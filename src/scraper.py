from __future__ import annotations

import asyncio
import json
import os
import tempfile
from typing import Callable, Iterable
from urllib.parse import urljoin, urlparse

from playwright.async_api import TimeoutError as PlaywrightTimeoutError
from playwright.async_api import async_playwright

from .detectors import (
    detect_platform,
    extract_contacts,
    extract_price_and_currency,
    has_failure_text,
    has_human_verification_text,
    classify_navigation_error,
    resolve_and_filter_category_links,
)
from .excel_io import read_input_rows, write_output
from .models import InputRow, ProductRow, ScrapeSettings, StoreResult
from .utils import (
    domain_from_url,
    normalize_url,
    now_iso,
    random_delay_bounds,
    safe_str,
    split_multi_value,
    unique_keep_order,
)


Emit = Callable[[dict], None]


class PlaywrightScraper:
    def __init__(self, settings: ScrapeSettings, emit: Emit, stop_requested: Callable[[], bool] | None = None):
        self.settings = settings
        self.emit = emit
        self.stop_requested = stop_requested or (lambda: False)
        self._playwright = None
        self._context = None
        self._processed_since_restart = 0

    def log(self, message: str, level: str = "INFO") -> None:
        self.emit({"type": "log", "level": level, "message": message, "time": now_iso()})

    async def human_delay(self) -> None:
        await asyncio.sleep(random_delay_bounds(self.settings.min_delay_sec, self.settings.max_delay_sec))
        
    async def open_page_and_goto(self, url: str):
        if not self._context:
            raise RuntimeError("Browser context is not started")

        # مهم: لا نعيد استخدام about:blank مع التوازي
        # كل Task يأخذ تبويب مستقل حتى لا تتداخل الروابط
        page = await self._context.new_page()

        try:
            await page.bring_to_front()
        except Exception:
            pass

        self.log(f"فتح تبويب مستقل للرابط: {url}")

        response = None

        try:
            response = await page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=self.settings.navigation_timeout_ms
            )
        except PlaywrightTimeoutError:
            self.log(
                f"انتهت مهلة الفتح الأولي، سيتم فحص ما تم تحميله بدل إغلاق الصفحة مباشرة: {url}",
                "WARN"
            )
        except Exception as exc:
            self.log(f"فشل page.goto للرابط {url}: {exc}", "WARN")

        if page.url == "about:blank":
            self.log(f"التبويب ما زال about:blank؛ سيتم حقن الرابط مباشرة: {url}", "WARN")
            try:
                await page.evaluate(
                    "(targetUrl) => { window.location.href = targetUrl; }",
                    url
                )
                await page.wait_for_timeout(1500)
            except Exception as exc:
                self.log(f"فشل حقن الرابط مباشرة في التبويب: {exc}", "ERROR")

        self.log(f"الرابط الحالي داخل المتصفح: {page.url}")

        return page, response
    
    async def wait_page_ready(self, page, label: str = "") -> None:
        timeout_ms = max(3000, int(self.settings.navigation_timeout_ms))
        quick_timeout = min(timeout_ms, 6000)

        try:
            await page.wait_for_selector("body", state="attached", timeout=quick_timeout)
        except Exception:
            pass

        signal_selector = ",".join([
            "header",
            "footer",
            "salla-product-card",
            "salla-products-list",
            ".s-product",
            ".s-product-card-entry",
            ".s-product-card-image",
            ".s-cart-summary",
            ".s-cart-summary-wrapper",
            ".zid-product-card",
            ".product-card",
            ".product-item",
            ".product",
            "[data-product-id]",
            'input[type="search"]',
            ".s-search-input",
        ])

        loop = asyncio.get_running_loop()

        # حد أقصى سريع للانتظار بعد فتح الرابط
        deadline = loop.time() + min(max(timeout_ms / 1000, 4), 10)

        last_score = None
        stable_rounds = 0

        while loop.time() < deadline:
            if self.stop_requested():
                return

            try:
                info = await page.evaluate(
                    """(signalSelector) => {
                        const body = document.body;
                        const text = body ? (body.innerText || body.textContent || '') : '';
                        let signalCount = 0;

                        try {
                            signalCount = document.querySelectorAll(signalSelector).length;
                        } catch (e) {
                            signalCount = 0;
                        }

                        return {
                            ready: document.readyState || '',
                            textLength: text.trim().length,
                            scrollHeight: body ? body.scrollHeight : 0,
                            signalCount
                        };
                    }""",
                    signal_selector
                )
            except Exception:
                info = {}

            ready = info.get("ready", "")
            text_len = int(info.get("textLength") or 0)
            scroll_height = int(info.get("scrollHeight") or 0)
            signal_count = int(info.get("signalCount") or 0)

            score = text_len + scroll_height + signal_count * 10000

            if signal_count > 0 and text_len > 50:
                return

            if ready in ("interactive", "complete") and text_len > 300:
                return

            if score == last_score and text_len > 100:
                stable_rounds += 1
                if stable_rounds >= 2:
                    return
            else:
                stable_rounds = 0
                last_score = score

            await asyncio.sleep(0.5)

    async def start(self):
        self._playwright = await async_playwright().start()
        await self.launch_context()

    async def close(self):
        if self._context:
            try:
                await self._context.close()
            except Exception:
                pass
            self._context = None
        if self._playwright:
            try:
                await self._playwright.stop()
            except Exception:
                pass
            self._playwright = None

    async def launch_context(self):
        if self._context:
            try:
                await self._context.close()
            except Exception:
                pass

        user_data_dir = self.settings.user_data_dir.strip()
        if not user_data_dir:
            user_data_dir = os.path.join(tempfile.gettempdir(), "salla_zid_scraper_playwright_profile")
            os.makedirs(user_data_dir, exist_ok=True)

        args = ["--start-maximized"]
        if self.settings.profile_dir.strip():
            args.append(f"--profile-directory={self.settings.profile_dir.strip()}")

        kwargs = {
            "user_data_dir": user_data_dir,
            "headless": bool(self.settings.headless),
            "viewport": None,
            "args": args,
            "locale": "ar-SA",
            "timeout": self.settings.navigation_timeout_ms,
        }
        if self.settings.browser_exe.strip():
            kwargs["executable_path"] = self.settings.browser_exe.strip()

        self.log("فتح جلسة المتصفح...")
        self._context = await self._playwright.chromium.launch_persistent_context(**kwargs)
        self._context.set_default_timeout(self.settings.navigation_timeout_ms)
        self._context.set_default_navigation_timeout(self.settings.navigation_timeout_ms)
        self._processed_since_restart = 0

    async def maybe_restart_periodically(self):
        if self.settings.restart_every <= 0:
            return
        if self._processed_since_restart >= self.settings.restart_every:
            self.log("إعادة تشغيل دورية لجلسة المتصفح.")
            await self.launch_context()

    async def restart_after_human_verification(self):
        wait_sec = max(1, int(self.settings.human_verification_wait_sec))
        self.log(f"تم اكتشاف تحقق بشري. سيتم إغلاق الجلسة وفتحها بعد {wait_sec} ثانية بدون محاولة تجاوز التحقق.", "WARN")
        if self._context:
            try:
                await self._context.close()
            except Exception:
                pass
            self._context = None

        for second in range(wait_sec, 0, -1):
            if self.stop_requested():
                return
            self.emit({"type": "human_wait", "remaining": second, "time": now_iso()})
            await asyncio.sleep(1)

        await self.launch_context()

    async def run_from_excel(self) -> list[StoreResult]:
        headers, rows, link_col = read_input_rows(self.settings.input_excel, self.settings.link_column)
        self.settings.link_column = link_col

        total = len(rows)

        self.emit({"type": "headers", "headers": headers, "link_column": link_col})
        self.emit({"type": "progress", "done": 0, "total": total, "remaining": total})
        self.log(f"تم تحميل {total} رابط صالح من ملف Excel. عمود الروابط: {link_col}")

        results: list[StoreResult] = []

        await self.start()

        try:
            chunk_size = max(1, int(self.settings.concurrency))

            for start in range(0, total, chunk_size):
                if self.stop_requested():
                    self.log("تم طلب إيقاف التشغيل.", "WARN")
                    break

                # إعادة التشغيل تكون بين الدفعات فقط، وليس أثناء عمل صفحات متوازية
                await self.maybe_restart_periodically()

                chunk = rows[start:start + chunk_size]

                tasks = [
                    asyncio.create_task(self.scrape_with_retry(row, total))
                    for row in chunk
                ]

                for task in asyncio.as_completed(tasks):
                    if self.stop_requested():
                        break

                    try:
                        item = await task
                    except Exception as exc:
                        self.log(f"خطأ غير متوقع في Task: {exc}", "ERROR")
                        continue

                    if isinstance(item, StoreResult):
                        results.append(item)
                        self._processed_since_restart += 1

                        self.emit({"type": "result", "result": item})

                        completed = len(results)
                        self.emit({
                            "type": "progress",
                            "done": completed,
                            "total": total,
                            "remaining": max(total - completed, 0)
                        })

                        if completed and completed % max(1, self.settings.save_every) == 0:
                            ordered = sorted(results, key=lambda r: r.index)
                            write_output(self.settings.output_excel, ordered)
                            self.log(f"تم الحفظ المرحلي: {self.settings.output_excel}")

            ordered_results = sorted(results, key=lambda r: r.index)

            write_output(self.settings.output_excel, ordered_results)
            self.log(f"تم الحفظ النهائي: {self.settings.output_excel}")

            self.emit({
                "type": "done",
                "results": ordered_results,
                "output_excel": self.settings.output_excel
            })

            return ordered_results

        finally:
            await self.close()


    async def scrape_with_retry(self, row: InputRow, total: int) -> StoreResult:
        attempts = max(0, int(self.settings.human_verification_retries))

        result = await self.scrape_one(row, total)

        if result.needs_retry_after_human_check and self.settings.concurrency > 1:
            self.log(
                "تم رصد تحقق بشري، لكن لن يتم إغلاق وإعادة فتح الجلسة أثناء التوازي حتى لا يتم إغلاق تبويبات أخرى تعمل. "
                "لو تريد إعادة المحاولة بعد التحقق البشري اجعل التوازي = 1.",
                "WARN"
            )
            return result

        for _ in range(attempts):
            if self.stop_requested() or not result.needs_retry_after_human_check:
                break

            await self.restart_after_human_verification()
            result = await self.scrape_one(row, total)

        return result

    async def scrape_one(self, row: InputRow, total: int) -> StoreResult:
        input_url = normalize_url(row.get(self.settings.link_column))
        result = StoreResult(
            index=row.index,
            excel_row=row.excel_row,
            input_url=input_url,
            started_at=now_iso(),
            status="قيد الفحص",
            store_name=domain_from_url(input_url),
        )

        if not input_url:
            result.status = "خطأ"
            result.status_reason = "رابط فارغ"
            result.finished_at = now_iso()
            return result

        self.log(f"[{row.index + 1}/{total}] فتح: {input_url}")
        page = None
        response_status = ""
        try:
            page, response = await self.open_page_and_goto(input_url)

            if response is not None:
                response_status = str(response.status)

            await self.wait_page_ready(page, input_url)
            await self.human_delay()
            await self.scroll_to_bottom(page)
            await self.wait_page_ready(page, input_url)
            await self.human_delay()

            data = await self.get_basic_page_data(page)
            body_text = data.get("bodyText", "")
            html = data.get("html", "")
            title = data.get("title", "")
            final_url = page.url

            result.final_url = final_url
            result.http_or_error = response_status
            result.store_name = self.choose_store_name(title, final_url)
            result.platform = detect_platform(final_url, body_text, html)

            human, human_reason = has_human_verification_text(body_text, html, title)
            if human:
                result.status = "تحقق بشري"
                result.status_reason = f"تم رصد صفحة تحقق: {human_reason}"
                result.needs_retry_after_human_check = True
                return result

            failed, reason = has_failure_text(body_text, final_url)
            if failed or (response_status and response_status.startswith(("4", "5"))):
                result.status = "لا يعمل"
                result.status_reason = reason or f"HTTP {response_status}"
                return result

            working = await self.has_working_signals(page)
            if not working:
                result.status = "غير مؤكد"
                result.status_reason = "لم تظهر إشارات متجر واضحة"
            else:
                result.status = "يعمل"
                result.status_reason = "ظهرت إشارات متجر فعّال"

            if "contacts" in self.settings.modes:
                contacts = extract_contacts(data.get("headerFooterText", ""), data.get("headerFooterLinks", []))
                result.emails = contacts["emails"]
                result.phones = contacts["phones"]
                result.whatsapp = contacts["whatsapp"]
                result.socials = contacts["socials"]
                result.address = contacts["address"]

            categories = []
            if "products" in self.settings.modes:
                categories = await self.get_category_urls(page, row, data)
                result.categories_found = len(categories)
                products = await self.scrape_products_for_categories(result, categories)
                result.products = products
                result.products_found = len(products)
                if products:
                    result.status_reason = f"تم سحب {len(products)} منتج من {len(categories)} تصنيف/صفحة"
                else:
                    result.status_reason = result.status_reason + " - لم يتم العثور على منتجات بالـ selectors الحالية"

            return result

        except PlaywrightTimeoutError as exc:
            result.status = "خطأ"
            result.status_reason = "انتهت مهلة تحميل الصفحة"
            result.http_or_error = safe_str(exc)[:500]
            return result
        except Exception as exc:
            result.status = "خطأ"
            result.status_reason = classify_navigation_error(exc)
            result.http_or_error = safe_str(exc)[:500]
            return result
        finally:
            result.finished_at = now_iso()
            if page:
                try:
                    await page.close()
                except Exception:
                    pass

    async def scroll_to_bottom(self, page):
        try:
            last_height = await page.evaluate("() => document.body ? document.body.scrollHeight : 0")
            for _ in range(8):
                if self.stop_requested():
                    break
                await page.evaluate("() => window.scrollTo(0, document.body ? document.body.scrollHeight : 0)")
                await asyncio.sleep(0.5)
                new_height = await page.evaluate("() => document.body ? document.body.scrollHeight : 0")
                if new_height == last_height:
                    break
                last_height = new_height
        except Exception:
            pass

    async def has_working_signals(self, page) -> bool:
        for selector in [
            "footer",
            "header",
            "salla-product-card",
            "salla-products-list",
            ".s-product",
            ".s-product-card-entry",
            ".s-cart-summary",
            ".zid-product-card",
            ".product-card",
            ".product-item",
            'input[type="search"]',
            ".s-search-input",
        ]:
            try:
                if await page.locator(selector).count() > 0:
                    return True
            except Exception:
                continue
        return False

    async def get_basic_page_data(self, page) -> dict:
        return await page.evaluate(
            """() => {
                const clean = value => (value || '').replace(/\\s+/g, ' ').trim();
                const scoped = Array.from(document.querySelectorAll('header, footer'));
                const roots = scoped.length ? scoped : [document.body].filter(Boolean);
                const linksFrom = root => Array.from(root.querySelectorAll('a[href]')).map(a => ({
                    text: clean(a.innerText || a.textContent || a.title || ''),
                    href: a.href || a.getAttribute('href') || ''
                }));
                const headerFooterLinks = roots.flatMap(linksFrom).slice(0, 1000);
                const allLinks = Array.from(document.querySelectorAll('a[href]')).map(a => ({
                    text: clean(a.innerText || a.textContent || a.title || ''),
                    href: a.href || a.getAttribute('href') || ''
                })).slice(0, 2500);
                return {
                    title: document.title || '',
                    headerFooterText: roots.map(n => n.innerText || '').join('\\n').slice(0, 50000),
                    bodyText: (document.body ? document.body.innerText : '').slice(0, 120000),
                    html: (document.documentElement ? document.documentElement.innerHTML : '').slice(0, 120000),
                    headerFooterLinks,
                    allLinks
                };
            }"""
        )

    def choose_store_name(self, title: str, url: str) -> str:
        title = safe_str(title)
        if title and len(title) <= 100:
            for suffix in ["| سلة", "| zid", "| زد", "- Salla", "| Salla"]:
                title = title.replace(suffix, "")
            return title.strip(" -|")
        return domain_from_url(url)

    async def get_category_urls(self, page, row: InputRow, basic_data: dict) -> list[str]:
        from_input = []
        if self.settings.category_urls_column and self.settings.category_urls_column in row.data:
            from_input = [normalize_url(x) for x in split_multi_value(row.get(self.settings.category_urls_column))]

        selector_links = []
        selector = self.settings.selectors.get("category_links", "").strip()
        if selector:
            try:
                selector_links = await page.evaluate(
                    """(selector) => Array.from(document.querySelectorAll(selector)).map(a => ({
                        text: (a.innerText || a.textContent || '').trim(),
                        href: a.href || a.getAttribute('href') || ''
                    })).slice(0, 500)""",
                    selector,
                )
            except Exception as exc:
                self.log(f"تعذر استخدام selector التصنيفات: {exc}", "WARN")

        auto_links = resolve_and_filter_category_links(
            page.url,
            (selector_links or []) + basic_data.get("allLinks", []),
            self.settings.max_categories_per_store,
        )
        all_categories = unique_keep_order([*from_input, *auto_links])
        if not all_categories:
            all_categories = [page.url]
        return all_categories[: max(1, self.settings.max_categories_per_store)]

    async def scrape_products_for_categories(self, store_result: StoreResult, categories: list[str]) -> list[ProductRow]:
        all_products: list[ProductRow] = []
        seen_keys = set()

        for category_url in categories:
            if self.stop_requested() or len(all_products) >= self.settings.max_products_per_store:
                break

            next_url = normalize_url(category_url)
            page_count = 0
            while next_url and page_count < max(1, self.settings.max_category_pages):
                if self.stop_requested() or len(all_products) >= self.settings.max_products_per_store:
                    break

                page_count += 1
                page = None
                try:
                    page, _ = await self.open_page_and_goto(next_url)

                    await self.wait_page_ready(page, next_url)
                    await self.human_delay()
                    await self.scroll_to_bottom(page)
                    await self.wait_page_ready(page, next_url)

                    products = await self.extract_products_from_current_page(page, store_result, category_url)
                    for product in products:
                        key = (product.product_url or product.name or product.raw_text).lower()
                        if key and key not in seen_keys:
                            seen_keys.add(key)
                            all_products.append(product)
                            if len(all_products) >= self.settings.max_products_per_store:
                                break

                    next_url = await self.find_next_page_url(page, next_url)
                except Exception as exc:
                    self.log(f"تعذر سحب منتجات من {next_url}: {exc}", "WARN")
                    next_url = ""
                finally:
                    if page:
                        try:
                            await page.close()
                        except Exception:
                            pass

        return all_products

    async def extract_products_from_current_page(self, page, store_result: StoreResult, category_url: str) -> list[ProductRow]:
        selectors = self.settings.selectors or {}
        raw_products = await page.evaluate(
            """(selectors) => {
                const clean = value => (value || '').replace(/\\s+/g, ' ').trim();
                const first = (root, selectorList) => {
                    for (const selector of selectorList.filter(Boolean)) {
                        try {
                            const node = root.querySelector(selector);
                            if (node) return node;
                        } catch (e) {}
                    }
                    return null;
                };
                const attrOrText = (node, attr) => {
                    if (!node) return '';
                    if (attr && node.getAttribute(attr)) return clean(node.getAttribute(attr));
                    if (node.href) return node.href;
                    if (node.src) return node.src;
                    return clean(node.innerText || node.textContent || '');
                };

                const cardSelector = selectors.product_card || [
                    'salla-product-card',
                    '.s-product-card-entry',
                    '.s-product',
                    '.zid-product-card',
                    '.product-card',
                    '.product-item',
                    '.product',
                    '[data-product-id]'
                ].join(',');
                let cards = [];
                try { cards = Array.from(document.querySelectorAll(cardSelector)); } catch (e) { cards = []; }

                const products = [];
                for (const card of cards.slice(0, 500)) {
                    const nameNode = selectors.product_name
                        ? first(card, [selectors.product_name])
                        : first(card, ['[itemprop="name"]', '.s-product-card-content-title', '.s-product-card__title', '.product-title', '.title', 'h1', 'h2', 'h3', 'a[title]']);
                    const priceNode = selectors.product_price
                        ? first(card, [selectors.product_price])
                        : first(card, ['[itemprop="price"]', '.s-product-card-price', '.price', '.product-price', '.amount', '[class*="price"]']);
                    const linkNode = selectors.product_url
                        ? first(card, [selectors.product_url])
                        : first(card, ['a[href]']);
                    const imgNode = selectors.product_image
                        ? first(card, [selectors.product_image])
                        : first(card, ['img[src]', 'img[data-src]', 'source[srcset]']);

                    let image = '';
                    if (imgNode) {
                        image = imgNode.src || imgNode.getAttribute('data-src') || imgNode.getAttribute('srcset') || imgNode.getAttribute('src') || '';
                    }
                    let productUrl = '';
                    if (linkNode) {
                        productUrl = linkNode.href || linkNode.getAttribute('href') || '';
                    }

                    products.push({
                        name: attrOrText(nameNode),
                        priceText: attrOrText(priceNode),
                        product_url: productUrl,
                        image,
                        raw_text: clean(card.innerText || card.textContent || '').slice(0, 1000)
                    });
                }

                if (products.length === 0) {
                    for (const script of Array.from(document.querySelectorAll('script[type="application/ld+json"]')).slice(0, 50)) {
                        try {
                            const parsed = JSON.parse(script.textContent || '{}');
                            const items = Array.isArray(parsed) ? parsed : [parsed];
                            for (const item of items) {
                                const graph = item['@graph'] || [item];
                                for (const node of graph) {
                                    const type = node['@type'];
                                    const isProduct = type === 'Product' || (Array.isArray(type) && type.includes('Product'));
                                    if (!isProduct) continue;
                                    products.push({
                                        name: clean(node.name || ''),
                                        priceText: clean((node.offers && (node.offers.price || node.offers.lowPrice)) || ''),
                                        product_url: node.url || '',
                                        image: Array.isArray(node.image) ? (node.image[0] || '') : (node.image || ''),
                                        raw_text: clean(JSON.stringify(node)).slice(0, 1000)
                                    });
                                }
                            }
                        } catch (e) {}
                    }
                }
                return products;
            }""",
            selectors,
        )

        out: list[ProductRow] = []
        for item in raw_products or []:
            price, currency = extract_price_and_currency(item.get("priceText", ""))
            product_url = safe_str(item.get("product_url"))
            if product_url:
                product_url = urljoin(page.url, product_url)
            raw_text = safe_str(item.get("raw_text"))
            availability = ""
            low_raw = raw_text.lower()
            if any(k in low_raw for k in ["نفدت", "غير متوفر", "out of stock", "sold out"]):
                availability = "غير متوفر"
            elif raw_text:
                availability = "متوفر/غير مؤكد"

            name = safe_str(item.get("name"))
            if not name and raw_text:
                name = raw_text[:120]

            if name or product_url:
                out.append(ProductRow(
                    store_index=store_result.index,
                    store_url=store_result.final_url or store_result.input_url,
                    category_url=category_url,
                    product_url=product_url,
                    name=name,
                    price=price,
                    currency=currency,
                    image=safe_str(item.get("image")),
                    availability=availability,
                    raw_text=raw_text[:500],
                ))
        return out

    async def find_next_page_url(self, page, current_url: str) -> str:
        try:
            href = await page.evaluate(
                """() => {
                    const clean = v => (v || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const candidates = Array.from(document.querySelectorAll('a[href]'));
                    const next = candidates.find(a => {
                        const rel = (a.getAttribute('rel') || '').toLowerCase();
                        const aria = clean(a.getAttribute('aria-label'));
                        const text = clean(a.innerText || a.textContent);
                        return rel === 'next' || aria.includes('next') || aria.includes('التالي') ||
                               text === 'next' || text === '>' || text.includes('التالي');
                    });
                    return next ? (next.href || next.getAttribute('href') || '') : '';
                }"""
            )
            if not href:
                return ""
            href = urljoin(current_url, href)
            current_host = urlparse(current_url).netloc
            if urlparse(href).netloc != current_host:
                return ""
            return href.split("#")[0]
        except Exception:
            return ""


async def run_scraper(settings: ScrapeSettings, emit: Emit, stop_requested: Callable[[], bool] | None = None) -> list[StoreResult]:
    scraper = PlaywrightScraper(settings, emit, stop_requested)
    return await scraper.run_from_excel()


def parse_selectors_json(text: str) -> dict[str, str]:
    text = safe_str(text)
    if not text:
        return {}
    parsed = json.loads(text)
    if not isinstance(parsed, dict):
        raise ValueError("selectors يجب أن تكون JSON object")
    return {str(k): str(v) for k, v in parsed.items() if str(k).strip() and str(v).strip()}
