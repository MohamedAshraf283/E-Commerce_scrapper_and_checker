from __future__ import annotations

import re
from urllib.parse import urljoin, urlparse

from .utils import domain_from_url, safe_str, unique_keep_order


FAILURE_KEYWORDS = [
    "تحت الصيانة", "قيد الصيانة", "تم إيقاف المتجر", "المتجر موقوف",
    "المتجر غير متاح", "غير متاح", "نعود قريب", "المتجر غير موجود",
    "store unavailable", "under maintenance", "temporarily disabled",
    "suspended", "not found", "server not found", "dns_probe_finished",
    "this store is unavailable", "coming soon", "maintenance",
]

HUMAN_VERIFICATION_KEYWORDS = [
    "verify you are human", "checking your browser", "are you human",
    "captcha", "recaptcha", "hcaptcha", "cloudflare", "attention required",
    "challenge", "human verification", "security check", "تحقق من أنك إنسان",
    "التحقق من أنك لست روبوت", "يرجى التحقق", "جاري التحقق", "كابتشا",
]

WORKING_SELECTORS = [
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
]

SOCIAL_HOSTS = [
    "instagram.com",
    "twitter.com",
    "x.com",
    "facebook.com",
    "snapchat.com",
    "tiktok.com",
    "youtube.com",
    "linkedin.com",
    "wa.me",
    "api.whatsapp.com",
    "whatsapp.com",
    "telegram.me",
    "t.me",
]

CATEGORY_HINTS = [
    "category",
    "categories",
    "collection",
    "collections",
    "taxon",
    "products",
    "product-category",
    "تصنيف",
    "منتجات",
    "الأقسام",
    "الاقسام",
]

EMAIL_RE = re.compile(r"(?i)\b[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}\b")
PHONE_RE = re.compile(r"(?:(?:\+|00)?966|0)?\s?5\d(?:[\s\-.]?\d){7}|(?:\+|00)?\d{8,15}")
PRICE_RE = re.compile(r"(?P<price>\d+(?:[,.]\d{1,2})?)\s*(?P<currency>ر\.س|SAR|SR|ريال|د\.إ|AED|USD|EGP)?", re.I)


def classify_navigation_error(exc: Exception) -> str:
    text = safe_str(exc).lower()
    if any(k in text for k in ["err_name_not_resolved", "dns_probe", "name_not_resolved", "getaddrinfo"]):
        return "مشكلة DNS أو الدومين غير موجود"
    if any(k in text for k in ["timeout", "timed out"]):
        return "انتهت مهلة التحميل"
    if any(k in text for k in ["ssl", "certificate", "cert"]):
        return "مشكلة SSL/Certificate"
    if any(k in text for k in ["connection refused", "connection reset", "net::err_connection"]):
        return "مشكلة اتصال بالخادم"
    return "خطأ أثناء فتح الصفحة"


def detect_platform(url: str, text: str = "", html: str = "") -> str:
    joined = " ".join([safe_str(url), safe_str(text)[:5000], safe_str(html)[:5000]]).lower()
    host = domain_from_url(url).lower()
    if "salla" in joined or host.endswith(".salla.sa") or "salla.sa" in host:
        return "Salla"
    if "zid" in joined or "zid.sa" in host or host.endswith(".zid.store"):
        return "Zid"
    return "غير معروف"


def has_failure_text(text: str, url: str = "") -> tuple[bool, str]:
    hay = f"{safe_str(text).lower()} {safe_str(url).lower()}"
    for keyword in FAILURE_KEYWORDS:
        if keyword.lower() in hay:
            return True, keyword
    if "/maintenance" in hay:
        return True, "maintenance"
    return False, ""


def has_human_verification_text(text: str, html: str = "", title: str = "") -> tuple[bool, str]:
    hay = " ".join([safe_str(text), safe_str(html), safe_str(title)]).lower()
    for keyword in HUMAN_VERIFICATION_KEYWORDS:
        if keyword.lower() in hay:
            return True, keyword
    return False, ""


def extract_contacts(text: str, links: list[dict]) -> dict[str, str]:
    text = safe_str(text)
    emails = unique_keep_order(EMAIL_RE.findall(text))

    phones = []
    for match in PHONE_RE.findall(text):
        cleaned = re.sub(r"[^\d+]", "", match)
        if len(cleaned) >= 8:
            phones.append(cleaned)
    phones = unique_keep_order(phones)

    socials = []
    whatsapp = []
    for item in links or []:
        href = safe_str(item.get("href"))
        if not href:
            continue
        low = href.lower()
        if any(host in low for host in SOCIAL_HOSTS):
            if "wa.me" in low or "whatsapp" in low:
                whatsapp.append(href)
            socials.append(href)

    address = ""
    address_match = re.search(r"(?:العنوان|address)\s*[:：\-]?\s*(.{10,160})", text, re.I)
    if address_match:
        address = " ".join(address_match.group(1).split())

    return {
        "emails": " | ".join(unique_keep_order(emails)),
        "phones": " | ".join(unique_keep_order(phones)),
        "whatsapp": " | ".join(unique_keep_order(whatsapp)),
        "socials": " | ".join(unique_keep_order(socials)),
        "address": address,
    }


def looks_like_category_url(href: str, text: str = "") -> bool:
    joined = f"{safe_str(href).lower()} {safe_str(text).lower()}"
    if not href:
        return False
    if any(x in joined for x in ["#", "javascript:", "mailto:", "tel:", "whatsapp"]):
        return False
    return any(hint.lower() in joined for hint in CATEGORY_HINTS)


def resolve_and_filter_category_links(base_url: str, links: list[dict], limit: int) -> list[str]:
    out = []
    base_domain = urlparse(base_url).netloc.lower().replace("www.", "")
    for item in links or []:
        href = safe_str(item.get("href"))
        text = safe_str(item.get("text"))
        if not looks_like_category_url(href, text):
            continue
        url = urljoin(base_url, href)
        domain = urlparse(url).netloc.lower().replace("www.", "")
        if domain and domain != base_domain:
            continue
        out.append(url.split("#")[0])
        if len(out) >= max(1, limit):
            break
    return unique_keep_order(out)[:limit]


def extract_price_and_currency(value: str) -> tuple[str, str]:
    value = safe_str(value)
    match = PRICE_RE.search(value)
    if not match:
        return value, ""
    return safe_str(match.group("price")).replace(",", "."), safe_str(match.group("currency"))
