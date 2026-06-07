from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ScrapeSettings:
    input_excel: str
    output_excel: str
    link_column: str = "link"

    browser_exe: str = ""
    user_data_dir: str = ""
    profile_dir: str = "Default"
    headless: bool = False

    modes: list[str] = field(default_factory=lambda: ["status", "contacts", "products"])

    concurrency: int = 2
    navigation_timeout_ms: int = 30000
    save_every: int = 5
    restart_every: int = 100

    min_delay_sec: float = 1.0
    max_delay_sec: float = 3.0
    human_verification_wait_sec: int = 60
    human_verification_retries: int = 1

    max_categories_per_store: int = 20
    max_category_pages: int = 3
    max_products_per_store: int = 200

    category_urls_column: str = ""
    selectors: dict[str, str] = field(default_factory=dict)

    sort_after_finish: bool = False
    sort_column: str = ""
    sort_descending: bool = False
    sort_numeric: bool = False


@dataclass
class InputRow:
    index: int
    excel_row: int
    data: dict[str, Any]

    def get(self, key: str, default: Any = "") -> Any:
        return self.data.get(key, default)


@dataclass
class ProductRow:
    store_index: int
    store_url: str
    category_url: str
    product_url: str = ""
    name: str = ""
    price: str = ""
    currency: str = ""
    image: str = ""
    availability: str = ""
    raw_text: str = ""


@dataclass
class StoreResult:
    index: int
    excel_row: int
    input_url: str
    final_url: str = ""
    platform: str = ""
    store_name: str = ""
    status: str = ""
    status_reason: str = ""
    http_or_error: str = ""
    emails: str = ""
    phones: str = ""
    whatsapp: str = ""
    socials: str = ""
    address: str = ""
    categories_found: int = 0
    products_found: int = 0
    started_at: str = ""
    finished_at: str = ""
    needs_retry_after_human_check: bool = False
    products: list[ProductRow] = field(default_factory=list)
