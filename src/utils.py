from __future__ import annotations

import datetime as _dt
import os
import random
import re
from urllib.parse import urlparse, urlunparse


def now_iso() -> str:
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_url(url: str) -> str:
    url = safe_str(url)
    if not url:
        return ""
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    parsed = urlparse(url)
    if not parsed.netloc:
        return url
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path or "/", parsed.params, parsed.query, ""))


def domain_from_url(url: str) -> str:
    try:
        return urlparse(normalize_url(url)).netloc.replace("www.", "")
    except Exception:
        return safe_str(url)


def clamp_int(value, default: int, min_value: int, max_value: int) -> int:
    try:
        value = int(value)
    except Exception:
        return default
    return max(min_value, min(max_value, value))


def clamp_float(value, default: float, min_value: float, max_value: float) -> float:
    try:
        value = float(value)
    except Exception:
        return default
    return max(min_value, min(max_value, value))


def split_multi_value(value: str) -> list[str]:
    value = safe_str(value)
    if not value:
        return []
    parts = re.split(r"[\n;,|]+", value)
    return [p.strip() for p in parts if p.strip()]


def unique_keep_order(items: list[str]) -> list[str]:
    seen = set()
    out = []
    for item in items:
        key = item.strip().lower()
        if key and key not in seen:
            seen.add(key)
            out.append(item.strip())
    return out


def random_delay_bounds(min_delay: float, max_delay: float) -> float:
    if max_delay < min_delay:
        max_delay = min_delay
    return random.uniform(min_delay, max_delay)


def ensure_parent_dir(path: str) -> None:
    directory = os.path.dirname(os.path.abspath(path))
    if directory:
        os.makedirs(directory, exist_ok=True)
