"""Tracked new-product introductions rolled up by sales channel."""
from __future__ import annotations

import re
from decimal import Decimal
from typing import Iterable

from reporting.models import (
    PLATFORM_KEYS,
    PLATFORM_LABELS,
    PLATFORM_SHORT,
    NewProductChannelRow,
    SkuRow,
)

# Canonical SKUs (Shopify / Sellerboard may prefix FBA / WM- and add size suffixes).
TRACKED_NEW_PRODUCTS: tuple[dict[str, object], ...] = (
    {
        "key": "travel_skyline_stripe",
        "label": "Travel Umbrella – Skyline Stripe",
        "sku_tokens": ("10001-000",),
        "title_all": ("skyline stripe",),
        "title_require_any": (),
    },
    {
        "key": "travel_dusty_lavender",
        "label": "Travel Umbrella – Dusty Lavender",
        "sku_tokens": ("10001-511",),
        "title_all": ("dusty lavender",),
        # Prefer Travel family; SKU match alone is authoritative.
        "title_require_any": ("travel", "compact mini"),
    },
    {
        "key": "trek_rusty_orange",
        "label": "Trek Umbrella – Rusty Orange",
        "sku_tokens": ("12005-801",),
        "title_all": ("rusty orange",),
        # Avoid Walk / Kids / Travel Rusty Orange when matching by title only.
        "title_require_any": ("trek", "collapsible"),
    },
)

_PLATFORM_LABEL_TO_KEY = {
    **{label.lower(): key for key, label in PLATFORM_LABELS.items()},
    **{short.lower(): key for key, short in PLATFORM_SHORT.items()},
    "shopify": "shopify_direct",
    "shopify direct": "shopify_direct",
    "dick's sporting goods": "dsg",
    "dicks sporting goods": "dsg",
}


def normalize_sku(sku: str) -> str:
    """Strip FBA/WM prefixes and non-alphanumerics for fuzzy SKU matching."""
    text = str(sku or "").upper()
    text = re.sub(r"^(?:FBA\d*[\s_\-]*)+", "", text)
    text = re.sub(r"^WM[\s_\-]*", "", text)
    return re.sub(r"[^A-Z0-9]", "", text)


def platform_key_from_label(label: str) -> str | None:
    return _PLATFORM_LABEL_TO_KEY.get(str(label or "").strip().lower())


def _sku_matches(sku: str, tokens: Iterable[str]) -> bool:
    norm = normalize_sku(sku)
    if not norm:
        return False
    for token in tokens:
        token_norm = normalize_sku(str(token))
        if token_norm and token_norm in norm:
            return True
    return False


def _title_matches(
    title: str,
    *,
    title_all: Iterable[str],
    title_require_any: Iterable[str],
) -> bool:
    text = str(title or "").lower()
    required = tuple(title_all)
    if not required or not all(part.lower() in text for part in required):
        return False
    any_needed = tuple(title_require_any)
    if not any_needed:
        return True
    return any(part.lower() in text for part in any_needed)


def match_tracked_product_key(sku: str, title: str) -> str | None:
    """Return tracked product key if SKU/title matches a new introduction."""
    for product in TRACKED_NEW_PRODUCTS:
        tokens = product["sku_tokens"]  # type: ignore[assignment]
        if _sku_matches(sku, tokens):  # type: ignore[arg-type]
            return str(product["key"])
    for product in TRACKED_NEW_PRODUCTS:
        if _title_matches(
            title,
            title_all=product["title_all"],  # type: ignore[arg-type]
            title_require_any=product["title_require_any"],  # type: ignore[arg-type]
        ):
            return str(product["key"])
    return None


def empty_new_product_rows() -> list[NewProductChannelRow]:
    return [
        NewProductChannelRow(
            key=str(p["key"]),
            label=str(p["label"]),
            units_by_platform={k: 0 for k in PLATFORM_KEYS},
            revenue_by_platform={k: Decimal("0") for k in PLATFORM_KEYS},
        )
        for p in TRACKED_NEW_PRODUCTS
    ]


def build_new_product_rows(sku_rows: list[SkuRow]) -> list[NewProductChannelRow]:
    """Roll SKU drilldown lines into the fixed new-product × channel matrix."""
    rows = {row.key: row for row in empty_new_product_rows()}
    for sku_row in sku_rows:
        product_key = match_tracked_product_key(sku_row.sku, sku_row.item)
        if not product_key or product_key not in rows:
            continue
        platform_key = platform_key_from_label(sku_row.platform)
        if not platform_key:
            continue
        target = rows[product_key]
        target.units_by_platform[platform_key] = int(
            target.units_by_platform.get(platform_key, 0) or 0
        ) + int(sku_row.units or 0)
        prior = Decimal(str(target.revenue_by_platform.get(platform_key, Decimal("0"))))
        target.revenue_by_platform[platform_key] = prior + Decimal(str(sku_row.revenue or 0))
    return [rows[str(p["key"])] for p in TRACKED_NEW_PRODUCTS]
