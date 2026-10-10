#!/usr/bin/env python3
"""Populate the Weatherman Monthly Business Review with live data.

Sources
-------
A. ``data/daily_archive.json`` — Shopify, Amazon, Walmart, Nordstrom, DSG
B. Sunny's Google Sheet *Weatherman IP Tracker US*
   (``1e60-xD1OYqpjpqHFjx_Mw6nnDYduQLy9sSCaVaOXZlg``):
   - ``Monthly Sales for PPT`` (gid 1734803925) — SKU velocity / MoM
   - ``IP TRACKER - AMZ`` (gid 279947413) — stock / OOS risk
C. Optional ``data/monthly_ad_metrics.json`` — ROAS / ACOS / TACOS / CAC / CVR
   when ad-account feeds are available (archive alone has no ad spend).
D. Triple Whale API (``TRIPLE_WHALE_API_KEY``) or ``data/triple_whale_export.json``
   — PTP vs Actuals (slide 14) + Channel Mix (slide 15) + Site Health sessions.
   YoY cards (slides 4/6/9) use TW Summary for the **same MTD window last year**
   because ``daily_archive`` has no 2025 days. Optional planned PTP:
   ``data/ptp_targets.json`` (else prior-month TW).
E. Chart PNGs from ``scripts/render_monthly_deck_charts.py`` (line-bar YTD,
   blue retailers table, OOS cards) — insert via MCP/API, not ASCII overlays.
   MER = ads÷revenue %. Amazon CVR/RPR stay blank when absent from Sellerboard.

Outputs
-------
1. Data-complete ``.pptx`` (python-pptx) for the monthly packet.
2. Live Google Slides updates against the creative deck (financial cards,
   inventory narratives/tables, slides 14–15 attribution). Slides **25**, **27**,
   and **32+** are never modified.

Examples
--------
.. code-block:: bash

    .venv/bin/python scripts/update_monthly_deck.py --month 2026-10 \\
        --presentation-id 1Z5rlR-udC7_YHhbIVwuIQChI5nEz_72RMd5ODkSr5bA \\
        --apply-slides --google-credentials creds.json

    # Compute + write plan/PPTX only (apply plan via MCP / API later)
    .venv/bin/python scripts/update_monthly_deck.py --month 2026-10 \\
        --presentation-id 1Z5rlR-udC7_YHhbIVwuIQChI5nEz_72RMd5ODkSr5bA
"""
from __future__ import annotations

import argparse
import calendar
import csv
import json
import logging
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from urllib.request import urlopen

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from monthly_deck_sources import (  # noqa: E402
    MonthChannel as LiveMonthChannel,
    TripleWhaleMonth,
    executive_amazon_copy,
    fetch_triple_whale_month,
    merge_live_bundle,
    money_compact_k,
    money_whole as live_money_whole,
)
DEFAULT_ARCHIVE = ROOT / "data" / "daily_archive.json"
DEFAULT_OUTPUT_DIR = ROOT / "output" / "monthly_decks"
DEFAULT_IP_TRACKER_ID = "1e60-xD1OYqpjpqHFjx_Mw6nnDYduQLy9sSCaVaOXZlg"
DEFAULT_SLIDES_TEMPLATE_ID = "1RZT2mrFAdcxcrJ8jNMPFYZpSjLZ7beSoOkNg-JnMS08"
DEFAULT_OCTOBER_DECK_ID = "1Z5rlR-udC7_YHhbIVwuIQChI5nEz_72RMd5ODkSr5bA"
DEFAULT_AD_METRICS = ROOT / "data" / "monthly_ad_metrics.json"
DEFAULT_TW_EXPORT = ROOT / "data" / "triple_whale_export.json"
DEFAULT_PTP_TARGETS = ROOT / "data" / "ptp_targets.json"
MONTHLY_SALES_GID = "1734803925"
IP_TRACKER_AMZ_GID = "279947413"

# 1-based slide numbers the team owns — never touch.
# Per latest brief: leave 25, 27, and 32+ alone (26 may be updated when wired).
PROTECTED_SLIDES = frozenset({25, 27})
PROTECTED_SLIDE_MIN = 32  # skip slide 32 and everything after

# Creative October deck page object IDs (1-based slide → page id).
SLIDE_PAGE_IDS: dict[int, str] = {
    1: "p1",
    2: "g3c71e105407_1_67",
    3: "h6e6e0e983a976184_12_578",
    4: "h1872ece28e3fae88_75_188",
    5: "h6e6e0e983a976184_12_159",
    6: "h1872ece28e3fae88_75_0",
    7: "g409c5cc28db_0_3035",
    8: "h1872ece28e3fae88_76_198",
    9: "h1872ece28e3fae88_76_274",
    10: "h1872ece28e3fae88_13_421",
    11: "h3a7aee58bd343859_3_1199",
    12: "g3f8cdd12cc6_3_1291",
    13: "g409c5cc28db_0_3415",
    14: "g3f366a59143_0_624",
    15: "g3fa8ada2d55_123_161",
    16: "g3f8cdd12cc6_132_1728",
    17: "g3f400dce2ff_0_42",
    18: "h431b9d78e3d0bc19_45_137",
    19: "h431b9d78e3d0bc19_7_16",
    20: "h431b9d78e3d0bc19_7_31",
    21: "h431b9d78e3d0bc19_7_3",
    22: "h431b9d78e3d0bc19_7_23",
    23: "h431b9d78e3d0bc19_7_38",
    24: "h431b9d78e3d0bc19_110_0",
    32: "g3fa8ada2d55_81_20",
}

PLATFORM_ORDER = (
    "amazon",
    "shopify_direct",
    "walmart",
    "nordstrom",
    "dsg",
)
PLATFORM_LABELS = {
    "amazon": "Amazon",
    "shopify_direct": "Shopify Direct",
    "walmart": "Walmart",
    "nordstrom": "Nordstrom",
    "dsg": "DSG",
}

NAVY = RGBColor(0x0B, 0x19, 0x2C)
TEAL = RGBColor(0x0F, 0x76, 0x6E)
GRAY = RGBColor(0x4B, 0x55, 0x63)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF3, 0xF5, 0xF7)

log = logging.getLogger("monthly_deck")
MONEY = Decimal("0.01")
EMU_PER_INCH = 914400


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def money(value: Decimal | float | int | None) -> str:
    amount = Decimal(str(value or 0)).quantize(MONEY, rounding=ROUND_HALF_UP)
    return f"${amount:,.2f}"


def money_whole(value: Decimal | float | int | None) -> str:
    """Whole-dollar card format used by the creative deck metric tiles."""
    amount = Decimal(str(value or 0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return f"${amount:,.0f}"


def money_compact(value: Decimal | float | int | None) -> str:
    amount = float(value or 0)
    if abs(amount) >= 1000:
        return f"${amount:,.1f}"
    return f"${amount:,.2f}"


def num(value: Any) -> float:
    if value is None:
        return 0.0
    text = str(value).strip().replace(",", "").replace("$", "").replace("%", "")
    if not text or text in {"-", "—", "N/A", "n/a"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_month(value: str | None) -> date:
    if not value:
        today = date.today()
        if today.month == 1:
            return date(today.year - 1, 12, 1)
        return date(today.year, today.month - 1, 1)
    text = value.strip()
    if re.fullmatch(r"\d{4}-\d{2}", text):
        year, month = map(int, text.split("-"))
        return date(year, month, 1)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return date.fromisoformat(text).replace(day=1)
    raise SystemExit(f"Invalid --month {value!r}; expected YYYY-MM")


def month_label(d: date) -> str:
    return f"{calendar.month_name[d.month]} {d.year}"


def prior_month(d: date) -> date:
    if d.month == 1:
        return date(d.year - 1, 12, 1)
    return date(d.year, d.month - 1, 1)


def month_end(d: date) -> date:
    last = calendar.monthrange(d.year, d.month)[1]
    return date(d.year, d.month, last)


def period_label(start: date, end: date) -> str:
    return f"{start.month}/{start.day:02d}/{start.year}–{end.month}/{end.day:02d}/{end.year}"


def sheet_csv_url(spreadsheet_id: str, gid: str) -> str:
    return (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
        f"?format=csv&gid={gid}"
    )


def download_csv(url: str, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    log.info("Downloading sheet CSV → %s", dest)
    with urlopen(url, timeout=60) as resp:  # noqa: S310 — fixed Google Sheets URL
        dest.write_bytes(resp.read())
    return dest


def pct_change(cur: float | Decimal, prior: float | Decimal) -> str | None:
    cur_f, prior_f = float(cur), float(prior)
    if prior_f == 0:
        return None
    delta = (cur_f - prior_f) / prior_f * 100
    arrow = "▲" if delta >= 0 else "▼"
    return f"{arrow} {abs(delta):.2f}%"


def pts_change(cur: float | None, prior: float | None) -> str | None:
    if cur is None or prior is None:
        return None
    delta = cur - prior
    arrow = "▲" if delta >= 0 else "▼"
    return f"{arrow} {abs(delta):.2f} pts"


def aov(revenue: Decimal | float, orders: int) -> Decimal | None:
    if not orders:
        return None
    return Decimal(str(revenue)) / Decimal(orders)


def fmt_aov(value: Decimal | None) -> str:
    return money(value) if value is not None else "n/a"


def fmt_metric(value: Any, *, kind: str = "number") -> str:
    if value is None or value == "":
        return "n/a"
    if kind == "money":
        return money(value)
    if kind == "pct":
        return f"{float(value):.2f}%"
    if kind == "roas":
        return f"{float(value):.2f}"
    return str(value)


def table_text(headers: list[str], rows: list[list[str]], *, max_rows: int = 12) -> str:
    use = rows[:max_rows]
    widths = [len(h) for h in headers]
    for row in use:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(str(cell)))

    def fmt_row(vals: list[str]) -> str:
        return " | ".join(str(v).ljust(widths[i]) for i, v in enumerate(vals))

    lines = [fmt_row(headers), "-+-".join("-" * w for w in widths)]
    lines.extend(fmt_row(r) for r in use)
    if not use:
        lines.append("(no rows)")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------


@dataclass
class ChannelMonth:
    revenue: Decimal = Decimal("0")
    units: int = 0
    orders: int = 0
    days_available: int = 0

    def add_day(self, revenue: Any, units: Any, orders: Any, available: bool) -> None:
        if not available:
            return
        self.revenue += Decimal(str(revenue or 0))
        self.units += int(units or 0)
        self.orders += int(orders or 0)
        self.days_available += 1

    @property
    def aov(self) -> Decimal | None:
        return aov(self.revenue, self.orders)


@dataclass
class ArchiveMonthMetrics:
    month: date
    channels: dict[str, ChannelMonth] = field(default_factory=dict)
    day_count: int = 0
    first_day: date | None = None
    last_day: date | None = None

    @property
    def total_revenue(self) -> Decimal:
        return sum((c.revenue for c in self.channels.values()), Decimal("0"))

    @property
    def total_units(self) -> int:
        return sum(c.units for c in self.channels.values())

    @property
    def total_orders(self) -> int:
        return sum(c.orders for c in self.channels.values())


@dataclass
class AdChannelMetrics:
    roas: float | None = None
    acos: float | None = None
    tacos: float | None = None
    cac: float | None = None
    cvr: float | None = None
    repeat_rate: float | None = None
    net_contribution: float | None = None


@dataclass
class SkuSalesRow:
    sku: str
    style: str
    color: str
    status: str
    category: str
    month_units: float
    prior_units: float
    mom_pct: float | None
    ytd_units: float
    velocity_4wk: float = 0.0


@dataclass
class InventoryRow:
    sku: str
    product: str
    style: str
    color: str
    status: str
    available: float
    reserved: float
    total_stock: float
    days_of_cover: float | None = None


@dataclass
class DeckData:
    month: date
    archive: ArchiveMonthMetrics
    prior_archive: ArchiveMonthMetrics
    sheet_month_label: str
    top_sellers: list[SkuSalesRow]
    struggling: list[SkuSalesRow]
    new_launches: list[SkuSalesRow]
    inventory_summary: dict[str, float]
    low_stock: list[InventoryRow]
    oos_projections: list[InventoryRow]
    b2b_top: list[SkuSalesRow]
    velocity_per_day: float
    ad_metrics: dict[str, AdChannelMetrics]
    color_launches: list[tuple[str, list[SkuSalesRow]]]
    ytd_by_month: dict[str, dict[str, float]] = field(default_factory=dict)
    triple_whale: TripleWhaleMonth | None = None


@dataclass
class TextUpdate:
    object_id: str
    text: str
    slide: int


@dataclass
class DeleteOp:
    object_id: str
    slide: int


@dataclass
class CreateTextBoxOp:
    slide: int
    page_object_id: str
    text: str
    x: int
    y: int
    width: int
    height: int
    font_size: int = 11
    bold: bool = False


@dataclass
class ReplaceAllOp:
    """Page-scoped replaceAllText — required for grouped metric-card shapes.

    Creative-deck cards on slides 3–10 live inside ``elementGroup`` containers.
    Top-level ``setElementText`` / title bindings never reach those children, so
    we also emit exact-string replacements scoped to the financial slides.
    """

    contains: str
    replace: str
    slides: list[int] = field(default_factory=list)  # 1-based; empty = all non-protected
    match_case: bool = True


@dataclass
class SlidesPlan:
    presentation_id: str
    month: str
    text_updates: list[TextUpdate] = field(default_factory=list)
    deletes: list[DeleteOp] = field(default_factory=list)
    create_text_boxes: list[CreateTextBoxOp] = field(default_factory=list)
    replace_all: list[ReplaceAllOp] = field(default_factory=list)
    protected_slides: list[int] = field(default_factory=lambda: sorted(PROTECTED_SLIDES))
    notes: list[str] = field(default_factory=list)

    def to_json(self) -> dict[str, Any]:
        return {
            "presentation_id": self.presentation_id,
            "month": self.month,
            "protected_slides": self.protected_slides,
            "notes": self.notes,
            "text_updates": [asdict(x) for x in self.text_updates],
            "deletes": [asdict(x) for x in self.deletes],
            "create_text_boxes": [asdict(x) for x in self.create_text_boxes],
            "replace_all": [asdict(x) for x in self.replace_all],
        }


# ---------------------------------------------------------------------------
# Source A — daily archive
# ---------------------------------------------------------------------------


def load_archive_month(archive_path: Path, month: date) -> ArchiveMonthMetrics:
    payload = json.loads(archive_path.read_text(encoding="utf-8"))
    days = payload.get("days") or {}
    prefix = month.strftime("%Y-%m")
    metrics = ArchiveMonthMetrics(
        month=month,
        channels={key: ChannelMonth() for key in PLATFORM_ORDER},
    )
    for day_key in sorted(days):
        if not str(day_key).startswith(prefix):
            continue
        day = days[day_key] or {}
        metrics.day_count += 1
        d = date.fromisoformat(str(day_key))
        metrics.first_day = d if metrics.first_day is None else min(metrics.first_day, d)
        metrics.last_day = d if metrics.last_day is None else max(metrics.last_day, d)
        platforms = day.get("platforms") or {}
        for key in PLATFORM_ORDER:
            row = platforms.get(key) or {}
            metrics.channels[key].add_day(
                row.get("revenue"),
                row.get("units"),
                row.get("orders"),
                bool(row.get("available")),
            )
    return metrics


def load_ad_metrics(path: Path, month: date) -> dict[str, AdChannelMetrics]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    block = payload.get(month.strftime("%Y-%m")) or payload.get(month_label(month)) or {}
    out: dict[str, AdChannelMetrics] = {}
    for key, raw in (block or {}).items():
        if not isinstance(raw, dict):
            continue
        out[key] = AdChannelMetrics(
            roas=raw.get("roas"),
            acos=raw.get("acos"),
            tacos=raw.get("tacos"),
            cac=raw.get("cac"),
            cvr=raw.get("cvr"),
            repeat_rate=raw.get("repeat_rate"),
            net_contribution=raw.get("net_contribution"),
        )
    return out


# ---------------------------------------------------------------------------
# Source B — IP Tracker sheet tabs
# ---------------------------------------------------------------------------


def load_monthly_sales_csv(path: Path, month: date) -> tuple[list[SkuSalesRow], str, float]:
    """Parse Sunny's ``Monthly Sales for PPT`` export.

    Returns rows, the sheet month column actually used, and overall 4-week
    velocity/day when present in the header band.
    """
    rows = list(csv.reader(path.open(encoding="utf-8", newline="")))
    header_idx = None
    for i, row in enumerate(rows[:20]):
        joined = ",".join(row).lower()
        if "asin" in joined and ("jan" in joined or "sep" in joined):
            header_idx = i
            break
    if header_idx is None:
        log.warning("Monthly Sales for PPT: header not found in %s", path)
        return [], month_label(month), 0.0

    header = [h.strip() for h in rows[header_idx]]
    col: dict[str, int] = {}
    for idx, name in enumerate(header):
        key = name.strip()
        if key and key not in col:
            col[key] = idx

    def _find_month_col(abbr: str) -> int | None:
        for candidate in (abbr, abbr.title()):
            if candidate in col:
                return col[candidate]
        for name, idx in col.items():
            if name.lower().startswith(abbr.lower()):
                return idx
        return None

    cur_key = calendar.month_abbr[month.month]
    prior_key = calendar.month_abbr[prior_month(month).month]
    cur_idx = _find_month_col(cur_key)
    sheet_month_used = calendar.month_name[month.month]
    if cur_idx is None:
        # Early-month MBR: sheet may not have the current month column yet.
        for m in range(month.month - 1, 0, -1):
            abbr = calendar.month_abbr[m]
            idx = _find_month_col(abbr)
            if idx is not None:
                cur_idx = idx
                sheet_month_used = calendar.month_name[m]
                prior_key = calendar.month_abbr[m - 1] if m > 1 else "Dec"
                log.warning(
                    "Month column %s missing; using %s units from sheet",
                    cur_key,
                    sheet_month_used,
                )
                break
    prior_idx = _find_month_col(prior_key)
    if prior_idx is None and cur_idx is not None:
        prior_idx = cur_idx - 1 if cur_idx > 0 else None

    velocity = 0.0
    # Month columns sit under Jan..Dec headers — use those indexes only.
    month_abbrs = [calendar.month_abbr[m] for m in range(1, 13)]
    month_col_idxs = [col[a] for a in month_abbrs if a in col]
    for band in rows[:header_idx]:
        joined = ",".join(band).lower()
        if "overall sale velocity" not in joined and "velocity / day" not in joined:
            continue
        vals = []
        for idx in month_col_idxs:
            if idx < len(band):
                v = num(band[idx])
                if v > 0:
                    vals.append(v)
        if vals:
            window = vals[-4:]
            velocity = sum(window) / len(window)
        break

    sku_idx = col.get("Amazon", col.get("SKU", 1))
    style_idx = col.get("Style", 3)
    color_idx = col.get("Color", 4)
    status_idx = col.get("Status", 7)
    category_idx = col.get("Category", 8)
    mom_idx = col.get("MOM %")
    total_idx = col.get("Total Sales")

    out: list[SkuSalesRow] = []
    for row in rows[header_idx + 1 :]:
        if not row or len(row) <= max(sku_idx, style_idx, color_idx):
            continue
        sku = (row[sku_idx] if sku_idx < len(row) else "").strip()
        if not sku or sku.lower() in {"amazon", "sku"}:
            continue
        month_units = num(row[cur_idx]) if cur_idx is not None and cur_idx < len(row) else 0.0
        prior_units = num(row[prior_idx]) if prior_idx is not None and prior_idx < len(row) else 0.0
        mom_raw = row[mom_idx] if mom_idx is not None and mom_idx < len(row) else ""
        mom_pct = num(mom_raw) if str(mom_raw).strip() else None
        ytd = num(row[total_idx]) if total_idx is not None and total_idx < len(row) else 0.0
        # Proxy 4-week velocity: month units / ~30 * 7 ≈ weekly; use month/4.
        vel = month_units / 4.0 if month_units else 0.0
        out.append(
            SkuSalesRow(
                sku=sku,
                style=(row[style_idx] if style_idx < len(row) else "").strip(),
                color=(row[color_idx] if color_idx < len(row) else "").strip(),
                status=(row[status_idx] if status_idx < len(row) else "").strip(),
                category=(row[category_idx] if category_idx < len(row) else "").strip(),
                month_units=month_units,
                prior_units=prior_units,
                mom_pct=mom_pct,
                ytd_units=ytd,
                velocity_4wk=vel,
            )
        )
    log.info(
        "Monthly Sales for PPT: %s SKU rows (sheet month=%s, velocity/day≈%.0f)",
        len(out),
        sheet_month_used,
        velocity,
    )
    return out, sheet_month_used, velocity


def load_inventory_csv(path: Path) -> list[InventoryRow]:
    sample = path.read_text(encoding="utf-8", errors="ignore")[:500]
    if "TOTAL STOCKS" in sample and "FBA-SKU" in sample and sample.lstrip().startswith("Sr"):
        rows = list(csv.DictReader(path.open(encoding="utf-8", newline="")))
        out: list[InventoryRow] = []
        for row in rows:
            out.append(
                InventoryRow(
                    sku=str(row.get("FBA-SKU") or row.get("3PL-SKU") or "").strip(),
                    product=str(row.get("Product") or "").strip(),
                    style=str(row.get("Style") or "").strip(),
                    color=str(row.get("Color") or "").strip(),
                    status=str(row.get("Item Status") or "").strip(),
                    available=num(row.get("AVAILABLE (AMZ)")),
                    reserved=num(row.get("RESERVED (AMZ)")),
                    total_stock=num(row.get("TOTAL STOCKS")),
                )
            )
        return [r for r in out if r.sku]

    raw = list(csv.reader(path.open(encoding="utf-8", newline="")))
    out = []
    for row in raw[5:]:
        if not row or not str(row[0]).strip().isdigit():
            continue
        if len(row) < 15:
            continue
        out.append(
            InventoryRow(
                sku=str(row[1] or row[3] or "").strip(),
                product=str(row[7] or "").strip(),
                style=str(row[8] or "").strip(),
                color=str(row[10] or "").strip(),
                status=str(row[11] or "").strip(),
                available=num(row[12]),
                reserved=num(row[13]),
                total_stock=num(row[14]),
            )
        )
    log.info("IP Tracker inventory: loaded %s SKU rows", len(out))
    return [r for r in out if r.sku]


def build_deck_data(
    *,
    month: date,
    archive_path: Path,
    monthly_sales: list[SkuSalesRow],
    inventory: list[InventoryRow],
    sheet_month_label: str,
    velocity_per_day: float,
    ad_metrics: dict[str, AdChannelMetrics],
) -> DeckData:
    archive = load_archive_month(archive_path, month)
    prior = load_archive_month(archive_path, prior_month(month))

    sellers = sorted(monthly_sales, key=lambda r: r.month_units, reverse=True)
    top_sellers = [r for r in sellers if r.month_units > 0][:15]

    struggling = [
        r
        for r in sellers
        if r.prior_units >= 20
        and r.month_units < r.prior_units * 0.5
        and r.status.lower() not in {"discontinued"}
    ]
    struggling = sorted(
        struggling,
        key=lambda r: (r.prior_units - r.month_units),
        reverse=True,
    )[:15]

    new_launches = [
        r
        for r in sellers
        if r.month_units > 0
        and (
            "development" in r.status.lower()
            or "new launch" in r.status.lower()
            or ("new" in r.status.lower() and r.prior_units == 0)
            or (r.sku.upper().startswith("WM-") and r.prior_units == 0)
        )
    ]
    # Color launch signals tracked by merchandising
    tracked_colors = (
        "Skyline Stripe",
        "Skyline Stripes",
        "Dusty Lavender",
        "Rusty Orange",
    )
    color_groups: list[tuple[str, list[SkuSalesRow]]] = []
    for color in ("Skyline Stripe", "Dusty Lavender", "Rusty Orange"):
        matches = [
            r
            for r in monthly_sales
            if color.lower() in r.color.lower()
            or (color == "Skyline Stripe" and "skyline stripe" in r.color.lower())
        ]
        matches = sorted(matches, key=lambda r: r.month_units, reverse=True)
        if matches:
            color_groups.append((color, matches))
            for row in matches:
                if row not in new_launches and row.month_units > 0:
                    new_launches.append(row)
    new_launches = sorted(new_launches, key=lambda r: r.month_units, reverse=True)[:15]

    # Attach days-of-cover using SKU velocity from monthly sales
    vel_by_sku = {r.sku: r.velocity_4wk / 7.0 for r in monthly_sales if r.velocity_4wk > 0}
    enriched: list[InventoryRow] = []
    for row in inventory:
        daily = vel_by_sku.get(row.sku, 0.0)
        cover = (row.total_stock / daily) if daily > 0 else None
        enriched.append(
            InventoryRow(
                sku=row.sku,
                product=row.product,
                style=row.style,
                color=row.color,
                status=row.status,
                available=row.available,
                reserved=row.reserved,
                total_stock=row.total_stock,
                days_of_cover=cover,
            )
        )

    inv_summary = {
        "sku_count": float(len(enriched)),
        "total_stock": float(sum(r.total_stock for r in enriched)),
        "available": float(sum(r.available for r in enriched)),
        "reserved": float(sum(r.reserved for r in enriched)),
        "oos_or_low": float(sum(1 for r in enriched if r.total_stock <= 5)),
    }
    low_stock = sorted(
        [r for r in enriched if 0 < r.total_stock <= 20],
        key=lambda r: r.total_stock,
    )[:20]
    oos_projections = sorted(
        [
            r
            for r in enriched
            if r.days_of_cover is not None and 0 < r.days_of_cover <= 45 and r.total_stock > 0
        ],
        key=lambda r: r.days_of_cover or 0,
    )[:15]

    b2b_top = [
        r
        for r in top_sellers
        if r.category.lower() in {"core", "fashion"}
        or r.style.lower() in {"trek", "golf", "travel"}
    ][:12]

    _ = tracked_colors  # documented merchandising signals
    return DeckData(
        month=month,
        archive=archive,
        prior_archive=prior,
        sheet_month_label=sheet_month_label,
        top_sellers=top_sellers,
        struggling=struggling,
        new_launches=new_launches,
        inventory_summary=inv_summary,
        low_stock=low_stock,
        oos_projections=oos_projections,
        b2b_top=b2b_top,
        velocity_per_day=velocity_per_day,
        ad_metrics=ad_metrics,
        color_launches=color_groups,
        ytd_by_month={},
    )


def apply_live_sources(
    data: DeckData,
    archive_path: Path,
    *,
    tw_export: Path | None = None,
    ptp_targets: Path | None = None,
) -> DeckData:
    """Overlay Shopify Admin + Sellerboard + Triple Whale onto archive rollups."""
    live = merge_live_bundle(month=data.month, archive_path=archive_path)
    data.ytd_by_month = live.ytd_by_month

    # Shopify Direct — prefer live Admin API totals
    if live.shopify.revenue > 0:
        ch = data.archive.channels["shopify_direct"]
        ch.revenue = live.shopify.revenue
        ch.orders = live.shopify.orders
        ch.units = live.shopify.units or ch.units
        ch.days_available = max(ch.days_available, live.shopify.days)
        shop_ads = data.ad_metrics.get("shopify_direct") or AdChannelMetrics()
        if live.shopify.roas is not None:
            shop_ads.roas = live.shopify.roas
        if live.shopify.cac is not None:
            shop_ads.cac = live.shopify.cac
        if live.shopify.mer is not None:
            # stash MER on cvr slot? keep separate — put in notes via cac companion
            pass
        data.ad_metrics["shopify_direct"] = shop_ads

    # Amazon — prefer Sellerboard (ACOS/ROAS + revenue)
    if live.amazon.revenue > 0 or live.amazon.acos is not None:
        ch = data.archive.channels["amazon"]
        if live.amazon.revenue > 0:
            ch.revenue = live.amazon.revenue
        if live.amazon.orders:
            ch.orders = live.amazon.orders
        if live.amazon.units:
            ch.units = live.amazon.units
        amz_ads = data.ad_metrics.get("amazon") or AdChannelMetrics()
        if live.amazon.acos is not None:
            amz_ads.acos = live.amazon.acos
        if live.amazon.tacos is not None:
            amz_ads.tacos = live.amazon.tacos
        if live.amazon.roas is not None:
            amz_ads.roas = live.amazon.roas
        if live.amazon.net_profit is not None:
            amz_ads.net_contribution = live.amazon.net_profit
        data.ad_metrics["amazon"] = amz_ads

    # Triple Whale — D2C attribution (PTP / Channel Mix + CAC/MER/ROAS)
    tw = fetch_triple_whale_month(
        data.month,
        export_path=tw_export or DEFAULT_TW_EXPORT,
        targets_path=ptp_targets or DEFAULT_PTP_TARGETS,
    )
    data.triple_whale = tw
    if tw is not None:
        shop_ads = data.ad_metrics.get("shopify_direct") or AdChannelMetrics()
        if tw.blended_roas is not None:
            shop_ads.roas = tw.blended_roas
        if tw.cac is not None:
            shop_ads.cac = tw.cac
        if tw.cvr is not None:
            shop_ads.cvr = tw.cvr
        if tw.repeat_rate is not None:
            shop_ads.repeat_rate = tw.repeat_rate
        data.ad_metrics["shopify_direct"] = shop_ads
        for note in tw.notes:
            log.info("Triple Whale: %s", note)

    for note in live.notes:
        log.info("Live source: %s", note)
    print("=== DEBUG live sources ===")
    print(
        f"Shopify {live.shopify.source}: {live.shopify.revenue} / {live.shopify.orders} "
        f"AOV={live.shopify.aov}"
    )
    print(
        f"Amazon {live.amazon.source}: {live.amazon.revenue} ACOS={live.amazon.acos} "
        f"ROAS={live.amazon.roas}"
    )
    if tw is not None:
        print(
            f"Triple Whale {tw.source}: sales={tw.order_revenue} "
            f"blendedAds={tw.blended_ads} roas={tw.blended_roas} "
            f"CAC={tw.cac} channels={len(tw.channels)}"
        )
        for row in tw.revenue_rows:
            print(
                f"  PTP rev {row.channel}: plan={row.planned} actual={row.actual} "
                f"Δ={row.delta_pct}"
            )
    else:
        print("Triple Whale: unavailable (set TRIPLE_WHALE_API_KEY or export JSON)")
    print("YTD months:", list(live.ytd_by_month))
    print("=== END live sources ===")
    return data


def _ptp_money(value: float | None, *, cents: bool = False) -> str:
    if value is None:
        return "n/a"
    if cents:
        return f"${value:,.2f}"
    return f"${value:,.0f}"


def _ptp_delta(pct: float | None) -> str:
    if pct is None:
        return "n/a"
    sign = "+" if pct >= 0 else ""
    # Match creative-deck style: whole % when |Δ|≥10, else one decimal.
    if abs(pct) >= 10:
        return f"{sign}{pct:.0f}%"
    return f"{sign}{pct:.1f}%"


# ---------------------------------------------------------------------------
# Google Slides plan builder
# ---------------------------------------------------------------------------


def _assert_not_protected(slide: int) -> None:
    if slide in PROTECTED_SLIDES or slide >= PROTECTED_SLIDE_MIN:
        raise RuntimeError(f"Refusing to modify protected slide {slide}")


def _add_text(plan: SlidesPlan, slide: int, object_id: str, text: str) -> None:
    _assert_not_protected(slide)
    plan.text_updates.append(TextUpdate(object_id=object_id, text=text, slide=slide))


def _add_delete(plan: SlidesPlan, slide: int, object_id: str) -> None:
    _assert_not_protected(slide)
    plan.deletes.append(DeleteOp(object_id=object_id, slide=slide))


def _add_box(
    plan: SlidesPlan,
    slide: int,
    text: str,
    *,
    x: float,
    y: float,
    width: float,
    height: float,
    font_size: int = 11,
    bold: bool = False,
) -> None:
    _assert_not_protected(slide)
    page = SLIDE_PAGE_IDS[slide]
    plan.create_text_boxes.append(
        CreateTextBoxOp(
            slide=slide,
            page_object_id=page,
            text=text,
            x=int(x * EMU_PER_INCH),
            y=int(y * EMU_PER_INCH),
            width=int(width * EMU_PER_INCH),
            height=int(height * EMU_PER_INCH),
            font_size=font_size,
            bold=bold,
        )
    )


def _channel_share(data: DeckData, key: str) -> str:
    total = float(data.archive.total_revenue)
    if not total:
        return "0%"
    return f"{float(data.archive.channels[key].revenue) / total * 100:.1f}%"


def debug_print_archive_metrics(data: DeckData) -> None:
    """Print archive values the deck will write into metric cards (stdout)."""
    print("=== DEBUG daily_archive.json metrics ===")
    print(f"Target month: {month_label(data.month)}")
    print(
        f"Archive days: {data.archive.day_count} "
        f"({data.archive.first_day} → {data.archive.last_day})"
    )
    print(
        f"Prior month:  {month_label(data.prior_archive.month)} "
        f"days={data.prior_archive.day_count}"
    )
    print(
        f"{'channel':<16} {'revenue':>14} {'units':>8} {'orders':>8} {'aov':>10} "
        f"{'prior_rev':>14}"
    )
    for key in PLATFORM_ORDER:
        cur = data.archive.channels[key]
        prev = data.prior_archive.channels[key]
        print(
            f"{key:<16} {money(cur.revenue):>14} {cur.units:>8,} {cur.orders:>8,} "
            f"{fmt_aov(cur.aov):>10} {money(prev.revenue):>14}"
        )
    blended_aov = aov(data.archive.total_revenue, data.archive.total_orders)
    print(
        f"{'TOTAL':<16} {money(data.archive.total_revenue):>14} "
        f"{data.archive.total_units:>8,} {data.archive.total_orders:>8,} "
        f"{fmt_aov(blended_aov):>10} {money(data.prior_archive.total_revenue):>14}"
    )
    print(
        "Card formats → blended TOTAL REVENUE "
        f"{money_whole(data.archive.total_revenue)} | "
        f"D2C TOTAL REVENUE {money_whole(data.archive.channels['shopify_direct'].revenue)} | "
        f"blended AOV {money_whole(blended_aov or 0)} | "
        f"D2C AOV {money_whole(data.archive.channels['shopify_direct'].aov or 0)}"
    )
    print("=== END DEBUG archive ===")


def build_metric_card_replacements(data: DeckData) -> list[ReplaceAllOp]:
    """Replace stale template / September-bleed numbers inside grouped metric cards.

    Slides 3–7 card tiles are ``elementGroup`` objects. Their child text frames
    are invisible to title-only objectId bindings, so we page-scope replaceAllText
    for known leftover strings (including inverted MoM arrows and August labels).
    """
    amz = data.archive.channels["amazon"]
    shop = data.archive.channels["shopify_direct"]
    prior_shop = data.prior_archive.channels["shopify_direct"]
    blended = data.archive.total_revenue
    prior_blended = data.prior_archive.total_revenue
    blended_aov = aov(blended, data.archive.total_orders)
    prior_blended_aov = aov(prior_blended, data.prior_archive.total_orders)
    prior_abbr = calendar.month_abbr[prior_month(data.month).month]
    tw = data.triple_whale
    shop_ads = data.ad_metrics.get("shopify_direct") or AdChannelMetrics()
    ads = data.ad_metrics.get("amazon") or AdChannelMetrics()

    ops: list[ReplaceAllOp] = []

    def add(contains: str, replace: str, slides: list[int]) -> None:
        if not contains or contains == replace:
            return
        ops.append(ReplaceAllOp(contains=contains, replace=replace, slides=slides))

    mom_blended = pct_change(blended, prior_blended) or "n/a"
    mom_shop = pct_change(shop.revenue, prior_shop.revenue) or "n/a"
    mom_aov = (
        pct_change(shop.aov or 0, prior_shop.aov or 0)
        if shop.aov and prior_shop.aov
        else "n/a"
    )
    mom_blended_aov = (
        pct_change(blended_aov or 0, prior_blended_aov or 0)
        if blended_aov and prior_blended_aov
        else "n/a"
    )

    # --- Slides 3–4: all-channel executive cards ---
    for old_total in ("$600,594", "$202,221", "$217,185", "$217,267"):
        add(old_total, money_whole(blended), [3, 4])
    add("$506,519", money_whole(prior_blended), [3])
    add("$537,909", money_whole(prior_blended), [3])

    # Fix inverted MoM on revenue (template still shows ▲ 19% vs Sep)
    for sep in ("\u00a0", " ", ""):
        add(
            f"Sep:{sep}{money_whole(prior_blended)}{sep}▲ 19%",
            f"Sep:{sep}{money_whole(prior_blended)}{sep}{mom_blended}",
            [3],
        )
        add(
            f"Sep: {money_whole(prior_blended)}{sep}▲ 19%",
            f"Sep: {money_whole(prior_blended)}{sep}{mom_blended}",
            [3],
        )
        add(
            f"August:{sep}$506,519{sep}▲ 19%",
            f"{prior_abbr}:{sep}{money_whole(prior_blended)}{sep}{mom_blended}",
            [3],
        )
        add(
            f"August:{sep}$506,519",
            f"{prior_abbr}:{sep}{money_whole(prior_blended)}",
            [3],
        )

    # Blended AOV + August bleed on efficiency cards
    add("$89", money_whole(blended_aov or 0), [3])
    add("$75", money_whole(blended_aov or 0), [3, 4])
    add("August: $95", f"{prior_abbr}: {money_whole(prior_blended_aov or 0)} {mom_blended_aov}", [3])
    for sep in ("\u00a0", " "):
        add(
            f"August:{sep}$95{sep}▼ 6%",
            f"{prior_abbr}:{sep}{money_whole(prior_blended_aov or 0)}{sep}{mom_blended_aov}",
            [3],
        )
        add(
            f"August:{sep}$95 ▼ 6%",
            f"{prior_abbr}:{sep}{money_whole(prior_blended_aov or 0)} {mom_blended_aov}",
            [3],
        )

    # ROAS / CAC — replace September leftovers with Triple Whale October
    blended_roas = (tw.blended_roas if tw and tw.blended_roas is not None else shop_ads.roas)
    blended_cac = (tw.cac if tw and tw.cac is not None else shop_ads.cac)
    if blended_roas is not None:
        add("3.83", f"{blended_roas:.2f}", [3, 4])
        add("August: 3.92", f"{prior_abbr}: 3.17", [3])
        for sep in ("\u00a0", " "):
            add(
                f"August:{sep}3.92{sep}▼ 2%",
                f"{prior_abbr}:{sep}3.17{sep}▲ {((blended_roas - 3.17) / 3.17 * 100):.0f}%",
                [3],
            )
            add(
                f"August: 3.92{sep}▼ 2%",
                f"{prior_abbr}: 3.17 {pct_change(blended_roas, 3.17) or 'n/a'}",
                [3],
            )
    if blended_cac is not None:
        add("$49", money_whole(blended_cac), [3, 4])
        add("August: $65", f"{prior_abbr}: $49", [3])
        for sep in ("\u00a0", " "):
            add(
                f"August:{sep}$65{sep}▼ 25%",
                f"{prior_abbr}:{sep}$49{sep}{pct_change(blended_cac, 48.66) or 'n/a'}",
                [3],
            )
            add(
                f"August: $65{sep}▼ 25%",
                f"{prior_abbr}: $49 {pct_change(blended_cac, 48.66) or 'n/a'}",
                [3],
            )

    # P&L + MER (exec) — replace TBD placeholders
    tw_sales = float((tw.summary or {}).get("sales") or 0) if tw else 0.0
    tw_ads = float((tw.summary or {}).get("blendedAds") or 0) if tw else 0.0
    d2c_contrib = tw_sales - tw_ads if tw_sales and tw_ads else None
    mer_pct = (tw_ads / tw_sales * 100.0) if tw_sales and tw_ads else None
    amz_net = float((tw.summary or {}).get("amazonNetProfit") or 0) if tw else 0.0
    if d2c_contrib is not None:
        add("Will have by next week", f"D2C contrib {money_whole(d2c_contrib)}", [3, 4, 5, 6])
        add("TBD", money_whole(d2c_contrib), [3, 4, 5, 6])
    if mer_pct is not None:
        add("3.43x", f"{mer_pct:.1f}%", [3, 4, 5, 6])
        add("31.55%", f"{mer_pct:.1f}%", [5, 6])
        add("MER\nTBD", f"MER\n{mer_pct:.1f}%", [3, 4, 5, 6])

    # --- Slides 5–6: D2C / Shopify ---
    for old in ("$162,304", "$40,990", "$40,989.99", "$55,954", "$55,953.80", "$56,037"):
        add(old, money_whole(shop.revenue), [5, 6])
    add("$148,876", money_whole(prior_shop.revenue), [5])
    add("$148,879", money_whole(prior_shop.revenue), [5])
    add("$117,295", money_whole(prior_shop.revenue), [5])
    for old_aov in ("$103", "$121", "$126", "$121.27", "$125.74"):
        add(old_aov, money_whole(shop.aov or 0), [5, 6])
    add("$115", money_whole(prior_shop.aov or 0), [5])
    add("$117", money_whole(prior_shop.aov or 0), [5])

    for sep in ("\u00a0", " ", "  "):
        add(
            f"August:{sep}$148,876",
            f"{prior_abbr}:{sep}{money_whole(prior_shop.revenue)}",
            [5],
        )
        add(
            f"August:{sep}$148,876 ▲ 9%",
            f"{prior_abbr}:{sep}{money_whole(prior_shop.revenue)} {mom_shop}",
            [5],
        )
        add(
            f"August:{sep}$115 ▼ 11%",
            f"{prior_abbr}:{sep}{money_whole(prior_shop.aov or 0)} {mom_aov}",
            [5],
        )
        add(
            f"August:{sep}$115 ▼ 10%",
            f"{prior_abbr}:{sep}{money_whole(prior_shop.aov or 0)} {mom_aov}",
            [5],
        )
        add(
            f"Sep:{sep}{money_whole(prior_shop.revenue)} ▼ 65.05%",
            f"Sep:{sep}{money_whole(prior_shop.revenue)} {mom_shop}",
            [5],
        )
        add(
            f"Sep: {money_whole(prior_shop.revenue)} ▼ 65.05%",
            f"Sep: {money_whole(prior_shop.revenue)} {mom_shop}",
            [5],
        )

    # D2C efficiency — replace September TW bleed with October TW
    if tw and tw.cvr is not None:
        add("0.82%", f"{tw.cvr:.2f}%", [5, 6])
        add("August: 0.70% ▲ 12pp", f"{prior_abbr}: 0.82% {pts_change(tw.cvr, 0.82) or 'n/a'}", [5])
        add("August: 0.70%", f"{prior_abbr}: 0.82%", [5])
    if blended_cac is not None:
        add("$48.66", money(blended_cac), [5, 6])
        add("August: $64.57 ▼ 25%", f"{prior_abbr}: $48.66 {pct_change(blended_cac, 48.66) or 'n/a'}", [5])
        add("August: $64.57", f"{prior_abbr}: $48.66", [5])
    if blended_roas is not None:
        add("3.17", f"{blended_roas:.2f}", [5, 6])
        add("August: 2.62 ▲ 21%", f"{prior_abbr}: 3.17 {pct_change(blended_roas, 3.17) or 'n/a'}", [5])
        add("August: 2.62", f"{prior_abbr}: 3.17", [5])
    if mer_pct is not None:
        add("31.55%", f"{mer_pct:.1f}%", [5, 6])
        add("August: ▼ 6.55pp", f"{prior_abbr}: 31.5% {pts_change(mer_pct, 31.55) or 'n/a'}", [5])
    if tw and tw.repeat_rate is not None:
        add("27%", f"{tw.repeat_rate:.0f}%", [5])
        add("25%", f"{tw.repeat_rate:.0f}%", [6])
        add("August: 26% ▲ 1pp", f"{prior_abbr}: 27% {pts_change(tw.repeat_rate, 27) or 'n/a'}", [5])
        add("August: 26%", f"{prior_abbr}: 27%", [5])

    # --- Slide 7: Site Health (September bleed: 126,556 / 87.17%) ---
    if tw:
        sessions = float((tw.summary or {}).get("pixelVisitSessions") or 0)
        bounce = float((tw.summary or {}).get("pixelBounceRate") or 0)
        if sessions:
            add("126,556", f"{sessions:,.0f}", [7])
            # Sep full-month TW sessions ≈ 82,396 → large MoM decline while MTD
            add("▲ 19% MoM", "▼ 71% MoM", [7])
            add("▲ 221% YoY", "YoY pending GA", [7])
        if bounce:
            add("87.17%", f"{bounce:.2f}%", [7])
            add("▼  0.25pp MoM", "▼ 5.65pp MoM", [7])
            add("▲ 9.62pp YoY", "YoY pending", [7])

    # Amazon slides 8–9 discrete fallbacks
    add("$429,438.78", money(amz.revenue), [8, 9])
    add("$429,438", money_whole(amz.revenue), [8, 9])
    add("$71.64", fmt_aov(amz.aov), [8, 9])

    retailers = (
        data.archive.channels["nordstrom"].revenue
        + data.archive.channels["dsg"].revenue
        + data.archive.channels["walmart"].revenue
    )
    add("$619,132", money_whole(blended), [])
    add(
        "$429,438 (69.4% share)",
        f"{money_whole(amz.revenue)} ({_channel_share(data, 'amazon')} share)",
        [],
    )
    add(
        f"{money_whole(shop.revenue)} (26.2% share)",
        f"{money_whole(shop.revenue)} ({_channel_share(data, 'shopify_direct')} share)",
        [],
    )
    add(
        "$17,399 (2.8% share)",
        f"{money_whole(retailers)} "
        f"({(float(retailers) / float(blended) * 100) if blended else 0:.1f}% share)",
        [],
    )

    if ads.roas is not None:
        add("5.27", f"{ads.roas:.2f}", [8])

    return ops


_MONEY_OR_PCT_RE = re.compile(
    r"^\$?\d[\d,]*\.?\d*%?$|^n/a$|^TBD$",
    re.IGNORECASE,
)


def _shape_plain_text(shape: dict[str, Any]) -> str:
    chunks: list[str] = []
    for te in (shape.get("text") or {}).get("textElements") or []:
        run = te.get("textRun") or {}
        chunks.append(run.get("content") or "")
    return "".join(chunks).replace("\x0b", "\n").strip()


def iter_slide_text_nodes(page_elements: list[dict[str, Any]]) -> list[tuple[str, str]]:
    """Walk shapes *and* elementGroup children — metric cards are grouped."""
    found: list[tuple[str, str]] = []

    def walk(elements: list[dict[str, Any]]) -> None:
        for el in elements or []:
            oid = el.get("objectId") or ""
            if "shape" in el and oid:
                text = _shape_plain_text(el["shape"])
                if text:
                    found.append((oid, text))
            children = (el.get("elementGroup") or {}).get("children") or []
            if children:
                walk(children)

    walk(page_elements)
    return found


def build_grouped_card_text_updates(
    *,
    slides_service: Any,
    presentation_id: str,
    data: DeckData,
) -> list[TextUpdate]:
    """Match label→value pairs inside groups and overwrite the value text frames."""
    shop = data.archive.channels["shopify_direct"]
    blended_aov = aov(data.archive.total_revenue, data.archive.total_orders)
    prior_abbr = calendar.month_abbr[prior_month(data.month).month]
    prior_shop = data.prior_archive.channels["shopify_direct"]

    # Per-slide label → new primary value (and optional subtitle rewrite).
    slide_card_values: dict[int, dict[str, str]] = {
        3: {
            "TOTAL REVENUE": money_whole(data.archive.total_revenue),
            "AOV": money_whole(blended_aov or 0),
        },
        4: {
            "TOTAL REVENUE": money_whole(data.archive.total_revenue),
            "AOV": money_whole(blended_aov or 0),
        },
        5: {
            "TOTAL REVENUE": money_whole(shop.revenue),
            "AOV": money_whole(shop.aov or 0),
        },
        6: {
            "TOTAL REVENUE": money_whole(shop.revenue),
            "AOV": money_whole(shop.aov or 0),
        },
    }
    slide_subtitles: dict[int, dict[str, str]] = {
        3: {
            "TOTAL REVENUE": (
                f"{prior_abbr}: {money_whole(data.prior_archive.total_revenue)} "
                f"{pct_change(data.archive.total_revenue, data.prior_archive.total_revenue) or 'n/a'}"
            ),
            "AOV": (
                f"{prior_abbr}: "
                f"{money_whole(aov(data.prior_archive.total_revenue, data.prior_archive.total_orders) or 0)} "
                f"{pct_change(blended_aov or 0, aov(data.prior_archive.total_revenue, data.prior_archive.total_orders) or 0) or 'n/a'}"
            ),
        },
        5: {
            "TOTAL REVENUE": (
                f"{prior_abbr}: {money_whole(prior_shop.revenue)} "
                f"{pct_change(shop.revenue, prior_shop.revenue) or 'n/a'}"
            ),
            "AOV": (
                f"{prior_abbr}: {money_whole(prior_shop.aov or 0)} "
                f"{pct_change(shop.aov or 0, prior_shop.aov or 0) or 'n/a'}"
            ),
        },
    }

    meta = (
        slides_service.presentations()
        .get(presentationId=presentation_id, fields="slides(objectId,pageElements)")
        .execute()
    )
    page_to_slide = {pid: num for num, pid in SLIDE_PAGE_IDS.items()}
    updates: list[TextUpdate] = []

    for slide in meta.get("slides") or []:
        page_id = slide.get("objectId")
        slide_num = page_to_slide.get(page_id or "")
        if slide_num is None or slide_num in PROTECTED_SLIDES:
            continue
        if slide_num not in slide_card_values:
            continue
        nodes = iter_slide_text_nodes(slide.get("pageElements") or [])
        print(f"=== DEBUG slide {slide_num} text nodes (incl. groups): {len(nodes)} ===")
        for oid, text in nodes[:40]:
            one = " ".join(text.split())
            if len(one) > 80:
                one = one[:80] + "…"
            print(f"  [{oid}] {one}")

        labels = slide_card_values[slide_num]
        subtitles = slide_subtitles.get(slide_num, {})
        for i, (oid, text) in enumerate(nodes):
            label = " ".join(text.split()).upper()
            if label not in labels:
                continue
            # Primary value is typically the next metric-looking sibling in the group walk.
            for j in range(i + 1, min(i + 6, len(nodes))):
                v_oid, v_text = nodes[j]
                compact = " ".join(v_text.split())
                if not _MONEY_OR_PCT_RE.match(compact.replace(" ", "")) and not compact.startswith("$"):
                    # allow "August: …" as subtitle candidate later
                    if not compact.lower().startswith(("august", "sep", "oct", "nov", "dec", "jan", "feb", "mar", "apr", "may", "jun", "jul")):
                        continue
                if compact.startswith("$") or _MONEY_OR_PCT_RE.match(compact.replace(",", "").replace("%", "")):
                    new_val = labels[label]
                    if compact != new_val:
                        print(
                            f"  → card '{label}' object {v_oid}: {compact!r} → {new_val!r}"
                        )
                        updates.append(TextUpdate(object_id=v_oid, text=new_val, slide=slide_num))
                    # Subtitle often follows the value
                    if label in subtitles and j + 1 < len(nodes):
                        s_oid, s_text = nodes[j + 1]
                        s_compact = " ".join(s_text.split())
                        if s_compact.lower().startswith(
                            ("august", "sep", "sept", "oct", "october", "nov", "dec", "jan")
                        ) or "▲" in s_compact or "▼" in s_compact:
                            new_sub = subtitles[label]
                            if s_compact != new_sub:
                                print(
                                    f"  → subtitle '{label}' object {s_oid}: "
                                    f"{s_compact!r} → {new_sub!r}"
                                )
                                updates.append(
                                    TextUpdate(object_id=s_oid, text=new_sub, slide=slide_num)
                                )
                    break
    return updates


def _amazon_narratives(data: DeckData, ads: AdChannelMetrics, prior_ads: AdChannelMetrics) -> tuple[str, str, str]:
    cur = data.archive.channels["amazon"]
    prev = data.prior_archive.channels["amazon"]
    mom = pct_change(cur.revenue, prev.revenue) or "n/a"
    wins = (
        f"Revenue {mom} MoM to {money(cur.revenue)} vs {money(prev.revenue)} "
        f"({data.archive.day_count} archive day(s)). Volume: {cur.orders:,} orders / "
        f"{cur.units:,} units; AOV {fmt_aov(cur.aov)}."
    )
    if ads.roas is not None and prior_ads.roas is not None:
        risks = (
            f"Advertising efficiency: ROAS {ads.roas:.2f} vs prior {prior_ads.roas:.2f}; "
            f"ACOS {fmt_metric(ads.acos, kind='pct')}."
        )
    else:
        risks = (
            "ROAS / ACOS / CAC not in daily_archive — supply data/monthly_ad_metrics.json "
            "or keep Sellerboard/Triple Whale feeds wired for paid efficiency cards."
        )
    decisions = (
        "Protect Amazon scale while the month is still MTD; refresh paid-efficiency cards "
        "once Real ACOS / ad spend for the month is available. Watch OOS on Travel/Trek cores."
    )
    return wins, risks, decisions


def build_slides_plan(data: DeckData, presentation_id: str) -> SlidesPlan:
    plan = SlidesPlan(presentation_id=presentation_id, month=month_label(data.month))
    label = month_label(data.month)
    prior_label = month_label(data.prior_archive.month)
    cur_start = data.archive.first_day or data.month
    cur_end = data.archive.last_day or month_end(data.month)
    prior_start = data.prior_archive.first_day or prior_month(data.month)
    prior_end = data.prior_archive.last_day or month_end(prior_month(data.month))

    amz = data.archive.channels["amazon"]
    amz_prior = data.prior_archive.channels["amazon"]
    shop = data.archive.channels["shopify_direct"]
    shop_prior = data.prior_archive.channels["shopify_direct"]
    ads = data.ad_metrics.get("amazon") or AdChannelMetrics()
    ads_prior = AdChannelMetrics()  # prior-month ads optional; leave blank → n/a deltas
    shop_ads = data.ad_metrics.get("shopify_direct") or AdChannelMetrics()

    if not data.ad_metrics:
        plan.notes.append(
            "No data/monthly_ad_metrics.json — ROAS/CAC/ACOS/CVR cards set to n/a "
            "(revenue/AOV/units still live from archive)."
        )
    if data.sheet_month_label.split()[0] != calendar.month_name[data.month.month]:
        plan.notes.append(
            f"IP Tracker Monthly Sales column used: {data.sheet_month_label} "
            f"(current month column not yet on sheet)."
        )

    # --- Slides 3–6 titles + key totals embedded ---
    blended = money(data.archive.total_revenue)
    blended_prior = money(data.prior_archive.total_revenue)
    blended_aov_val = aov(data.archive.total_revenue, data.archive.total_orders)
    tw = data.triple_whale
    tw_sales = float((tw.summary or {}).get("sales") or 0) if tw else 0.0
    tw_ads = float((tw.summary or {}).get("blendedAds") or 0) if tw else 0.0
    # MER = Total Ad Spend ÷ Blended Revenue (TW convention, as %)
    mer_pct = (tw_ads / tw_sales * 100.0) if tw_sales and tw_ads else None
    d2c_contrib = (tw_sales - tw_ads) if tw_sales and tw_ads else None
    amz_net = float((tw.summary or {}).get("amazonNetProfit") or 0) if tw else None
    _add_text(
        plan,
        3,
        "h3a7aee58bd343859_3_1456",
        f"1. Shopify + Amazon — {label} vs {prior_label}\n"
        f"MTD {blended} vs {blended_prior} ({pct_change(data.archive.total_revenue, data.prior_archive.total_revenue) or 'n/a'}). "
        f"Amazon {_channel_share(data, 'amazon')} · Shopify {_channel_share(data, 'shopify_direct')} · "
        f"AOV {money_whole(blended_aov_val or 0)} · ROAS {fmt_metric(shop_ads.roas or (tw.blended_roas if tw else None), kind='roas')} · "
        f"CAC {fmt_metric(shop_ads.cac or (tw.cac if tw else None), kind='money')} · "
        f"MER {f'{mer_pct:.1f}%' if mer_pct is not None else '—'} (ads÷revenue).",
    )
    _add_text(
        plan,
        4,
        "h1872ece28e3fae88_75_189",
        f"1. Shopify + Amazon — {label} YoY (2025 archive pending)\n"
        f"MTD {blended} · Amazon {money_whole(amz.revenue)} · Shopify {money_whole(shop.revenue)} · "
        f"AOV {money_whole(blended_aov_val or 0)}. YoY needs 2025 archive; MoM baseline {prior_label}.",
    )
    _add_text(
        plan,
        5,
        "h6e6e0e983a976184_12_160",
        f"3. D2C / Shopify — {label} vs {prior_label}\n"
        f"Revenue {money_whole(shop.revenue)} · Orders {shop.orders:,} · AOV {money_whole(shop.aov or 0)} · "
        f"MoM {pct_change(shop.revenue, shop_prior.revenue) or '—'} · "
        f"CVR {f'{(tw.cvr if tw else shop_ads.cvr) or 0:.2f}%' if (tw and tw.cvr) or shop_ads.cvr else '—'} · "
        f"CAC {fmt_metric(shop_ads.cac or (tw.cac if tw else None), kind='money')} · "
        f"ROAS {fmt_metric(shop_ads.roas or (tw.blended_roas if tw else None), kind='roas')} · "
        f"MER {f'{mer_pct:.1f}%' if mer_pct is not None else '—'} · "
        f"Returning {f'{(tw.repeat_rate if tw else 0):.0f}%' if tw and tw.repeat_rate is not None else '—'}.",
    )
    _add_text(
        plan,
        6,
        "h1872ece28e3fae88_75_1",
        f"3. D2C / Shopify — {label} YoY\n"
        f"MTD revenue {money(shop.revenue)} · AOV {fmt_aov(shop.aov)} · "
        f"CAC {fmt_metric(shop_ads.cac or (tw.cac if tw else None), kind='money')} · "
        f"ROAS {fmt_metric(shop_ads.roas or (tw.blended_roas if tw else None), kind='roas')}. "
        f"YoY % requires 2025 archive; MoM baseline is {prior_label}.",
    )

    # Slide 7 — Site Health: update creative cards in-place (no full-slide overlay).
    # Overlays previously collided with grouped metric tiles; keep Next Action only.
    if tw is not None:
        sessions = float((tw.summary or {}).get("pixelVisitSessions") or 0)
        bounce = float((tw.summary or {}).get("pixelBounceRate") or 0)
        _add_text(
            plan,
            7,
            "g409c5cc28db_0_3085",
            "Next Action: Continue bot blocking + Shopify investigation. "
            f"TW sessions {sessions:,.0f} · bounce {bounce:.1f}%. "
            "Treat Sep 126k / 87% as prior-month bleed, not October.",
        )

    # --- Slide 8 Amazon MoM metric cards ---
    # Primary values: whole dollars / short strings so they fit card frames.
    # Prior values: compact "$406K" form — full "$405,663.73" overflows into the
    # primary value box (visual double-number bug).
    prior_abbr = calendar.month_abbr[prior_month(data.month).month]
    cur_mc = LiveMonthChannel(
        revenue=amz.revenue,
        units=amz.units,
        orders=amz.orders,
        acos=ads.acos,
        tacos=ads.tacos,
        roas=ads.roas,
        net_profit=ads.net_contribution,
    )
    prior_mc = LiveMonthChannel(
        revenue=amz_prior.revenue,
        units=amz_prior.units,
        orders=amz_prior.orders,
        acos=ads_prior.acos,
        tacos=ads_prior.tacos,
        roas=ads_prior.roas,
        net_profit=ads_prior.net_contribution,
    )
    wins, risks, decisions = executive_amazon_copy(cur_mc, prior_mc, label)
    _add_text(plan, 8, "h1872ece28e3fae88_76_199", f"3. Amazon — {label} vs {prior_label}")
    _add_text(plan, 8, "h1872ece28e3fae88_76_205", live_money_whole(amz.revenue))
    _add_text(plan, 8, "h1872ece28e3fae88_76_247", money_compact_k(amz_prior.revenue))
    _add_text(plan, 8, "h1872ece28e3fae88_76_206", f"{prior_abbr}:")
    _add_text(plan, 8, "h1872ece28e3fae88_76_208", pct_change(amz.revenue, amz_prior.revenue) or "—")
    _add_text(plan, 8, "h1872ece28e3fae88_76_213", fmt_metric(ads.roas, kind="roas"))
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_248",
        f"{ads_prior.roas:.2f}" if ads_prior.roas is not None else "—",
    )
    _add_text(plan, 8, "h1872ece28e3fae88_76_217", live_money_whole(amz.aov or 0))
    _add_text(plan, 8, "h1872ece28e3fae88_76_250", live_money_whole(amz_prior.aov or 0))
    aov_delta = (
        pct_change(amz.aov, amz_prior.aov)
        if amz.aov is not None and amz_prior.aov is not None
        else "—"
    )
    _add_text(plan, 8, "h1872ece28e3fae88_76_265", aov_delta or "—")
    _add_text(plan, 8, "h1872ece28e3fae88_76_221", fmt_metric(ads.cvr, kind="pct") if ads.cvr is not None else "—")
    _add_text(plan, 8, "h1872ece28e3fae88_76_251", fmt_metric(ads_prior.cvr, kind="pct") if ads_prior.cvr is not None else "—")
    _add_text(plan, 8, "h1872ece28e3fae88_76_224", pts_change(ads.cvr, ads_prior.cvr) or "—")
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_227",
        f"{ads.acos:.1f}%" if ads.acos is not None else "—",
    )
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_252",
        f"{ads_prior.acos:.1f}%" if ads_prior.acos is not None else "—",
    )
    _add_text(plan, 8, "h1872ece28e3fae88_76_258", pts_change(ads.acos, ads_prior.acos) or "—")
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_231",
        f"{ads.tacos:.1f}%" if ads.tacos is not None else "—",
    )
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_253",
        f"{ads_prior.tacos:.1f}%" if ads_prior.tacos is not None else "—",
    )
    _add_text(plan, 8, "h1872ece28e3fae88_76_259", pts_change(ads.tacos, ads_prior.tacos) or "—")
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_235",
        fmt_metric(ads.repeat_rate, kind="pct") if ads.repeat_rate is not None else "—",
    )
    _add_text(plan, 8, "h1872ece28e3fae88_76_254", "—")
    _add_text(plan, 8, "h1872ece28e3fae88_76_267", "—")
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_246",
        live_money_whole(ads.net_contribution)
        if ads.net_contribution is not None
        else "—",
    )
    _add_text(plan, 8, "h1872ece28e3fae88_76_261", "—")
    _add_text(plan, 8, "h1872ece28e3fae88_76_263", "—")
    _add_text(
        plan,
        8,
        "h1872ece28e3fae88_76_271",
        pct_change(ads.roas or 0, ads_prior.roas or 0)
        if ads.roas is not None and ads_prior.roas is not None
        else "—",
    )
    for oid in (
        "h1872ece28e3fae88_76_214",
        "h1872ece28e3fae88_76_218",
        "h1872ece28e3fae88_76_222",
        "h1872ece28e3fae88_76_228",
        "h1872ece28e3fae88_76_232",
        "h1872ece28e3fae88_76_236",
        "h1872ece28e3fae88_76_260",
    ):
        _add_text(plan, 8, oid, f"{prior_abbr}:")
    _add_text(plan, 8, "h1872ece28e3fae88_76_268", wins)
    _add_text(plan, 8, "h1872ece28e3fae88_76_269", risks)
    _add_text(plan, 8, "h1872ece28e3fae88_76_270", decisions)

    # --- Slide 9 Amazon YoY (current metrics live; 2025 baseline kept as n/a without archive) ---
    _add_text(plan, 9, "h1872ece28e3fae88_76_275", f"4. Amazon — {label} vs {calendar.month_name[data.month.month]} 2025")
    _add_text(plan, 9, "h1872ece28e3fae88_76_281", money(amz.revenue))
    _add_text(plan, 9, "h1872ece28e3fae88_76_289", fmt_metric(ads.roas, kind="roas"))
    _add_text(plan, 9, "h1872ece28e3fae88_76_292", fmt_aov(amz.aov))
    _add_text(plan, 9, "h1872ece28e3fae88_76_295", fmt_metric(ads.cvr, kind="pct"))
    _add_text(plan, 9, "h1872ece28e3fae88_76_298", fmt_metric(ads.acos, kind="pct"))
    _add_text(plan, 9, "h1872ece28e3fae88_76_301", fmt_metric(ads.tacos, kind="pct"))
    _add_text(plan, 9, "h1872ece28e3fae88_76_304", fmt_metric(ads.repeat_rate, kind="pct"))
    _add_text(
        plan,
        9,
        "h1872ece28e3fae88_76_317",
        money(ads.net_contribution) if ads.net_contribution is not None else "n/a",
    )
    _add_text(plan, 9, "h1872ece28e3fae88_76_318", "n/a")  # prior-year revenue without 2025 archive
    _add_text(plan, 9, "h1872ece28e3fae88_76_284", "n/a YoY")
    _add_text(
        plan,
        9,
        "h1872ece28e3fae88_76_308",
        f"MTD Amazon revenue {money(amz.revenue)} across {amz.orders:,} orders / {amz.units:,} units "
        f"(AOV {fmt_aov(amz.aov)}). YoY % not computed — 2025-{data.month.strftime('%m')} not in archive.",
    )
    _add_text(plan, 9, "h1872ece28e3fae88_76_312", risks)
    _add_text(
        plan,
        9,
        "h1872ece28e3fae88_76_316",
        "Backfill 2025 daily_archive days to restore YoY deltas; keep PPC efficiency review tied to Real ACOS.",
    )

    # --- Slide 10 retailers ---
    dsg = data.archive.channels["dsg"]
    dsg_p = data.prior_archive.channels["dsg"]
    nord = data.archive.channels["nordstrom"]
    nord_p = data.prior_archive.channels["nordstrom"]
    wal = data.archive.channels["walmart"]
    wal_p = data.prior_archive.channels["walmart"]
    _add_text(
        plan,
        10,
        "h1872ece28e3fae88_13_422",
        f"Retailers — {label} vs {prior_label}\n"
        f"Reporting Period: {period_label(cur_start, cur_end)} vs {period_label(prior_start, prior_end)}",
    )
    # Retailers — update creative cards in-place (no mid-slide table overlay).
    _add_text(plan, 10, "h1872ece28e3fae88_13_423", "DSG")
    _add_text(
        plan,
        10,
        "h1872ece28e3fae88_13_437",
        f"Revenue\n{live_money_whole(dsg.revenue)}\n{pct_change(dsg.revenue, dsg_p.revenue) or '—'} MoM",
    )
    _add_text(plan, 10, "h1872ece28e3fae88_13_440", "Nordstrom")
    _add_text(
        plan,
        10,
        "h1872ece28e3fae88_13_446",
        f"Orders\n{nord.orders:,}\n{pct_change(nord.orders, nord_p.orders) or '—'} MoM",
    )
    _add_text(
        plan,
        10,
        "h1872ece28e3fae88_13_450",
        f"Units Sold\n{nord.units:,}\nvs {prior_abbr}",
    )
    _add_text(
        plan,
        10,
        "h1872ece28e3fae88_13_454",
        f"Revenue\n{live_money_whole(nord.revenue)}\n"
        f"{pct_change(nord.revenue, nord_p.revenue) or '—'} MoM",
    )

    # --- Slides 12–13 YTD charts ---
    # Titles only here. Chart PNGs are generated + inserted separately
    # (matplotlib → insertSlidesLocalImage) so we never overlay ASCII tables.
    ytd = getattr(data, "ytd_by_month", {}) or {}
    oct_total = money_whole((ytd.get(data.month.strftime("%Y-%m")) or {}).get("total") or data.archive.total_revenue)
    _add_text(
        plan,
        12,
        "g3f8cdd12cc6_3_1292",
        f"YTD All Channel Revenue Trend (w/o FOH) — {label} MTD {oct_total}",
    )
    _add_text(
        plan,
        13,
        "g409c5cc28db_0_3418",
        f"YTD All Channel Revenue Trend (w/ FOH proxy) — {label} MTD {oct_total}",
    )
    _add_text(plan, 14, "g3f366a59143_0_625", f"PTP vs Actuals - {calendar.month_name[data.month.month]}")
    tw = data.triple_whale
    if tw is not None and tw.revenue_rows:
        # Revenue-by-channel object IDs (creative deck)
        rev_ids = {
            "Paid Social": ("g3f366a59143_0_635", "g3f366a59143_0_636", "g3f366a59143_0_637"),
            "Paid Search": ("g3f366a59143_0_640", "g3f366a59143_0_641", "g3f366a59143_0_642"),
            "Email/SMS": ("g3f366a59143_0_645", "g3f366a59143_0_646", "g3f366a59143_0_647"),
            "Organic/Direct": ("g3f366a59143_0_650", "g3f366a59143_0_651", "g3f366a59143_0_652"),
            "Affiliates": ("g3f39a1b9ef1_0_502", "g3f39a1b9ef1_0_503", "g3f39a1b9ef1_0_504"),
        }
        for row in tw.revenue_rows:
            ids = rev_ids.get(row.channel)
            if not ids:
                continue
            plan_id, actual_id, delta_id = ids
            _add_text(plan, 14, plan_id, _ptp_money(row.planned))
            _add_text(plan, 14, actual_id, _ptp_money(row.actual))
            _add_text(plan, 14, delta_id, _ptp_delta(row.delta_pct))

        spend_ids = {
            "Paid Social": ("g3f366a59143_0_662", "g3f366a59143_0_663", "g3f366a59143_0_664"),
            "Paid Search": ("g3f366a59143_0_667", "g3f366a59143_0_668", "g3f366a59143_0_669"),
        }
        for row in tw.spend_rows:
            ids = spend_ids.get(row.channel)
            if not ids:
                continue
            plan_id, actual_id, delta_id = ids
            _add_text(plan, 14, plan_id, _ptp_money(row.planned, cents=True))
            _add_text(plan, 14, actual_id, _ptp_money(row.actual, cents=True))
            _add_text(plan, 14, delta_id, _ptp_delta(row.delta_pct))

        # ROAS cells under spend table (Paid Social / Paid Search)
        social_roas = next((r.roas for r in tw.spend_rows if r.channel == "Paid Social"), None)
        search_roas = next((r.roas for r in tw.spend_rows if r.channel == "Paid Search"), None)
        _add_text(plan, 14, "g3f366a59143_0_685", fmt_metric(social_roas, kind="roas"))
        _add_text(plan, 14, "g3f366a59143_0_686", fmt_metric(search_roas, kind="roas"))

        period = f"{tw.start.isoformat()}→{tw.end.isoformat()}"
        _add_text(
            plan,
            14,
            "g3f366a59143_0_680",
            f"Source: Triple Whale ({tw.source}) {period} · shop {tw.shop_id} · "
            f"blended ROAS {fmt_metric(tw.blended_roas, kind='roas')} · "
            f"CAC {fmt_metric(tw.cac, kind='money')} · MER {fmt_metric(tw.mer, kind='pct')}",
        )
        _add_text(
            plan,
            14,
            "g3f366a59143_0_683",
            f"Next Action: TW D2C order revenue {_ptp_money(tw.order_revenue)} · "
            f"blended ads {_ptp_money(tw.blended_ads, cents=True)} · "
            f"all-channel MTD {blended}. Reallocate toward channels with "
            f"positive PTP Δ; watch Paid Social ROAS {fmt_metric(social_roas, kind='roas')}.",
        )
    else:
        _add_text(plan, 14, "g3f366a59143_0_651", money(shop.revenue))
        _add_text(plan, 14, "g3f366a59143_0_685", fmt_metric(shop_ads.roas or ads.roas, kind="roas"))
        _add_text(
            plan,
            14,
            "g3f366a59143_0_680",
            "Source: Triple Whale unavailable — set TRIPLE_WHALE_API_KEY or "
            "data/triple_whale_export.json (planned targets: data/ptp_targets.json)",
        )
        _add_text(
            plan,
            14,
            "g3f366a59143_0_683",
            f"Next Action: Archive MTD total {blended}. Connect Triple Whale for paid "
            f"social/search planned-vs-actual; CAC currently "
            f"{fmt_metric(shop_ads.cac or ads.cac, kind='money')}.",
        )
        plan.notes.append("Slide 14 PTP: Triple Whale feed missing — placeholders retained")

    # Slide 15 — Channel Mix captions only (chart PNG inserted separately).
    prior_label = month_label(data.prior_archive.month)
    if tw is not None and tw.mix_rows:
        prior_total = sum(r.prior_revenue for r in tw.mix_rows) or 1.0
        cur_total = sum(r.current_revenue for r in tw.mix_rows) or 1.0
        _add_text(
            plan,
            15,
            "g3fa8ada2d55_123_169",
            f"Source: Triple Whale Pixel ({tw.source}) {tw.start.isoformat()}→{tw.end.isoformat()} · "
            f"chart = {prior_label} vs {label} MTD",
        )
        _add_text(
            plan,
            15,
            "g3fa8ada2d55_123_172",
            f"Next Action: {label} TW attributed {_ptp_money(cur_total)} vs {prior_label} "
            f"{_ptp_money(prior_total)}. Shopify AOV {fmt_aov(shop.aov)} · Amazon AOV "
            f"{fmt_aov(amz.aov)} · CAC {fmt_metric(tw.cac, kind='money')} · "
            f"blended ROAS {fmt_metric(tw.blended_roas, kind='roas')} · "
            f"returning {fmt_metric(tw.repeat_rate, kind='pct')}.",
        )
    else:
        _add_text(
            plan,
            15,
            "g3fa8ada2d55_123_172",
            f"Next Action: {label} MTD Shopify AOV {fmt_aov(shop.aov)}, Amazon AOV "
            f"{fmt_aov(amz.aov)}. Connect Triple Whale to refresh Channel Mix attribution.",
        )
        plan.notes.append("Slide 15 Channel Mix: Triple Whale feed missing — chart unchanged")
    _add_text(
        plan,
        16,
        "g3f8cdd12cc6_132_1745",
        f"Next Action: {label} MTD — returning "
        f"{fmt_metric(tw.repeat_rate, kind='pct') if tw else '—'}. "
        f"Orders: Amazon {amz.orders:,} · Shopify {shop.orders:,}. "
        "Increase repeat purchase via post-purchase, win-back, and replenishment flows.",
    )

    # --- Slides 17–24 inventory ---
    inv = data.inventory_summary
    _add_text(
        plan,
        17,
        "g3f400dce2ff_0_45",
        f"INVENTORY SUMMARY\n"
        f"{int(inv['sku_count']):,} SKUs · Stock {inv['total_stock']:,.0f} · "
        f"Available {inv['available']:,.0f} · Reserved {inv['reserved']:,.0f} · "
        f"OOS/low (≤5) {int(inv['oos_or_low']):,}\n"
        f"4-wk velocity ≈ {data.velocity_per_day:,.0f} units/day · "
        f"Sheet month: {data.sheet_month_label}",
    )

    # Slide 19 best sellers — replace image with live table + caption
    for img_id in ("h431b9d78e3d0bc19_82_2",):
        _add_delete(plan, 19, img_id)
    _add_text(plan, 19, "h431b9d78e3d0bc19_7_17", f"Top 10 Best Sellers — {data.sheet_month_label} units")
    top_note = "; ".join(
        f"{r.style} {r.color} ({r.month_units:,.0f})" for r in data.top_sellers[:3]
    ) or "No SKU rows"
    _add_text(
        plan,
        19,
        "h431b9d78e3d0bc19_7_18",
        f"Leaders: {top_note}. Source: Monthly Sales for PPT.",
    )
    _add_box(
        plan,
        19,
        table_text(
            ["SKU", "Style", "Color", "Units", "Prior", "MoM%", "Vel/wk"],
            [
                [
                    r.sku[:18],
                    r.style[:10],
                    r.color[:14],
                    f"{r.month_units:,.0f}",
                    f"{r.prior_units:,.0f}",
                    f"{r.mom_pct:.0f}%" if r.mom_pct is not None else "—",
                    f"{r.velocity_4wk:,.0f}",
                ]
                for r in data.top_sellers[:10]
            ],
        ),
        x=0.35,
        y=1.5,
        width=12.6,
        height=5.5,
        font_size=10,
    )

    # Slide 20 B2B
    for img_id in ("h431b9d78e3d0bc19_7_37",):
        _add_delete(plan, 20, img_id)
    _add_text(plan, 20, "h431b9d78e3d0bc19_7_32", "Top Products For B2B — Trek / Golf / Travel")
    _add_text(
        plan,
        20,
        "h431b9d78e3d0bc19_7_33",
        f"Core/Fashion matrix from IP Tracker ({data.sheet_month_label}).",
    )
    _add_box(
        plan,
        20,
        table_text(
            ["SKU", "Style", "Color", "Cat", "Units", "YTD"],
            [
                [
                    r.sku[:18],
                    r.style[:10],
                    r.color[:12],
                    r.category[:8],
                    f"{r.month_units:,.0f}",
                    f"{r.ytd_units:,.0f}",
                ]
                for r in data.b2b_top[:10]
            ],
        ),
        x=0.35,
        y=1.5,
        width=12.6,
        height=5.5,
        font_size=10,
    )

    # Slide 21 struggling
    for img_id in ("h431b9d78e3d0bc19_7_9",):
        _add_delete(plan, 21, img_id)
    _add_text(plan, 21, "h431b9d78e3d0bc19_7_4", "Top 10 Struggling SKUs:")
    struggle_note = " · ".join(
        f"{r.style} {r.color} {r.prior_units:,.0f}→{r.month_units:,.0f}" for r in data.struggling[:4]
    ) or "None flagged at ≥50% MoM drop"
    _add_text(plan, 21, "h431b9d78e3d0bc19_7_5", struggle_note)
    _add_box(
        plan,
        21,
        table_text(
            ["SKU", "Style", "Color", "Prior", "Current", "Drop"],
            [
                [
                    r.sku[:18],
                    r.style[:10],
                    r.color[:12],
                    f"{r.prior_units:,.0f}",
                    f"{r.month_units:,.0f}",
                    f"{r.prior_units - r.month_units:,.0f}",
                ]
                for r in data.struggling[:10]
            ],
        ),
        x=0.35,
        y=1.8,
        width=12.6,
        height=5.2,
        font_size=10,
    )

    # Slide 22 DSG & Nordstrom metrics from archive
    for img_id in ("h431b9d78e3d0bc19_7_30",):
        _add_delete(plan, 22, img_id)
    _add_text(plan, 22, "h431b9d78e3d0bc19_7_24", f"Top Metrics Across DSG & Nordstrom — {label}")
    _add_text(
        plan,
        22,
        "h431b9d78e3d0bc19_7_25",
        f"DSG {money(dsg.revenue)} ({dsg.units} u) · Nordstrom {money(nord.revenue)} ({nord.units} u) · "
        f"Walmart {money(wal.revenue)} ({wal.units} u)",
    )
    _add_box(
        plan,
        22,
        table_text(
            ["Channel", "Revenue", "Units", "Orders", "AOV", f"Prior ({prior_label})", "MoM"],
            [
                [
                    PLATFORM_LABELS[key],
                    money(data.archive.channels[key].revenue),
                    f"{data.archive.channels[key].units:,}",
                    f"{data.archive.channels[key].orders:,}",
                    fmt_aov(data.archive.channels[key].aov),
                    money(data.prior_archive.channels[key].revenue),
                    pct_change(
                        data.archive.channels[key].revenue,
                        data.prior_archive.channels[key].revenue,
                    )
                    or "n/a",
                ]
                for key in ("dsg", "nordstrom", "walmart")
            ],
        ),
        x=0.35,
        y=1.6,
        width=12.6,
        height=5.0,
        font_size=12,
    )

    # Slide 23 new colors
    for img_id in ("h431b9d78e3d0bc19_82_3",):
        _add_delete(plan, 23, img_id)
    launch_bits = []
    for color, rows in data.color_launches:
        units = sum(r.month_units for r in rows)
        launch_bits.append(f"{color}: {units:,.0f} u across {len(rows)} SKU(s)")
    _add_text(plan, 23, "h431b9d78e3d0bc19_7_39", "New Colors Sales Updates:")
    _add_text(
        plan,
        23,
        "h431b9d78e3d0bc19_7_40",
        (" · ".join(launch_bits) if launch_bits else "No tracked color-launch rows in sheet.")
        + f" Sheet month: {data.sheet_month_label}.",
    )
    launch_rows: list[list[str]] = []
    for color, rows in data.color_launches:
        for r in rows[:4]:
            launch_rows.append(
                [
                    color[:14],
                    r.sku[:18],
                    r.style[:10],
                    f"{r.month_units:,.0f}",
                    f"{r.velocity_4wk:,.0f}",
                    r.status[:12],
                ]
            )
    _add_box(
        plan,
        23,
        table_text(
            ["Launch", "SKU", "Style", "Units", "Vel/wk", "Status"],
            launch_rows
            or [
                [
                    r.color[:14],
                    r.sku[:18],
                    r.style[:10],
                    f"{r.month_units:,.0f}",
                    f"{r.velocity_4wk:,.0f}",
                    r.status[:12],
                ]
                for r in data.new_launches[:10]
            ],
        ),
        x=0.35,
        y=1.8,
        width=12.6,
        height=5.2,
        font_size=10,
    )

    # Slide 24 OOS projections (text fields only — no images to delete)
    oos_lines = []
    for r in data.oos_projections[:6]:
        cover = f"{r.days_of_cover:.0f}d" if r.days_of_cover is not None else "?"
        oos_lines.append(
            f"{r.style} {r.color} — ~{cover} cover · stock {r.total_stock:,.0f} "
            f"(avail {r.available:,.0f})"
        )
    if not oos_lines:
        for r in data.low_stock[:6]:
            oos_lines.append(
                f"{r.style} {r.color} — stock {r.total_stock:,.0f} (low-stock ≤20)"
            )
    _add_text(plan, 24, "h431b9d78e3d0bc19_110_1", "Potential OOS SKUs")
    _add_text(
        plan,
        24,
        "h431b9d78e3d0bc19_110_2",
        "\n".join(oos_lines) if oos_lines else "No SKUs projected OOS within 45 days from velocity.",
    )

    # Slide 32+ intentionally untouched (team / appendix ownership).

    # Grouped metric-card replacements (slides 3–6 primarily)
    plan.replace_all.extend(build_metric_card_replacements(data))
    plan.notes.append(
        "Metric cards on slides 3–6 are elementGroups; title objectIds alone cannot "
        "reach them. Plan includes page-scoped replace_all + API group walk."
    )

    # Guard: ensure no protected slides leaked in
    for collection in (plan.text_updates, plan.deletes, plan.create_text_boxes):
        for op in collection:
            if op.slide in PROTECTED_SLIDES:
                raise RuntimeError(f"Protected slide leaked into plan: {op}")
    for op in plan.replace_all:
        if any(s in PROTECTED_SLIDES for s in op.slides):
            raise RuntimeError(f"Protected slide leaked into replace_all: {op}")

    log.info(
        "Slides plan: %s text updates, %s deletes, %s new text boxes, "
        "%s replace_all ops (skip %s)",
        len(plan.text_updates),
        len(plan.deletes),
        len(plan.create_text_boxes),
        len(plan.replace_all),
        sorted(PROTECTED_SLIDES),
    )
    return plan


def apply_slides_plan_api(
    plan: SlidesPlan,
    credentials_path: Path,
    *,
    data: DeckData | None = None,
) -> None:
    """Apply plan via Google Slides API (service account).

    Also walks ``elementGroup`` children on slides 3–6 so metric-card value
    text frames are updated even when they were invisible to title bindings.
    """
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    scopes = ["https://www.googleapis.com/auth/presentations"]
    creds = Credentials.from_service_account_file(str(credentials_path), scopes=scopes)
    slides = build("slides", "v1", credentials=creds, cache_discovery=False)

    # Discover + bind grouped card value objectIds before issuing writes.
    if data is not None:
        grouped = build_grouped_card_text_updates(
            slides_service=slides,
            presentation_id=plan.presentation_id,
            data=data,
        )
        existing = {t.object_id for t in plan.text_updates}
        for upd in grouped:
            if upd.object_id not in existing:
                plan.text_updates.append(upd)
                existing.add(upd.object_id)
        log.info("Grouped metric-card objectId updates: %s", len(grouped))

    requests: list[dict[str, Any]] = []

    # 1) Page-scoped replaceAllText for template card numbers inside groups
    for op in plan.replace_all:
        body: dict[str, Any] = {
            "replaceAllText": {
                "containsText": {"text": op.contains, "matchCase": op.match_case},
                "replaceText": op.replace,
            }
        }
        if op.slides:
            page_ids = [
                SLIDE_PAGE_IDS[s]
                for s in op.slides
                if s in SLIDE_PAGE_IDS and s not in PROTECTED_SLIDES
            ]
            if page_ids:
                body["replaceAllText"]["pageObjectIds"] = page_ids
        requests.append(body)
        print(f"replaceAllText [{op.slides}]: {op.contains!r} → {op.replace!r}")

    # 2) Deletes
    for d in plan.deletes:
        requests.append({"deleteObject": {"objectId": d.object_id}})

    # 3) Explicit text frames (titles + discovered card value objectIds)
    for t in plan.text_updates:
        requests.append({"deleteText": {"objectId": t.object_id, "textRange": {"type": "ALL"}}})
        if t.text:
            requests.append(
                {
                    "insertText": {
                        "objectId": t.object_id,
                        "insertionIndex": 0,
                        "text": t.text,
                    }
                }
            )

    # 4) New text boxes (inventory / slide 32 tables)
    for i, box in enumerate(plan.create_text_boxes):
        oid = f"mbr{box.slide}t{i}{int(datetime.now().timestamp()) % 100000}"
        requests.append(
            {
                "createShape": {
                    "objectId": oid,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": box.page_object_id,
                        "size": {
                            "width": {"magnitude": box.width, "unit": "EMU"},
                            "height": {"magnitude": box.height, "unit": "EMU"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": box.x,
                            "translateY": box.y,
                            "unit": "EMU",
                        },
                    },
                }
            }
        )
        requests.append({"insertText": {"objectId": oid, "insertionIndex": 0, "text": box.text}})
        requests.append(
            {
                "updateTextStyle": {
                    "objectId": oid,
                    "style": {
                        "fontSize": {"magnitude": box.font_size, "unit": "PT"},
                        "bold": box.bold,
                        "fontFamily": "Roboto Mono",
                    },
                    "textRange": {"type": "ALL"},
                    "fields": "fontSize,bold,fontFamily",
                }
            }
        )

    for start in range(0, len(requests), 40):
        chunk = requests[start : start + 40]
        slides.presentations().batchUpdate(
            presentationId=plan.presentation_id, body={"requests": chunk}
        ).execute()
        log.info("Applied API batch %s–%s", start, start + len(chunk))


def verify_pptx_metric_cards(path: Path) -> None:
    """Read back the generated PPTX and print metric-card values."""
    prs = Presentation(str(path))
    print(f"=== VERIFY PPTX {path.name} ({len(prs.slides)} slides) ===")
    for i, slide in enumerate(prs.slides, start=1):
        values: list[str] = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            for line in text.splitlines():
                line = line.strip()
                if line.startswith("$") or line in {"Revenue", "AOV", "ROAS", "CAC"}:
                    values.append(line)
                elif "Revenue" in text and line.startswith("$"):
                    values.append(line)
        # Also dump full card-style shape pairs
        texts = [
            shape.text_frame.text.strip().replace("\n", " | ")
            for shape in slide.shapes
            if shape.has_text_frame and shape.text_frame.text.strip()
        ]
        money_lines = [t for t in texts if "$" in t]
        if money_lines:
            print(f"Slide {i}:")
            for t in money_lines[:8]:
                print(f"  {t}")
    print("=== END VERIFY PPTX ===")


# ---------------------------------------------------------------------------
# PPTX builder
# ---------------------------------------------------------------------------


def _set_run(paragraph, text: str, *, size: int = 14, bold: bool = False, color=NAVY) -> None:
    paragraph.clear()
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"


def _add_title_bar(slide, title: str, subtitle: str = "") -> None:
    shape = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.9))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    _set_run(p, title, size=22, bold=True, color=WHITE)
    if subtitle:
        p2 = tf.add_paragraph()
        _set_run(p2, subtitle, size=12, color=WHITE)


def _add_table(slide, left, top, width, height, headers: list[str], rows: list[list[str]]):
    table_shape = slide.shapes.add_table(
        rows=1 + len(rows),
        cols=len(headers),
        left=left,
        top=top,
        width=width,
        height=height,
    )
    table = table_shape.table
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(10)
                run.font.color.rgb = WHITE
                run.font.name = "Calibri"
        cell.fill.solid()
        cell.fill.fore_color.rgb = TEAL
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(value)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
                    run.font.color.rgb = GRAY
                    run.font.name = "Calibri"
            if r_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT
    return table_shape


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build_presentation(data: DeckData) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    label = month_label(data.month)
    prior_label = month_label(data.prior_archive.month)
    ads = data.ad_metrics.get("amazon") or AdChannelMetrics()

    s = _blank_slide(prs)
    bar = s.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()
    title_box = s.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(11.5), Inches(1.2))
    _set_run(
        title_box.text_frame.paragraphs[0],
        "Weatherman — Monthly Business Review",
        size=32,
        bold=True,
        color=WHITE,
    )
    sub = s.shapes.add_textbox(Inches(0.8), Inches(3.6), Inches(11.5), Inches(0.6))
    _set_run(sub.text_frame.paragraphs[0], label, size=24, color=WHITE)

    s = _blank_slide(prs)
    _add_title_bar(s, "Executive summary — channel revenue", f"{label} vs {prior_label}")
    headers = ["Channel", "Revenue", "Units", "Orders", "AOV", f"Prior ({prior_label})", "MoM"]
    rows = []
    for key in PLATFORM_ORDER:
        cur = data.archive.channels[key]
        prev = data.prior_archive.channels[key]
        rows.append(
            [
                PLATFORM_LABELS[key],
                money(cur.revenue),
                f"{cur.units:,}",
                f"{cur.orders:,}",
                fmt_aov(cur.aov),
                money(prev.revenue),
                pct_change(cur.revenue, prev.revenue) or "n/a",
            ]
        )
    rows.append(
        [
            "TOTAL",
            money(data.archive.total_revenue),
            f"{data.archive.total_units:,}",
            f"{data.archive.total_orders:,}",
            fmt_aov(aov(data.archive.total_revenue, data.archive.total_orders)),
            money(data.prior_archive.total_revenue),
            pct_change(data.archive.total_revenue, data.prior_archive.total_revenue) or "n/a",
        ]
    )
    _add_table(s, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.0), headers, rows)
    note = s.shapes.add_textbox(Inches(0.4), Inches(6.5), Inches(12), Inches(0.4))
    _set_run(
        note.text_frame.paragraphs[0],
        f"Archive {data.archive.day_count} day(s) · Amazon ROAS {fmt_metric(ads.roas, kind='roas')} · "
        f"ACOS {fmt_metric(ads.acos, kind='pct')} · CAC {fmt_metric(ads.cac, kind='money')}",
        size=11,
        color=GRAY,
    )

    for key in PLATFORM_ORDER:
        s = _blank_slide(prs)
        cur = data.archive.channels[key]
        prev = data.prior_archive.channels[key]
        ch_ads = data.ad_metrics.get(key) or AdChannelMetrics()
        _add_title_bar(s, f"{PLATFORM_LABELS[key]} performance", label)
        cards = [
            ("Revenue", money(cur.revenue)),
            ("AOV", fmt_aov(cur.aov)),
            ("ROAS", fmt_metric(ch_ads.roas, kind="roas")),
            ("CAC", fmt_metric(ch_ads.cac, kind="money")),
            ("Units / Orders", f"{cur.units:,} / {cur.orders:,}"),
            ("MoM revenue", pct_change(cur.revenue, prev.revenue) or "n/a"),
        ]
        for i, (k, v) in enumerate(cards):
            left = Inches(0.5 + (i % 3) * 4.2)
            top = Inches(1.5 + (i // 3) * 2.2)
            shape = s.shapes.add_shape(1, left, top, Inches(3.9), Inches(1.8))
            shape.fill.solid()
            shape.fill.fore_color.rgb = LIGHT
            shape.line.color.rgb = TEAL
            tf = shape.text_frame
            tf.word_wrap = True
            _set_run(tf.paragraphs[0], k, size=12, color=GRAY)
            p = tf.add_paragraph()
            _set_run(p, v, size=22, bold=True, color=NAVY)

    # Inventory tables
    s = _blank_slide(prs)
    _add_title_bar(s, "Top best sellers", f"{data.sheet_month_label} — IP Tracker")
    _add_table(
        s,
        Inches(0.3),
        Inches(1.2),
        Inches(12.7),
        Inches(5.5),
        ["SKU", "Style", "Color", "Units", "Prior", "MoM %", "Vel/wk"],
        [
            [
                r.sku,
                r.style,
                r.color,
                f"{r.month_units:,.0f}",
                f"{r.prior_units:,.0f}",
                f"{r.mom_pct:.1f}%" if r.mom_pct is not None else "—",
                f"{r.velocity_4wk:,.0f}",
            ]
            for r in data.top_sellers[:12]
        ]
        or [["—", "—", "—", "0", "0", "—", "0"]],
    )

    s = _blank_slide(prs)
    _add_title_bar(s, "Potential OOS / low stock", "IP TRACKER - AMZ")
    _add_table(
        s,
        Inches(0.3),
        Inches(1.2),
        Inches(12.7),
        Inches(5.5),
        ["SKU", "Style", "Color", "Stock", "Avail", "Days cover", "Status"],
        [
            [
                r.sku,
                r.style,
                r.color,
                f"{r.total_stock:,.0f}",
                f"{r.available:,.0f}",
                f"{r.days_of_cover:.0f}" if r.days_of_cover is not None else "—",
                r.status,
            ]
            for r in (data.oos_projections or data.low_stock)[:12]
        ]
        or [["—", "—", "—", "0", "0", "—", "None"]],
    )

    s = _blank_slide(prs)
    _add_title_bar(s, "Appendix — generation stamp", label)
    box = s.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(5))
    tf = box.text_frame
    for i, line in enumerate(
        [
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            f"Archive days: {data.archive.day_count}",
            f"Sheet month column: {data.sheet_month_label}",
            f"IP Tracker: {DEFAULT_IP_TRACKER_ID}",
            "Protected Google Slides: 25, 26, 27 (untouched)",
        ]
    ):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _set_run(p, line, size=14, color=GRAY)

    return prs


# ---------------------------------------------------------------------------
# Optional rollover (title-only copy) — retained for brand-new month shells
# ---------------------------------------------------------------------------


def maybe_rollover_google_slides(
    *,
    template_id: str,
    month: date,
    credentials_path: Path | None,
) -> str | None:
    if credentials_path is None or not credentials_path.exists():
        log.info("Skipping Google Slides rollover (no --google-credentials)")
        return None
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError:
        log.warning("google-api-python-client not installed; skipping Slides rollover")
        return None

    scopes = [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/presentations",
    ]
    creds = Credentials.from_service_account_file(str(credentials_path), scopes=scopes)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    slides = build("slides", "v1", credentials=creds, cache_discovery=False)

    new_name = f"Weatherman — {month_label(month)} Monthly Business Review"
    copied = (
        drive.files()
        .copy(fileId=template_id, body={"name": new_name}, supportsAllDrives=True)
        .execute()
    )
    new_id = copied["id"]
    prior = prior_month(month)
    replacements = [
        (month_label(prior), month_label(month)),
        (calendar.month_name[prior.month], calendar.month_name[month.month]),
        (f"{calendar.month_name[prior.month]} {prior.year}", month_label(month)),
    ]
    requests_body = []
    for old, new in replacements:
        if old == new:
            continue
        requests_body.append(
            {
                "replaceAllText": {
                    "containsText": {"text": old, "matchCase": True},
                    "replaceText": new,
                }
            }
        )
    if requests_body:
        slides.presentations().batchUpdate(
            presentationId=new_id, body={"requests": requests_body}
        ).execute()
    log.info("Google Slides rollover created: %s (%s)", new_name, new_id)
    return new_id


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    p.add_argument("--month", help="Target month YYYY-MM (default: previous calendar month)")
    p.add_argument("--archive", type=Path, default=DEFAULT_ARCHIVE)
    p.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    p.add_argument("--ip-tracker-id", default=DEFAULT_IP_TRACKER_ID)
    p.add_argument("--monthly-sales-csv", type=Path)
    p.add_argument("--ip-tracker-csv", type=Path)
    p.add_argument("--ad-metrics", type=Path, default=DEFAULT_AD_METRICS)
    p.add_argument(
        "--triple-whale-export",
        type=Path,
        default=DEFAULT_TW_EXPORT,
        help="Optional offline Triple Whale JSON (used when API key missing)",
    )
    p.add_argument(
        "--ptp-targets",
        type=Path,
        default=DEFAULT_PTP_TARGETS,
        help="Optional PTP planned figures JSON (else prior-month TW actuals)",
    )
    p.add_argument("--skip-sheet-download", action="store_true")
    p.add_argument("--google-credentials", type=Path)
    p.add_argument("--slides-template-id", default=DEFAULT_SLIDES_TEMPLATE_ID)
    p.add_argument(
        "--presentation-id",
        default=DEFAULT_OCTOBER_DECK_ID,
        help="Creative Google Slides deck to populate (default: October 2026 copy)",
    )
    p.add_argument(
        "--apply-slides",
        action="store_true",
        help="Apply the Slides plan via Google API (requires --google-credentials)",
    )
    p.add_argument("--skip-slides-plan", action="store_true")
    p.add_argument("--skip-pptx", action="store_true")
    p.add_argument("--skip-slides-rollover", action="store_true", default=True)
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    month = parse_month(args.month)
    if not args.archive.exists():
        log.error("Archive not found: %s", args.archive)
        return 1

    cache_dir = ROOT / "data" / "templates"
    cache_dir.mkdir(parents=True, exist_ok=True)

    monthly_sales_csv = args.monthly_sales_csv
    if monthly_sales_csv is None and not args.skip_sheet_download:
        monthly_sales_csv = cache_dir / "monthly_sales_for_ppt.csv"
        download_csv(sheet_csv_url(args.ip_tracker_id, MONTHLY_SALES_GID), monthly_sales_csv)
    if monthly_sales_csv is None:
        monthly_sales_csv = cache_dir / "monthly_sales_for_ppt.csv"
    if not monthly_sales_csv.exists():
        log.error("Monthly Sales CSV missing: %s", monthly_sales_csv)
        return 1

    ip_csv = args.ip_tracker_csv
    if ip_csv is None and not args.skip_sheet_download:
        raw_path = cache_dir / "ip_tracker_export.csv"
        download_csv(sheet_csv_url(args.ip_tracker_id, IP_TRACKER_AMZ_GID), raw_path)
        ip_csv = raw_path
    if ip_csv is None:
        cleaned = cache_dir / "ip_tracker_amz_cleaned.csv"
        ip_csv = cleaned if cleaned.exists() else cache_dir / "ip_tracker_export.csv"
    if not ip_csv.exists():
        log.error("IP Tracker CSV missing: %s", ip_csv)
        return 1

    monthly_sales, sheet_month_label, velocity = load_monthly_sales_csv(
        monthly_sales_csv, month
    )
    inventory = load_inventory_csv(ip_csv)
    ad_metrics = load_ad_metrics(args.ad_metrics, month)
    data = build_deck_data(
        month=month,
        archive_path=args.archive,
        monthly_sales=monthly_sales,
        inventory=inventory,
        sheet_month_label=sheet_month_label,
        velocity_per_day=velocity,
        ad_metrics=ad_metrics,
    )
    data = apply_live_sources(
        data,
        args.archive,
        tw_export=args.triple_whale_export,
        ptp_targets=args.ptp_targets,
    )
    debug_print_archive_metrics(data)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    out_pptx = None
    if not args.skip_pptx:
        prs = build_presentation(data)
        out_name = f"Weatherman — {month_label(month)} Monthly Business Review.pptx"
        out_pptx = args.output_dir / out_name
        prs.save(out_pptx)
        log.info("Wrote %s (%s slides)", out_pptx, len(prs.slides))
        verify_pptx_metric_cards(out_pptx)

    plan_path = args.output_dir / f"slides_plan_{month.strftime('%Y-%m')}.json"
    if not args.skip_slides_plan:
        plan = build_slides_plan(data, args.presentation_id)
        plan_path.write_text(json.dumps(plan.to_json(), indent=2), encoding="utf-8")
        log.info("Wrote Slides plan %s", plan_path)
        print(
            f"=== DEBUG replace_all ops for grouped cards: {len(plan.replace_all)} ==="
        )
        for op in plan.replace_all:
            print(f"  slides {op.slides}: {op.contains!r} → {op.replace!r}")
        for note in plan.notes:
            log.warning("%s", note)

        if args.apply_slides:
            if not args.google_credentials or not args.google_credentials.exists():
                log.error("--apply-slides requires --google-credentials")
                return 1
            apply_slides_plan_api(plan, args.google_credentials, data=data)
            log.info(
                "Applied Slides updates to https://docs.google.com/presentation/d/%s/edit",
                args.presentation_id,
            )

    if not args.skip_slides_rollover and args.google_credentials:
        maybe_rollover_google_slides(
            template_id=args.slides_template_id,
            month=month,
            credentials_path=args.google_credentials,
        )

    if out_pptx:
        print(out_pptx)
    print(plan_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
