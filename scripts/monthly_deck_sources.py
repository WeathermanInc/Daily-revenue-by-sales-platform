#!/usr/bin/env python3
"""Live data loaders for the Monthly Business Review deck.

Sources (in priority order for each channel):
1. Shopify Admin API (OAuth client credentials — same as MCP)
2. Sellerboard automation CSV (same as daily pipeline) or
   ``data/sellerboard_daily_cache.json`` rebuilt from Actions logs
3. ``data/daily_archive.json`` fallback
"""
from __future__ import annotations

import json
import logging
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

log = logging.getLogger("monthly_deck.sources")
MONEY = Decimal("0.01")


@dataclass
class MonthChannel:
    revenue: Decimal = Decimal("0")
    units: int = 0
    orders: int = 0
    days: int = 0
    acos: float | None = None  # %
    tacos: float | None = None
    roas: float | None = None
    net_profit: float | None = None
    ad_spend: float | None = None
    cvr: float | None = None
    cac: float | None = None
    mer: float | None = None
    repeat_rate: float | None = None
    source: str = "none"

    @property
    def aov(self) -> Decimal | None:
        if not self.orders:
            return None
        return (self.revenue / Decimal(self.orders)).quantize(MONEY, rounding=ROUND_HALF_UP)


@dataclass
class LiveMonthBundle:
    month: date
    amazon: MonthChannel = field(default_factory=MonthChannel)
    shopify: MonthChannel = field(default_factory=MonthChannel)
    walmart: MonthChannel = field(default_factory=MonthChannel)
    nordstrom: MonthChannel = field(default_factory=MonthChannel)
    dsg: MonthChannel = field(default_factory=MonthChannel)
    ytd_by_month: dict[str, dict[str, float]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    @property
    def total_revenue(self) -> Decimal:
        return (
            self.amazon.revenue
            + self.shopify.revenue
            + self.walmart.revenue
            + self.nordstrom.revenue
            + self.dsg.revenue
        )


def _money(v: Any) -> Decimal:
    return Decimal(str(v or 0)).quantize(MONEY, rounding=ROUND_HALF_UP)


def month_end(d: date) -> date:
    if d.month == 12:
        return date(d.year, 12, 31)
    return date(d.year, d.month + 1, 1) - timedelta(days=1)


def prior_month_start(d: date) -> date:
    if d.month == 1:
        return date(d.year - 1, 12, 1)
    return date(d.year, d.month - 1, 1)


# ---------------------------------------------------------------------------
# Credentials
# ---------------------------------------------------------------------------


def load_shopify_credentials() -> dict[str, str]:
    """Resolve Shopify OAuth client credentials from env or Cursor MCP config."""
    domain = os.getenv("SHOPIFY_STORE_DOMAIN", "").strip()
    client_id = os.getenv("SHOPIFY_CLIENT_ID", "").strip()
    client_secret = os.getenv("SHOPIFY_CLIENT_SECRET", "").strip()
    if domain and client_id and client_secret:
        return {
            "SHOPIFY_STORE_DOMAIN": domain.replace("https://", "").replace("http://", "").rstrip("/"),
            "SHOPIFY_CLIENT_ID": client_id,
            "SHOPIFY_CLIENT_SECRET": client_secret,
        }
    mcp_path = Path.home() / ".cursor" / "mcp.json"
    if mcp_path.exists():
        cfg = json.loads(mcp_path.read_text(encoding="utf-8"))
        env = ((cfg.get("mcpServers") or {}).get("shopify") or {}).get("env") or {}
        domain = str(env.get("SHOPIFY_STORE_DOMAIN") or "").strip()
        client_id = str(env.get("SHOPIFY_CLIENT_ID") or "").strip()
        client_secret = str(env.get("SHOPIFY_CLIENT_SECRET") or "").strip()
        if domain and client_id and client_secret:
            log.info("Loaded Shopify credentials from %s", mcp_path)
            return {
                "SHOPIFY_STORE_DOMAIN": domain.replace("https://", "").replace("http://", "").rstrip("/"),
                "SHOPIFY_CLIENT_ID": client_id,
                "SHOPIFY_CLIENT_SECRET": client_secret,
            }
    raise RuntimeError(
        "Shopify credentials missing. Set SHOPIFY_STORE_DOMAIN / "
        "SHOPIFY_CLIENT_ID / SHOPIFY_CLIENT_SECRET or configure the shopify MCP."
    )


def shopify_access_token(creds: dict[str, str]) -> str:
    domain = creds["SHOPIFY_STORE_DOMAIN"]
    if not domain.endswith(".myshopify.com"):
        domain = f"{domain}.myshopify.com"
    url = f"https://{domain}/admin/oauth/access_token"
    body = urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": creds["SHOPIFY_CLIENT_ID"],
            "client_secret": creds["SHOPIFY_CLIENT_SECRET"],
        }
    ).encode()
    req = Request(url, data=body, method="POST", headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urlopen(req, timeout=60) as resp:  # noqa: S310
        payload = json.loads(resp.read().decode())
    token = payload.get("access_token")
    if not token:
        raise RuntimeError(f"Shopify OAuth failed: {payload}")
    return str(token)


def shopify_graphql(domain: str, token: str, query: str, variables: dict | None = None) -> dict:
    if not domain.endswith(".myshopify.com"):
        domain = f"{domain}.myshopify.com"
    url = f"https://{domain}/admin/api/2025-10/graphql.json"
    payload = json.dumps({"query": query, "variables": variables or {}}).encode()
    req = Request(
        url,
        data=payload,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "X-Shopify-Access-Token": token,
        },
    )
    with urlopen(req, timeout=90) as resp:  # noqa: S310
        data = json.loads(resp.read().decode())
    if data.get("errors"):
        raise RuntimeError(f"Shopify GraphQL errors: {data['errors']}")
    return data.get("data") or {}


# ---------------------------------------------------------------------------
# Shopify month rollup
# ---------------------------------------------------------------------------


def fetch_shopify_month(month: date) -> MonthChannel:
    """Sum D2C Shopify orders for the calendar month (excludes marketplace tags)."""
    creds = load_shopify_credentials()
    token = shopify_access_token(creds)
    domain = creds["SHOPIFY_STORE_DOMAIN"]
    start = month.replace(day=1)
    end = month_end(month) + timedelta(days=1)
    # Shopify search uses store timezone; date-only bounds are accepted.
    query_str = (
        f"created_at:>={start.isoformat()} created_at:<{end.isoformat()} status:any"
    )
    gql = """
    query ($q: String!, $cursor: String) {
      orders(first: 100, after: $cursor, query: $q, sortKey: CREATED_AT) {
        pageInfo { hasNextPage endCursor }
        edges {
          node {
            name
            tags
            currentTotalPriceSet { shopMoney { amount } }
            currentSubtotalPriceSet { shopMoney { amount } }
            lineItems(first: 50) { edges { node { quantity } } }
          }
        }
      }
    }
    """
    marketplace_tag_re = re.compile(
        r"(dsg|dick.?s|nordstrom|walmart|amazon|mirakl|edi)", re.I
    )
    cursor = None
    revenue = Decimal("0")
    orders = 0
    units = 0
    pages = 0
    while True:
        data = shopify_graphql(
            domain, token, gql, {"q": query_str, "cursor": cursor}
        )
        conn = (data.get("orders") or {})
        edges = conn.get("edges") or []
        for edge in edges:
            node = edge.get("node") or {}
            tags = " ".join(node.get("tags") or [])
            # Keep only D2C / direct — same spirit as main.py channel classifier.
            if marketplace_tag_re.search(tags):
                continue
            name = str(node.get("name") or "")
            # EDI / marketplace order names often lack '#' web pattern; keep web + numeric.
            amt = _money(
                ((node.get("currentTotalPriceSet") or {}).get("shopMoney") or {}).get("amount")
            )
            if amt == 0:
                continue
            revenue += amt
            orders += 1
            for li in ((node.get("lineItems") or {}).get("edges") or []):
                units += int(((li.get("node") or {}).get("quantity") or 0))
        pages += 1
        page = conn.get("pageInfo") or {}
        if not page.get("hasNextPage"):
            break
        cursor = page.get("endCursor")
        if pages > 200:
            log.warning("Shopify pagination safety stop at 200 pages")
            break

    ch = MonthChannel(
        revenue=revenue,
        units=units,
        orders=orders,
        days=(month_end(month) - start).days + 1,
        source="shopify_admin_api",
    )
    # Optional ad spend for CAC / MER / ROAS
    ad_path = ROOT / "data" / "shopify_ad_spend.json"
    if ad_path.exists():
        ads = json.loads(ad_path.read_text(encoding="utf-8"))
        prefix = month.strftime("%Y-%m")
        spend = Decimal("0")
        for k, v in (ads.items() if isinstance(ads, dict) else []):
            if str(k).startswith(prefix):
                spend += _money(v if not isinstance(v, dict) else v.get("spend"))
        if spend > 0:
            ch.ad_spend = float(spend)
            ch.cac = float((spend / Decimal(orders)).quantize(MONEY)) if orders else None
            ch.mer = float((spend / revenue * 100).quantize(MONEY)) if revenue else None
            ch.roas = float((revenue / spend).quantize(Decimal("0.01"))) if spend else None
    log.info(
        "Shopify month %s: revenue=%s orders=%s units=%s aov=%s pages=%s",
        month.strftime("%Y-%m"),
        ch.revenue,
        ch.orders,
        ch.units,
        ch.aov,
        pages,
    )
    return ch


# ---------------------------------------------------------------------------
# Sellerboard month rollup
# ---------------------------------------------------------------------------


def _load_sellerboard_csv_url(url: str):
    import pandas as pd

    req = Request(url, headers={"User-Agent": "WeathermanMonthlyDeck/1.0"})
    with urlopen(req, timeout=120) as resp:  # noqa: S310
        raw = resp.read()
    from io import BytesIO

    return pd.read_csv(BytesIO(raw))


def fetch_sellerboard_month(month: date) -> MonthChannel:
    """Aggregate Sellerboard daily automation CSV for the month."""
    url = os.getenv("SELLERBOARD_DAILY_URL", "").strip()
    cache_path = ROOT / "data" / "sellerboard_daily_cache.json"

    # Prefer live CSV when configured
    if url:
        try:
            from main import (  # type: ignore
                _first_matching_column,
                _parse_sellerboard_dates,
                _series_numeric,
                _sum_sellerboard_metric_group,
            )

            df = _load_sellerboard_csv_url(url)
            date_col = _first_matching_column(df, ("Date", "Day"))
            if date_col is None:
                raise RuntimeError("Sellerboard CSV missing Date column")
            parsed = _parse_sellerboard_dates(df[date_col])
            df = df.copy()
            df["_day"] = parsed.dt.date
            start, end = month.replace(day=1), month_end(month)
            filtered = df[(df["_day"] >= start) & (df["_day"] <= end)]
            revenue, _ = _sum_sellerboard_metric_group(
                filtered, prefix="sales",
                exact_fallbacks=("Ordered Product Sales", "Gross Sales", "Sales", "Revenue"),
            )
            units_dec, _ = _sum_sellerboard_metric_group(
                filtered, prefix="units",
                exact_fallbacks=("Units Ordered", "Units Sold", "Units"),
            )
            orders_col = _first_matching_column(filtered, ("Orders", "Order Count"))
            acos_col = _first_matching_column(filtered, ("Real ACOS", "ACOS", "ACoS"))
            net_col = _first_matching_column(filtered, ("NetProfit", "Net Profit", "Profit"))
            spend_cols = [
                c
                for c in filtered.columns
                if any(
                    x in str(c).lower()
                    for x in ("sponsored", "google ads", "facebook ads", "ppc")
                )
                and "sales" not in str(c).lower()
                and "units" not in str(c).lower()
            ]
            orders = int(_series_numeric(filtered, orders_col).sum()) if orders_col else int(units_dec)
            acos = None
            if acos_col is not None and revenue > 0:
                weights = _series_numeric(filtered, _first_matching_column(filtered, ("SalesOrganic", "Sales")) or acos_col)
                series = _series_numeric(filtered, acos_col)
                if not series.empty and not weights.empty:
                    acos = float((series * weights).sum() / weights.sum())
            ad_spend = float(sum(float(_series_numeric(filtered, c).sum()) for c in spend_cols)) if spend_cols else None
            net = float(_series_numeric(filtered, net_col).sum()) if net_col else None
            roas = (100.0 / acos) if acos and acos > 0 else None
            if ad_spend and float(revenue) > 0 and not roas:
                roas = float(revenue) / ad_spend
            tacos = (ad_spend / float(revenue) * 100) if ad_spend and revenue else (acos if acos else None)
            ch = MonthChannel(
                revenue=_money(revenue),
                units=int(units_dec),
                orders=orders,
                days=int(filtered["_day"].nunique()) if len(filtered) else 0,
                acos=acos,
                tacos=tacos,
                roas=roas,
                net_profit=net,
                ad_spend=ad_spend,
                source="sellerboard_csv",
            )
            log.info(
                "Sellerboard month %s: rev=%s acos=%s roas=%s net=%s days=%s",
                month.strftime("%Y-%m"),
                ch.revenue,
                ch.acos,
                ch.roas,
                ch.net_profit,
                ch.days,
            )
            return ch
        except Exception as exc:  # noqa: BLE001
            log.warning("Live Sellerboard CSV failed (%s); trying cache", exc)

    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        days = payload.get("days") or {}
        prefix = month.strftime("%Y-%m")
        rows = [v for k, v in days.items() if str(k).startswith(prefix)]
        if rows:
            rev = sum(float(r.get("revenue") or 0) for r in rows)
            orders = sum(int(r.get("orders") or 0) for r in rows)
            w = sum(float(r.get("revenue") or 0) for r in rows if r.get("acos") is not None)
            acos = (
                sum(float(r["acos"]) * float(r["revenue"]) for r in rows if r.get("acos") is not None) / w
                if w
                else None
            )
            roas = (100.0 / acos) if acos and acos > 0 else None
            ch = MonthChannel(
                revenue=_money(rev),
                units=sum(int(r.get("units") or r.get("orders") or 0) for r in rows),
                orders=orders,
                days=len(rows),
                acos=acos,
                tacos=acos,  # without spend breakdown, TACOS ≈ Real ACOS proxy
                roas=roas,
                source="sellerboard_cache",
            )
            log.info(
                "Sellerboard cache %s: rev=%s acos=%s roas=%s days=%s",
                month.strftime("%Y-%m"),
                ch.revenue,
                ch.acos,
                ch.roas,
                ch.days,
            )
            return ch

    log.warning("No Sellerboard data for %s", month.strftime("%Y-%m"))
    return MonthChannel(source="unavailable")


# ---------------------------------------------------------------------------
# Archive fallback + YTD series
# ---------------------------------------------------------------------------


def archive_channel_month(archive_path: Path, month: date, key: str) -> MonthChannel:
    payload = json.loads(archive_path.read_text(encoding="utf-8"))
    days = payload.get("days") or {}
    prefix = month.strftime("%Y-%m")
    ch = MonthChannel(source="daily_archive")
    for day_key in sorted(days):
        if not str(day_key).startswith(prefix):
            continue
        row = ((days[day_key] or {}).get("platforms") or {}).get(key) or {}
        if not row.get("available"):
            continue
        ch.revenue += _money(row.get("revenue"))
        ch.units += int(row.get("units") or 0)
        ch.orders += int(row.get("orders") or 0)
        ch.days += 1
    return ch


def build_ytd_series(archive_path: Path, year: int) -> dict[str, dict[str, float]]:
    """Return {YYYY-MM: {amazon, shopify_direct, total, ...}} for chart tables."""
    payload = json.loads(archive_path.read_text(encoding="utf-8"))
    days = payload.get("days") or {}
    out: dict[str, dict[str, float]] = {}
    for day_key, day in days.items():
        if not str(day_key).startswith(f"{year}-"):
            continue
        ym = str(day_key)[:7]
        bucket = out.setdefault(
            ym,
            {
                "amazon": 0.0,
                "shopify_direct": 0.0,
                "walmart": 0.0,
                "nordstrom": 0.0,
                "dsg": 0.0,
                "total": 0.0,
            },
        )
        for key in ("amazon", "shopify_direct", "walmart", "nordstrom", "dsg"):
            row = ((day or {}).get("platforms") or {}).get(key) or {}
            if row.get("available"):
                rev = float(row.get("revenue") or 0)
                bucket[key] += rev
                bucket["total"] += rev
    return dict(sorted(out.items()))


def merge_live_bundle(
    *,
    month: date,
    archive_path: Path,
    use_shopify: bool = True,
    use_sellerboard: bool = True,
) -> LiveMonthBundle:
    bundle = LiveMonthBundle(month=month)
    bundle.ytd_by_month = build_ytd_series(archive_path, month.year)

    # Retailers + baseline from archive
    for key, attr in (
        ("walmart", "walmart"),
        ("nordstrom", "nordstrom"),
        ("dsg", "dsg"),
        ("shopify_direct", "shopify"),
        ("amazon", "amazon"),
    ):
        setattr(bundle, attr, archive_channel_month(archive_path, month, key))

    if use_shopify:
        try:
            live = fetch_shopify_month(month)
            if live.revenue > 0:
                bundle.shopify = live
                bundle.notes.append(f"Shopify live API: {live.revenue} / {live.orders} orders")
                cache_path = ROOT / "data" / "shopify_month_cache.json"
                try:
                    cache = {}
                    if cache_path.exists():
                        cache = json.loads(cache_path.read_text(encoding="utf-8"))
                    cache[month.strftime("%Y-%m")] = {
                        "revenue": float(live.revenue),
                        "orders": live.orders,
                        "units": live.units,
                        "aov": float(live.aov or 0),
                    }
                    cache_path.write_text(json.dumps(cache, indent=2), encoding="utf-8")
                except Exception as cache_exc:  # noqa: BLE001
                    log.debug("Shopify cache write skipped: %s", cache_exc)
        except Exception as exc:  # noqa: BLE001
            bundle.notes.append(f"Shopify live failed: {exc}")
            log.warning("Shopify live fetch failed: %s", exc)
            cache_path = ROOT / "data" / "shopify_month_cache.json"
            if cache_path.exists():
                try:
                    block = json.loads(cache_path.read_text(encoding="utf-8")).get(
                        month.strftime("%Y-%m")
                    ) or {}
                    if float(block.get("revenue") or 0) > 0:
                        bundle.shopify = MonthChannel(
                            revenue=_money(block.get("revenue")),
                            orders=int(block.get("orders") or 0),
                            units=int(block.get("units") or 0),
                            days=max(1, bundle.shopify.days),
                            source="shopify_cache",
                        )
                        bundle.notes.append(
                            f"Shopify cache: {bundle.shopify.revenue} / {bundle.shopify.orders} orders"
                        )
                except Exception as cache_exc:  # noqa: BLE001
                    log.warning("Shopify cache read failed: %s", cache_exc)

    if use_sellerboard:
        try:
            live = fetch_sellerboard_month(month)
            if live.revenue > 0 or live.acos is not None:
                # Keep archive units if sellerboard units missing
                if live.units <= 0:
                    live.units = bundle.amazon.units
                if live.orders <= 0:
                    live.orders = bundle.amazon.orders
                if live.revenue <= 0:
                    live.revenue = bundle.amazon.revenue
                bundle.amazon = live
                bundle.notes.append(
                    f"Sellerboard ({live.source}): acos={live.acos} roas={live.roas}"
                )
        except Exception as exc:  # noqa: BLE001
            bundle.notes.append(f"Sellerboard failed: {exc}")
            log.warning("Sellerboard fetch failed: %s", exc)

    # Enrich YTD with live month totals
    ym = month.strftime("%Y-%m")
    bundle.ytd_by_month[ym] = {
        "amazon": float(bundle.amazon.revenue),
        "shopify_direct": float(bundle.shopify.revenue),
        "walmart": float(bundle.walmart.revenue),
        "nordstrom": float(bundle.nordstrom.revenue),
        "dsg": float(bundle.dsg.revenue),
        "total": float(bundle.total_revenue),
    }
    return bundle


def money_whole(value: Decimal | float | int | None) -> str:
    amount = Decimal(str(value or 0)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return f"${amount:,.0f}"


def money_compact_k(value: Decimal | float | int | None) -> str:
    """Short prior-period label to avoid card overflow (e.g. Sep $406K)."""
    amount = float(value or 0)
    if abs(amount) >= 1000:
        return f"${amount/1000:.1f}K".replace(".0K", "K")
    return f"${amount:,.0f}"


def executive_amazon_copy(cur: MonthChannel, prior: MonthChannel, month_label: str) -> tuple[str, str, str]:
    mom = None
    if prior.revenue:
        delta = (float(cur.revenue) - float(prior.revenue)) / float(prior.revenue) * 100
        mom = f"{'▲' if delta >= 0 else '▼'} {abs(delta):.1f}%"
    wins = (
        f"{month_label} Amazon MTD {money_whole(cur.revenue)} "
        f"({cur.orders:,} orders, AOV {money_whole(cur.aov or 0)}"
        + (f", {mom} vs prior month" if mom else "")
        + ")."
    )
    if cur.roas is not None and cur.acos is not None:
        risks = (
            f"Paid efficiency: ROAS {cur.roas:.2f}, Real ACOS {cur.acos:.1f}%."
            + (
                f" Prior ROAS {prior.roas:.2f}."
                if prior.roas is not None
                else ""
            )
        )
    else:
        risks = "Paid efficiency metrics still catching up for the month; watch ACOS on PPC."
    decisions = (
        "Defend Amazon volume while MTD ramps; rebalance inefficient ASIN/keyword spend "
        "and protect Travel/Trek stock cover."
    )
    return wins, risks, decisions


# ---------------------------------------------------------------------------
# Triple Whale — summary page + pixel channel mix (slides 14–15)
# ---------------------------------------------------------------------------

TW_API_BASE = "https://api.triplewhale.com/api/v2"
DEFAULT_TW_SHOP = "4a3474-24.myshopify.com"
DEFAULT_PTP_TARGETS = ROOT / "data" / "ptp_targets.json"
DEFAULT_TW_EXPORT = ROOT / "data" / "triple_whale_export.json"
DEFAULT_TW_CACHE_DIR = ROOT / "data" / "triple_whale_cache"


@dataclass
class PtpRow:
    channel: str
    planned: float | None
    actual: float | None
    roas: float | None = None

    @property
    def delta_pct(self) -> float | None:
        if self.planned is None or self.actual is None or self.planned == 0:
            return None
        return (self.actual - self.planned) / self.planned * 100.0


@dataclass
class ChannelMixRow:
    channel: str
    prior_revenue: float
    current_revenue: float

    @property
    def prior_share(self) -> float:
        return 0.0

    @property
    def current_share(self) -> float:
        return 0.0


@dataclass
class TripleWhaleMonth:
    month: date
    start: date
    end: date
    shop_id: str
    source: str
    summary: dict[str, float] = field(default_factory=dict)
    channels: dict[str, dict[str, float]] = field(default_factory=dict)  # channel -> spend/revenue
    prior_channels: dict[str, dict[str, float]] = field(default_factory=dict)
    revenue_rows: list[PtpRow] = field(default_factory=list)
    spend_rows: list[PtpRow] = field(default_factory=list)
    mix_rows: list[ChannelMixRow] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def blended_roas(self) -> float | None:
        return self.summary.get("roas")

    @property
    def blended_ads(self) -> float | None:
        return self.summary.get("blendedAds")

    @property
    def cac(self) -> float | None:
        return self.summary.get("newCustomersCpa")

    @property
    def mer(self) -> float | None:
        return self.summary.get("mer")

    @property
    def repeat_rate(self) -> float | None:
        return self.summary.get("oldCustomersPercent")

    @property
    def cvr(self) -> float | None:
        return self.summary.get("pixelConversionRate")

    @property
    def order_revenue(self) -> float | None:
        return self.summary.get("sales")


def _load_dotenv_file(path: Path) -> None:
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def load_triple_whale_credentials() -> dict[str, str]:
    """Resolve TW API key + shop from env, project .env, or sibling Weather Tool .env."""
    for candidate in (
        ROOT / ".env",
        Path("/Extra/Yandex.Disk/Weatherman/Projects/WT2/Weather Tool/.env"),
        Path("/Extra/Yandex.Disk/Weatherman/Projects/Weather Tool/_github_repo/.env"),
    ):
        _load_dotenv_file(candidate)

    api_key = os.getenv("TRIPLE_WHALE_API_KEY", "").strip()
    shop = (
        os.getenv("TRIPLE_WHALE_SHOP_ID", "").strip()
        or os.getenv("TRIPLE_WHALE_SHOP_DOMAIN", "").strip()
        or DEFAULT_TW_SHOP
    )
    if not api_key:
        # Last-resort local note file used by the team (not committed).
        note = Path("/Extra/Yandex.Disk/Weatherman/triplewhaleapi.txt")
        if note.exists():
            for line in note.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if re.fullmatch(
                    r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
                    line,
                    flags=re.I,
                ):
                    api_key = line
                    break
    return {"api_key": api_key, "shop_id": shop}


def _tw_post(path: str, payload: dict[str, Any], api_key: str) -> Any:
    body = json.dumps(payload).encode("utf-8")
    req = Request(
        f"{TW_API_BASE}{path}",
        data=body,
        headers={
            "x-api-key": api_key,
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    with urlopen(req, timeout=120) as resp:  # noqa: S310 — fixed Triple Whale API
        return json.loads(resp.read().decode("utf-8"))


def _extract_sql_rows(body: Any) -> list[dict[str, Any]]:
    if isinstance(body, list):
        return body
    if not isinstance(body, dict):
        return []
    if body.get("success") is False:
        raise RuntimeError(f"Triple Whale SQL failed: {body.get('message', body)}")
    data = body.get("data")
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("rows", "results", "data"):
            nested = data.get(key)
            if isinstance(nested, list):
                return nested
    return []


def fetch_summary_metrics(
    *,
    api_key: str,
    shop_id: str,
    start: date,
    end: date,
) -> dict[str, float]:
    payload = {
        "shopDomain": shop_id,
        "period": {"start": start.isoformat(), "end": end.isoformat()},
    }
    # Some tenants accept shopId; retry with both keys if needed.
    try:
        data = _tw_post("/summary-page/get-data", payload, api_key)
    except Exception:  # noqa: BLE001
        data = _tw_post(
            "/summary-page/get-data",
            {
                "shopId": shop_id,
                "dateRange": {"start": start.isoformat(), "end": end.isoformat()},
            },
            api_key,
        )
    out: dict[str, float] = {}
    for metric in data.get("metrics") or []:
        mid = str(metric.get("id") or "")
        vals = metric.get("values") or {}
        cur = vals.get("current")
        if mid and cur is not None:
            try:
                out[mid] = float(cur)
            except (TypeError, ValueError):
                continue
    return out


def fetch_pixel_channels(
    *,
    api_key: str,
    shop_id: str,
    start: date,
    end: date,
) -> dict[str, dict[str, float]]:
    query = f"""
SELECT channel, sum(spend) AS spend, sum(order_revenue) AS revenue
FROM pixel_joined_table
WHERE event_date BETWEEN '{start.isoformat()}' AND '{end.isoformat()}'
GROUP BY channel
ORDER BY revenue DESC
""".strip()
    body = _tw_post(
        "/orcabase/api/sql",
        {
            "shopId": shop_id,
            "query": query,
            "currency": "USD",
            "period": {
                "startDate": start.isoformat(),
                "endDate": end.isoformat(),
            },
        },
        api_key,
    )
    out: dict[str, dict[str, float]] = {}
    for row in _extract_sql_rows(body):
        channel = str(row.get("channel") or "").strip()
        if not channel:
            continue
        out[channel] = {
            "spend": float(row.get("spend") or 0),
            "revenue": float(row.get("revenue") or 0),
        }
    return out


def _channel_sum(channels: dict[str, dict[str, float]], names: tuple[str, ...], field: str) -> float:
    total = 0.0
    for name in names:
        total += float((channels.get(name) or {}).get(field) or 0)
    return total


def load_ptp_targets(path: Path, month: date) -> dict[str, Any]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    key = month.strftime("%Y-%m")
    block = payload.get(key) or payload.get("default") or {}
    return block if isinstance(block, dict) else {}


def load_triple_whale_export(path: Path, month: date) -> TripleWhaleMonth | None:
    """Offline / MCP-exported JSON fallback.

    Expected shape::
        {
          "2026-10": {
            "summary": {"facebookAds": ..., "sales": ...},
            "channels": {"facebook-ads": {"spend": ..., "revenue": ...}},
            "prior_channels": {...}
          }
        }
    """
    if not path.exists():
        return None
    payload = json.loads(path.read_text(encoding="utf-8"))
    block = payload.get(month.strftime("%Y-%m")) or payload.get("month") or {}
    if not isinstance(block, dict) or not block.get("summary"):
        return None
    start = month
    end = min(month_end(month), date.today())
    tw = TripleWhaleMonth(
        month=month,
        start=start,
        end=end,
        shop_id=str(block.get("shop_id") or DEFAULT_TW_SHOP),
        source="export",
        summary={k: float(v) for k, v in (block.get("summary") or {}).items()},
        channels={
            str(k): {"spend": float((v or {}).get("spend") or 0), "revenue": float((v or {}).get("revenue") or 0)}
            for k, v in (block.get("channels") or {}).items()
        },
        prior_channels={
            str(k): {"spend": float((v or {}).get("spend") or 0), "revenue": float((v or {}).get("revenue") or 0)}
            for k, v in (block.get("prior_channels") or {}).items()
        },
    )
    tw.notes.append(f"Loaded Triple Whale export {path}")
    return tw


def _build_ptp_rows(
    summary: dict[str, float],
    channels: dict[str, dict[str, float]],
    targets: dict[str, Any],
) -> tuple[list[PtpRow], list[PtpRow]]:
    """Map Triple Whale metrics onto the PTP slide channel labels."""
    rev_targets = (targets.get("revenue") or {}) if isinstance(targets.get("revenue"), dict) else {}
    spend_targets = (targets.get("spend") or {}) if isinstance(targets.get("spend"), dict) else {}

    # Prefer Summary-page platform-reported conversion for paid (matches Sep MBR),
    # fall back to pixel_joined channel revenue.
    paid_social_rev = summary.get("facebookConversionValue")
    if paid_social_rev is None:
        paid_social_rev = _channel_sum(channels, ("facebook-ads",), "revenue")
    paid_search_rev = summary.get("googleConversionValue")
    if paid_search_rev is None:
        paid_search_rev = _channel_sum(channels, ("google-ads", "googlepaid"), "revenue")
    email_sms_rev = _channel_sum(channels, ("klaviyo", "attentive"), "revenue")
    if not email_sms_rev:
        email_sms_rev = float(summary.get("klaviyoPlacedOrderSales") or 0)
    organic_direct_rev = _channel_sum(
        channels,
        ("Direct", "organic", "organic_and_social", "Non-attributed", "shop_app"),
        "revenue",
    )
    affiliates_rev = summary.get("influencerConversionValue")
    if affiliates_rev is None:
        affiliates_rev = _channel_sum(channels, ("influencers",), "revenue")

    paid_social_spend = float(summary.get("facebookAds") or _channel_sum(channels, ("facebook-ads",), "spend"))
    paid_search_spend = float(summary.get("googleAds") or _channel_sum(channels, ("google-ads",), "spend"))

    revenue_rows = [
        PtpRow("Paid Social", _opt_float(rev_targets.get("Paid Social")), float(paid_social_rev or 0),
               summary.get("facebookRoas")),
        PtpRow("Paid Search", _opt_float(rev_targets.get("Paid Search")), float(paid_search_rev or 0),
               summary.get("googleRoas")),
        PtpRow("Email/SMS", _opt_float(rev_targets.get("Email/SMS")), float(email_sms_rev or 0)),
        PtpRow("Organic/Direct", _opt_float(rev_targets.get("Organic/Direct")), float(organic_direct_rev or 0)),
        PtpRow("Affiliates", _opt_float(rev_targets.get("Affiliates")), float(affiliates_rev or 0)),
    ]
    spend_rows = [
        PtpRow(
            "Paid Social",
            _opt_float(spend_targets.get("Paid Social")),
            paid_social_spend,
            summary.get("facebookRoas"),
        ),
        PtpRow(
            "Paid Search",
            _opt_float(spend_targets.get("Paid Search")),
            paid_search_spend,
            summary.get("googleRoas"),
        ),
    ]
    return revenue_rows, spend_rows


def _opt_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _build_mix_rows(
    current: dict[str, dict[str, float]],
    prior: dict[str, dict[str, float]],
) -> list[ChannelMixRow]:
    """Collapse pixel channels into the Channel Mix slide buckets."""
    buckets = (
        ("Paid Social", ("facebook-ads",)),
        ("Paid Search", ("google-ads", "googlepaid")),
        ("Email/SMS", ("klaviyo", "attentive")),
        ("Organic/Direct", ("Direct", "organic", "organic_and_social", "Non-attributed", "shop_app")),
        ("Affiliates", ("influencers",)),
        ("Other", ()),  # filled below
    )
    used: set[str] = set()
    for _, names in buckets:
        used.update(names)

    rows: list[ChannelMixRow] = []
    for label, names in buckets:
        if label == "Other":
            cur = sum(
                float(v.get("revenue") or 0)
                for k, v in current.items()
                if k not in used and k.lower() not in {"amazon", "excluded"}
            )
            prv = sum(
                float(v.get("revenue") or 0)
                for k, v in prior.items()
                if k not in used and k.lower() not in {"amazon", "excluded"}
            )
        else:
            cur = _channel_sum(current, names, "revenue")
            prv = _channel_sum(prior, names, "revenue")
        if cur or prv:
            rows.append(ChannelMixRow(label, prv, cur))
    rows.sort(key=lambda r: r.current_revenue, reverse=True)
    return rows


def fetch_triple_whale_month(
    month: date,
    *,
    export_path: Path | None = None,
    targets_path: Path | None = None,
    end_override: date | None = None,
    use_api: bool = True,
) -> TripleWhaleMonth | None:
    """Load Triple Whale attribution for ``month`` (API, then export, then cache)."""
    export_path = export_path or DEFAULT_TW_EXPORT
    targets_path = targets_path or DEFAULT_PTP_TARGETS
    start = month
    end = end_override or min(month_end(month), date.today())
    prior_start = prior_month_start(month)
    prior_end = month_end(prior_start)

    tw = load_triple_whale_export(export_path, month) if export_path else None
    creds = load_triple_whale_credentials()
    api_key = creds["api_key"]
    shop_id = creds["shop_id"]

    if tw is None and use_api and api_key:
        try:
            summary = fetch_summary_metrics(
                api_key=api_key, shop_id=shop_id, start=start, end=end
            )
            channels = fetch_pixel_channels(
                api_key=api_key, shop_id=shop_id, start=start, end=end
            )
            prior_channels = fetch_pixel_channels(
                api_key=api_key, shop_id=shop_id, start=prior_start, end=prior_end
            )
            tw = TripleWhaleMonth(
                month=month,
                start=start,
                end=end,
                shop_id=shop_id,
                source="api",
                summary=summary,
                channels=channels,
                prior_channels=prior_channels,
            )
            tw.notes.append(
                f"Triple Whale API {start}→{end}: sales={summary.get('sales')} "
                f"blendedAds={summary.get('blendedAds')} roas={summary.get('roas')}"
            )
            # Persist cache for offline re-runs
            DEFAULT_TW_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            cache_path = DEFAULT_TW_CACHE_DIR / f"{month.strftime('%Y-%m')}.json"
            cache_path.write_text(
                json.dumps(
                    {
                        "shop_id": shop_id,
                        "start": start.isoformat(),
                        "end": end.isoformat(),
                        "summary": summary,
                        "channels": channels,
                        "prior_channels": prior_channels,
                    },
                    indent=2,
                ),
                encoding="utf-8",
            )
        except Exception as exc:  # noqa: BLE001
            log.warning("Triple Whale API failed: %s", exc)

    if tw is None:
        cache_path = DEFAULT_TW_CACHE_DIR / f"{month.strftime('%Y-%m')}.json"
        if cache_path.exists():
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            tw = TripleWhaleMonth(
                month=month,
                start=date.fromisoformat(payload.get("start") or start.isoformat()),
                end=date.fromisoformat(payload.get("end") or end.isoformat()),
                shop_id=str(payload.get("shop_id") or shop_id or DEFAULT_TW_SHOP),
                source="cache",
                summary={k: float(v) for k, v in (payload.get("summary") or {}).items()},
                channels={
                    str(k): {
                        "spend": float((v or {}).get("spend") or 0),
                        "revenue": float((v or {}).get("revenue") or 0),
                    }
                    for k, v in (payload.get("channels") or {}).items()
                },
                prior_channels={
                    str(k): {
                        "spend": float((v or {}).get("spend") or 0),
                        "revenue": float((v or {}).get("revenue") or 0),
                    }
                    for k, v in (payload.get("prior_channels") or {}).items()
                },
            )
            tw.notes.append(f"Loaded Triple Whale cache {cache_path}")

    if tw is None:
        return None

    targets = load_ptp_targets(targets_path, month)
    # If planned targets missing, use prior-month TW actuals so Δ = MoM.
    if not targets.get("revenue") and tw.prior_channels:
        prior_summary_path = DEFAULT_TW_CACHE_DIR / f"{prior_start.strftime('%Y-%m')}.json"
        prior_summary: dict[str, float] = {}
        if prior_summary_path.exists():
            prior_summary = {
                k: float(v)
                for k, v in (
                    json.loads(prior_summary_path.read_text(encoding="utf-8")).get("summary") or {}
                ).items()
            }
        if not prior_summary and api_key and use_api:
            try:
                prior_summary = fetch_summary_metrics(
                    api_key=api_key, shop_id=shop_id, start=prior_start, end=prior_end
                )
            except Exception as exc:  # noqa: BLE001
                log.warning("Prior-month TW summary failed: %s", exc)
        prior_rev, prior_spend = _build_ptp_rows(prior_summary, tw.prior_channels, {})
        rev_map = {r.channel: r.actual for r in prior_rev}
        # Organic/Direct often lands at 0 via summary fields — use pixel mix prior.
        if not rev_map.get("Organic/Direct"):
            rev_map["Organic/Direct"] = _channel_sum(
                tw.prior_channels,
                ("Direct", "organic", "organic_and_social", "Non-attributed", "shop_app"),
                "revenue",
            )
        targets = {
            "revenue": rev_map,
            "spend": {r.channel: r.actual for r in prior_spend},
            "plan_source": "prior_month_triple_whale",
        }
        tw.notes.append("PTP planned figures = prior-month Triple Whale actuals (MoM Δ)")

    tw.revenue_rows, tw.spend_rows = _build_ptp_rows(tw.summary, tw.channels, targets)
    # If targets provided planned but some rows still None, leave planned blank → Δ n/a
    tw.mix_rows = _build_mix_rows(tw.channels, tw.prior_channels)
    if targets.get("plan_source"):
        tw.notes.append(f"PTP plan_source={targets.get('plan_source')}")
    elif targets:
        tw.notes.append(f"PTP plan_source={targets_path.name}")
    return tw
