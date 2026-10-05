"""WEATHERMAN dashboard brand + GitHub Pages base URL.

Brand is always **WEATHERMAN**. Pages host defaults to the WeathermanInc site;
``DASHBOARD_PUBLIC_URL`` may override. When running in Actions,
``GITHUB_REPOSITORY`` selects the matching ``*.github.io`` host so each fork's
CTAs stay self-referencing (brand header remains WEATHERMAN either way).
"""
from __future__ import annotations

import os
from dataclasses import dataclass

WEATHERMAN_BRAND = "WEATHERMAN"
WEATHERMAN_PAGES = "https://weathermaninc.github.io/Daily-revenue-by-sales-platform"
# Mirror Pages host when the pipeline runs on the mg22mex fork (brand unchanged).
MG22MEX_PAGES = "https://mg22mex.github.io/Daily-revenue-by-sales-platform"


@dataclass(frozen=True)
class SiteContext:
    brand: str
    pages_base: str
    repository: str
    owner: str


def _normalize_pages_base(raw: str, fallback: str) -> str:
    base = (raw or "").strip() or fallback
    base = base.rstrip("/")
    if base.lower().endswith(".html"):
        base = base.rsplit("/", 1)[0].rstrip("/")
    return base or fallback


def resolve_site_context() -> SiteContext:
    """Resolve hosted Pages root; brand is always WEATHERMAN."""
    repo = os.environ.get("GITHUB_REPOSITORY", "").strip()
    owner = repo.split("/", 1)[0].lower() if repo else ""

    if owner == "mg22mex":
        default_base = MG22MEX_PAGES
    else:
        # WeathermanInc Actions + local / unknown → Weatherman hosted dashboard.
        default_base = WEATHERMAN_PAGES

    secret_url = os.environ.get("DASHBOARD_PUBLIC_URL", "").strip()
    if secret_url:
        pages_base = _normalize_pages_base(secret_url, default_base)
    else:
        pages_base = default_base

    return SiteContext(
        brand=WEATHERMAN_BRAND,
        pages_base=pages_base,
        repository=repo,
        owner=owner or "local",
    )


def dated_dashboard_url(report_day_iso: str) -> str:
    """Return ``{pages_base}/{YYYY-MM-DD}.html`` for Brevo CTAs / deep links."""
    ctx = resolve_site_context()
    return f"{ctx.pages_base}/{report_day_iso}.html"
