"""Repo-aware dashboard brand + GitHub Pages base URL.

Detection order:
1. ``GITHUB_REPOSITORY`` (set automatically in Actions) — ``WeathermanInc/…`` → WEATHERMAN,
   ``mg22mex/…`` → MARCO.
2. Optional overrides: ``DASHBOARD_BRAND``, ``DASHBOARD_PUBLIC_URL``.
3. Local / unknown context defaults to MARCO + the mg22mex Pages site.
"""
from __future__ import annotations

import os
from dataclasses import dataclass

MG22MEX_PAGES = "https://mg22mex.github.io/Daily-revenue-by-sales-platform"
WEATHERMAN_PAGES = "https://weathermaninc.github.io/Daily-revenue-by-sales-platform"


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
    """Resolve brand name and hosted Pages root for the current deployment."""
    repo = os.environ.get("GITHUB_REPOSITORY", "").strip()
    owner = repo.split("/", 1)[0].lower() if repo else ""

    if owner == "weathermaninc":
        brand = "WEATHERMAN"
        default_base = WEATHERMAN_PAGES
    elif owner == "mg22mex":
        brand = "MARCO"
        default_base = MG22MEX_PAGES
    else:
        # Local runs and unknown hosts → MARCO / mg22mex Pages.
        brand = "MARCO"
        default_base = MG22MEX_PAGES

    brand_override = os.environ.get("DASHBOARD_BRAND", "").strip()
    if brand_override:
        brand = brand_override.upper()

    secret_url = os.environ.get("DASHBOARD_PUBLIC_URL", "").strip()
    if owner in {"weathermaninc", "mg22mex"}:
        # In Actions, prefer the Pages host that matches the running repo so a
        # stale DASHBOARD_PUBLIC_URL secret cannot cross-link forks.
        if secret_url and owner in secret_url.lower():
            pages_base = _normalize_pages_base(secret_url, default_base)
        else:
            pages_base = default_base
    else:
        pages_base = _normalize_pages_base(secret_url, default_base)

    return SiteContext(
        brand=brand,
        pages_base=pages_base,
        repository=repo,
        owner=owner or "local",
    )


def dated_dashboard_url(report_day_iso: str) -> str:
    """Return ``{pages_base}/{YYYY-MM-DD}.html`` for Brevo CTAs / deep links."""
    ctx = resolve_site_context()
    return f"{ctx.pages_base}/{report_day_iso}.html"
