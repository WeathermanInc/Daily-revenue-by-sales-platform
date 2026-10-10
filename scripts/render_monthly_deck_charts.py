#!/usr/bin/env python3
"""Render matplotlib chart PNGs for the Monthly Business Review Google Slides.

Outputs land in ``output/monthly_decks/charts/`` and are inserted into the
creative deck via Drive MCP ``insertSlidesLocalImage`` (or Slides API):

* Slides 12–13 — YTD **line-bar** (channel bars + all-channel line)
* Slide 22 — retailers metrics table (blue header)
* Slide 24 — potential OOS SKU cards

Example
-------
.. code-block:: bash

    .venv/bin/python scripts/render_monthly_deck_charts.py --month 2026-10
"""
from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from matplotlib.patches import FancyBboxPatch  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUT = ROOT / "output" / "monthly_decks" / "charts"


def render_ytd_linebar(
    out: Path,
    *,
    months: list[str],
    all_ch: list[float],
    shop: list[float],
    amz: list[float],
    title_note: str,
    filename: str,
) -> Path:
    x = np.arange(len(months))
    w = 0.22
    fig, ax = plt.subplots(figsize=(11.8, 5.4), dpi=150)
    ax.bar(x - w, shop, w, label="Shopify / D2C", color="#c45c26", zorder=2)
    ax.bar(x, amz, w, label="Amazon", color="#d4a017", zorder=2)
    ax.plot(x, all_ch, "o-", color="#1e3a5f", lw=3, ms=10, label="All Channel", zorder=3)
    for i, v in enumerate(all_ch):
        ax.annotate(
            f"${v:,.0f}" if i == len(all_ch) - 1 else f"${v / 1000:.0f}K",
            (i, v),
            textcoords="offset points",
            xytext=(0, 10),
            ha="center",
            fontsize=12 if i == len(all_ch) - 1 else 10,
            fontweight="bold",
            color="#1e3a5f",
        )
    ax.set_xticks(x)
    ax.set_xticklabels(months, fontsize=12)
    ax.set_ylabel("Revenue (USD)", fontsize=11)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"${v / 1000:.0f}K"))
    ax.legend(frameon=False, fontsize=10, loc="upper right")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_ylim(0, max(all_ch) * 1.2)
    ax.grid(axis="y", alpha=0.25, zorder=0)
    if title_note:
        ax.set_title(title_note, fontsize=11, color="#64748b", loc="left", pad=8)
    fig.tight_layout(pad=0.5)
    path = out / filename
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close()
    return path


def render_retailers_blue_table(out: Path, rows: list[list[str]], filename: str) -> Path:
    headers = ["Channel", "Revenue", "Units", "Orders", "AOV", "Prior (Sep)", "MoM"]
    fig, ax = plt.subplots(figsize=(13.2, 3.4), dpi=160)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.text(
        0.01,
        0.92,
        "Top Metrics Across DSG & Nordstrom — October 2026",
        fontsize=15,
        fontweight="bold",
        color="#1e3a5f",
        va="top",
    )
    col_w = [1.1, 1.2, 0.7, 0.7, 0.9, 1.2, 0.9]
    s = sum(col_w)
    col_w = [w / s for w in col_w]
    top = 0.78
    rh = 0.16
    x = 0.01
    for i, h in enumerate(headers):
        w = col_w[i] * 0.98
        ax.add_patch(
            FancyBboxPatch(
                (x, top - rh),
                w,
                rh,
                boxstyle="square,pad=0",
                facecolor="#1e3a5f",
                edgecolor="none",
            )
        )
        ax.text(x + 0.008, top - rh / 2, h, fontsize=10, fontweight="bold", color="white", va="center")
        x += w
    for r_i, row in enumerate(rows):
        y = top - (r_i + 1) * rh
        x = 0.01
        bg = "#e8f1f8" if r_i % 2 == 0 else "white"
        for i, cell in enumerate(row):
            w = col_w[i] * 0.98
            ax.add_patch(
                FancyBboxPatch(
                    (x, y - rh),
                    w,
                    rh,
                    boxstyle="square,pad=0",
                    facecolor=bg,
                    edgecolor="#c5d4e0",
                    lw=0.5,
                )
            )
            color = "#c0392b" if "▼" in cell else ("#1e7a46" if "▲" in cell else "#0f172a")
            ax.text(
                x + 0.008,
                y - rh / 2,
                cell,
                fontsize=11,
                color=color,
                va="center",
                fontweight="bold" if i == 0 else "normal",
            )
            x += w
    path = out / filename
    fig.savefig(path, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close()
    return path


def render_oos_pretty(out: Path, oos: list[tuple], filename: str) -> Path:
    fig, ax = plt.subplots(figsize=(13.2, 6.6), dpi=160)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.text(0.02, 0.95, "Potential OOS SKUs", fontsize=20, fontweight="bold", color="#1e3a5f", va="top")
    ax.text(
        0.02,
        0.90,
        "Ranked by days of cover · Available = AMZ sellable",
        fontsize=10,
        color="#64748b",
        va="top",
    )
    colors = {"Critical": "#c0392b", "Watch": "#d4a017", "Monitor": "#3b82f6", "OK": "#1e7a46"}
    bgc = {"Critical": "#fde8e8", "Watch": "#fff6e0", "Monitor": "#e8f1ff", "OK": "#e8f6ee"}
    for i, (name, days, stock, avail, status) in enumerate(oos):
        y = 0.78 - i * 0.12
        ax.add_patch(
            FancyBboxPatch(
                (0.02, y - 0.09),
                0.96,
                0.10,
                boxstyle="round,pad=0.008,rounding_size=0.02",
                facecolor=bgc[status],
                edgecolor=colors[status],
                lw=1.2,
            )
        )
        ax.text(0.04, y - 0.02, name, fontsize=13, fontweight="bold", color="#0f172a", va="top")
        ax.text(
            0.04,
            y - 0.055,
            f"~{days}d cover  ·  stock {stock}  ·  avail {avail}",
            fontsize=11,
            color="#334155",
            va="top",
        )
        ax.text(
            0.93,
            y - 0.035,
            status.upper(),
            fontsize=11,
            fontweight="bold",
            color=colors[status],
            va="top",
            ha="right",
        )
    path = out / filename
    fig.savefig(path, bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close()
    return path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--month", default="2026-10", help="YYYY-MM label for defaults")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)

    # October 2026 MTD defaults (Aug/Sep full months from daily_archive).
    months = ["Aug", "Sep", "Oct MTD"]
    all_ch = [514953.0, 537909.0, 218122.0]
    shop = [108140.0, 117295.0, 56891.0]
    amz = [381065.0, 405664.0, 154096.0]

    paths = [
        render_ytd_linebar(
            args.out,
            months=months,
            all_ch=all_ch,
            shop=shop,
            amz=amz,
            title_note="Bars = channel · Line = all-channel total",
            filename="ytd_linebar_wo_foh.png",
        ),
        render_ytd_linebar(
            args.out,
            months=months,
            all_ch=all_ch,
            shop=shop,
            amz=amz,
            title_note="FOH proxy · bars = channel · line = all-channel",
            filename="ytd_linebar_w_foh.png",
        ),
        render_retailers_blue_table(
            args.out,
            [
                ["DSG", "$364.30", "12", "9", "$40.48", "$2,566.01", "▼ 85.8%"],
                ["Nordstrom", "$6,405.05", "99", "79", "$81.08", "$5,975.61", "▲ 7.2%"],
                ["Walmart", "$364.98", "5", "5", "$73.00", "$6,409.00", "▼ 94.3%"],
            ],
            "table_retailers_blue.png",
        ),
        render_oos_pretty(
            args.out,
            [
                ("Travel Dusty Lavender", 1, 2, 0, "Critical"),
                ("Kid Whispy Blue", 6, 19, 0, "Critical"),
                ("Golf Red", 6, 16, 9, "Watch"),
                ("Kid FOH Blue 2025", 9, 26, 23, "Watch"),
                ("Kid Cosmic Rainbow", 11, 117, 111, "Monitor"),
                ("Trek 2026 America 250", 20, 160, 148, "OK"),
            ],
            "oos_pretty.png",
        ),
    ]
    for p in paths:
        print("wrote", p)


if __name__ == "__main__":
    main()
