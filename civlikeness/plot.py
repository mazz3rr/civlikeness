"""
Plot the Ch/H column of the Civ-likeness ledger.

Chassis  = P1 + P2 + P3 + P4 + P6a + P7   (max 30)  -> X axis
House    = P5 + P6b                       (max 10)  -> Y axis

Both axes use uniform, equal-length data units (House is NOT rescaled to
span the same pixel range as Chassis), so the plot is a true 1:1 grid
where 1 House point occupies the same physical length as 1 Chassis point.

Recomputing from the pillar columns avoids parsing the "Ch/H" text string;
if you prefer parsing, the equivalent line is given in a comment below.
"""
from __future__ import annotations

from pathlib import Path

import math

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.colors import LinearSegmentedColormap

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
XLSX_PATH = Path(__file__).parent.parent / "assets" / "ledger.xlsx"
SHEET = "ledger"
HEADER_ROW = 8  # 1-based; data starts here (row 7 is the header)
VAR_DIR = Path("var")
VAR_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_PNG = VAR_DIR / "ch_h_plot.png"

# Column layout of the ledger sheet (1-based, matches header in row 6/7):
#   A Game | B P1 Map | C P2 Expansion | D P3 Economy | E P4 Ladder
#   F P5 Peers (H) | G P6a Abstraction | H P6b Polity/Endings (H) | I P7 Time
#   J Ch/H | K Total | ...
COL_GAME = 1
# Chassis pillars (non-(H) pillars):
CHASSIS_COLS = [2, 3, 4, 5, 7, 9]   # P1, P2, P3, P4, P6a, P7
# House pillars (the ones marked "(H)"):
HOUSE_COLS = [6, 8]                # P5, P6b

# Short, recognizable labels for the dots. Anything not listed falls back to a
# heuristic shortener below, so adding new games never breaks the plot.
SHORT_NAMES = {
    "Civ4 (2005, BtS)":                                     "Civ4",
    "Civ6 (2016, GS)":                                      "Civ6",
    "Civ3 (2001, Conquests)":                               "Civ3",
    "Civ5 (2010, BNW)":                                     "Civ5",
    "Alpha Centauri (1999, SMAX)":                          "SMAC",
    "Call to Power II (2000, retail final)":                "CtP2",
    "Galactic Civilizations 2 (2006–08, Twilight)":         "GalCiv2",
    "Civ2 (1996)":                                          "Civ2",
    "Old World (2021, 7 DLC)":                              "OW",
    "Humankind (2021, TWR)":                                "HK",
    "Civ1 (1991)":                                          "Civ1",
    "Call to Power (1999, v1.2)":                           "CtP",
    "Civ7 (2025, current)":                                 "Civ7",
    "Endless Legend (2014, Symbiosis)":                     "EL",
    "Beyond Earth (2014, +RT)":                             "BE",
    "Stellaris (2016–current)":                             "Stellaris",
    "Galactic Civilizations 3 (2014–19, Retribution)":      "GalCiv3",
    "MOO2 (1996)":                                          "MOO2",
    "Galactic Civilizations 4 (2022–current, Supernova)":   "GalCiv4",
    "MOO1 (1993)":                                          "MOO1",
    "Master of Magic (1994, CoM)":                          "MoM",
    "Colonization (1994)":                                  "Col",
    "Age of Wonders 4 (2023–current)":                      "AoW4",
    "Anno 1800 (2019–2022)":                                "A1800",
    "Age of Wonders 3 (2014)":                              "AoW3",
    "Age of Wonders: Shadow Magic (2003)":                  "AoW:SM",
    "Dominions 6 (2024–current, 6.37)":                     "Dom6",
    "EU4 (2013–2025, final)":                               "EU4",
    "Total War: Warhammer III (2022, IE + 9.0)":            "TWW3",
    "CK3 (2020–2026)":                                      "CK3",
    "EVE Online (current)":                                 "EVE",
    "Age of Empires II: DE (2019–current)":                 "AoE2",
    "Heroes of Might & Magic III (1999, Complete)":         "HoMM3",
    "Conquest of Elysium 5 (2021–current)":                 "CoE5",
    "Offworld Trading Company (2016, final)":               "OTC",
    "StarCraft II (melee, vs-AI)":                          "SC2",
    "Warlords III (1997–98, + DLR)":                        "WL3",
    "Warcraft III: TFT (2003, melee vs-AI)":                "WC3",
    "SimCity 4 (2003, Deluxe+641)":                         "SC4",
    "Football Manager 26 (final)":                          "FM26",
    "XCOM 2 (2017, WotC)":                                  "XCOM2",
    "The Sims 4 (live)":                                    "Sims4",
    "Into the Breach (2018, AE)":                           "ItB",
    "Doom (1993, Ultimate)":                                "Doom",
}


def short_name(full: str) -> str:
    """Return a short label for a game, with a heuristic fallback.
    """
    if sn := SHORT_NAMES.get(full):
        return sn
    # Fallback: take everything before the first '(' and trim.
    base = full.split("(", 1)[0].strip() if "(" in full else full.strip()
    if len(base) <= 14:
        return base
    # Otherwise keep initials of the first few words.
    words = base.split()
    if len(words) >= 2:
        return "".join(w[0] for w in words[:4]) or base[:12]
    return base[:12]


def load_rows() -> list[tuple[str, float, float]]:
    """Return list of (short_name, chassis, house) recomputed from pillars.
    """
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb[SHEET]

    out: list[tuple[str, float, float]] = []
    for row in ws.iter_rows(min_row=HEADER_ROW, values_only=True):
        game = row[COL_GAME - 1]
        if game is None or str(game).strip() == "":
            continue
        vals = list(row)

        def col(idx):
            v = vals[idx - 1]
            return float(v) if v is not None else 0.0

        chassis = sum(col(c) for c in CHASSIS_COLS)   # /30
        house = sum(col(c) for c in HOUSE_COLS)       # /10
        out.append((short_name(str(game).strip()), chassis, house))
    return out


def main() -> None:
    print("Loading data...")
    data = load_rows()

    print("Preparing plot...")
    x = np.array([d[1] for d in data], dtype=float)   # Chassis 0..30
    y = np.array([d[2] for d in data], dtype=float)   # House   0..10

    # Color by total score (Chassis + House) on the nominal 0-40 scale.
    totals = x + y

    # Custom colormap: viridis extended past its yellow endpoint to red, so the
    # very top of the scale (a perfect 40) is red instead of stopping at yellow.
    viridis_rgb = plt.cm.viridis(np.linspace(0, 1, 200))[:, :3]
    yellow = plt.cm.viridis(1.0)[:3]
    red = np.array([1.0, 0.0, 0.0])
    n_ext = 60
    ext_rgb = np.array([np.linspace(yellow[i], red[i], n_ext) for i in range(3)]).T
    cmap = LinearSegmentedColormap.from_list(
        "viridis_red", np.vstack([viridis_rgb, ext_rgb[1:]])
    )

    fig, ax = plt.subplots(figsize=(16, 9))

    # Uniform units: equal aspect ratio so 1 House unit == 1 Chassis unit in
    # physical length. The axes' limits then define the visible region.
    sc = ax.scatter(
        x,
        y,
        c=totals,
        cmap=cmap,
        vmin=0,
        vmax=40,
        s=70,
        edgecolors="white",
        linewidths=0.6,
        zorder=3
    )

    # Initial axes limits — generous padding so the dense top-right cluster
    # can fan labels out into the empty space. adjustText with expand_axes=True
    # may expand these further if it needs more room.
    ax.set_xlim(-1, 37)
    ax.set_ylim(-0.5, 11.5)
    ax.set_aspect("equal", adjustable="box")   # <- the "uniform units" rule

    # Label every dot with adjustText, passing the scatter points themselves as
    # obstacles. The key parameters that make this work for the dense top-right
    # cluster (~20 games in a 4x2.5 area):
    #   - max_move=None: allow labels to displace arbitrarily far into the empty
    #     space above/right of the cluster (default is only 10 points!)
    #   - min_arrow_len: force a minimum leader-line length so labels are always
    #     clearly separated from their dots
    #   - objects=scatter: labels must avoid the dots themselves, not just each
    #     other, so they don't sit on top of nearby points
    try:
        from adjustText import adjust_text  # type: ignore
        # White halo behind each label so grid lines and the reference
        # diagonal never cut through the text, even when a label sits near a
        # dot or another label.
        print("Adjusting labels...")
        _LABEL_BBOX = dict(
            boxstyle="round,pad=0.15", facecolor="white", edgecolor="none", alpha=0.8)
        texts = [
            ax.text(xi, yi, n, fontsize=7, zorder=6, bbox=_LABEL_BBOX)
            for n, xi, yi in data
        ]
        adjust_text(
            texts,
            objects=sc,
            # Allow unlimited displacement into the expanded plot margins.
            max_move=None,
            # Stronger repulsion so the dense cluster actually separates.
            # force_text=(1.2, 1.5),
            force_text=(1.5, 1.8),
            force_static=(0.7, 0.9),
            # Pull labels back toward their anchor dots so they don't drift
            # far away (e.g. "FM26" floating loose, "Old World" hanging off
            # the end of its leader line). Applied once a label strays
            # further than pull_threshold points from its dot.
            force_pull=(0.15, 0.15),
            pull_threshold=8,
            # Expand bounding boxes more aggressively so labels get real space.
            expand=(1.4, 1.7),
            # Let adjustText expand the axes limits if it needs more room.
            expand_axes=True,
            # Minimum leader-line length — ensures each label is pulled far
            # enough from its dot to be traceable.
            min_arrow_len=15,
            iter_lim=5000,
            # time_lim removed — let iter_lim drive convergence.
            # Let labels move freely in all directions into the empty space.
            only_move={
                "text": "xy",
                "static": "xy",
                "explode": "xy",
                "pull": "xy",
            },
            # Don't force labels inside the tick range — they can use the
            # padded margin space we reserved.
            ensure_inside_axes=False,
            arrowprops=dict(
                arrowstyle="-",
                color="0.5",
                lw=0.4,
                shrinkA=3,
                shrinkB=3,
            ),
        )
    except ImportError:
        for n, xi, yi in data:
            ax.annotate(
                n, (xi, yi), textcoords="offset points", xytext=(7, 5), fontsize=8, zorder=5)

    # Axes limits are set above (before adjust_text) so expand_axes can
    # push them further if needed. The aspect ratio is already set to equal.
    ax.set_xlabel("Chassis  (P1+P2+P3+P4+P6a+P7)   / 30")
    ax.set_ylabel("House  (P5+P6b)   / 10")
    ax.set_title("Civ-likeness: Chassis vs. House")

    # Light reference grid at every integer point.
    ax.set_xticks(np.arange(0, 31, 2))
    ax.set_yticks(np.arange(0, 11, 1))
    ax.grid(True, which="both", linestyle=":", linewidth=0.5, color="0.8", zorder=0)

    cbar = fig.colorbar(sc, ax=ax, shrink=0.7, pad=0.02)
    cbar.set_label("Total Ch+H")
    # Nominal 0-40 scale with 5-unit ticks (8 segments).
    cbar.set_ticks(np.arange(0, 41, 5))

    # Reference line: the "3:1 ratio" diagonal from (0,0) to (30,10), i.e. a
    # House share of 1/4 of the total. Games above it are House-heavy
    # (relationship/ends-leaning); games below are Chassis-heavy (operational).
    # Dashed, semi-transparent, drawn under the points so it never crowds them.
    ax.plot(
        [0, 30],
        [0, 10],
        linestyle=(0, (5, 4)),
        color="#444444",
        alpha=0.6,
        linewidth=1.3,
        zorder=1
    )
    ax.text(
        1.2,
        0.55,
        "3:1 ratio line  (House = ¼ of total)",
        fontsize=8,
        color="#555555",
        rotation=math.degrees(math.atan2(10, 30)),
        zorder=1
    )

    fig.tight_layout()

    fig.savefig(OUTPUT_PNG, dpi=150)
    print(f"Plotted {len(data)} games -> {OUTPUT_PNG.resolve()}")


if __name__ == "__main__":
    main()
