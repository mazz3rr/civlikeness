"""
Plot the Ch/H column of the Civ-likeness ledger.

Chassis  = P1 + P2 + P3 + P4 + P6a + P7   (max 30)  -> X axis
House    = P5 + P6b                       (max 10)  -> Y axis

Both axes use uniform, equal-length data units (House is NOT rescaled to
span the same pixel range as Chassis), so the plot is a true 1:1 grid
where 1 House point occupies the same physical length as 1 Chassis point.

Recomputing from the pillar columns avoids parsing the "Ch/H" text string;
"""

# first version by Perplexity AI

import enum
from pathlib import Path

import math
from typing import Any, NamedTuple, Self

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.axes import Axes
from matplotlib.collections import PathCollection
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.figure import Figure
from numpy import dtype, ndarray

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
INPUT_XLSX = Path(__file__).parent.parent / "assets" / "ledger.xlsx"
SHEET = "ledger"
FIRST_ROW = 8  # 1-based; data starts here (row 7 is the header)
VAR_DIR = Path("var")
VAR_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_PNG = VAR_DIR / "ch_h_plot.png"

# Column layout of the ledger sheet (1-based, matches header in row 6/7):
#   A Game | B P1 Map | C P2 Expansion | D P3 Economy | E P4 Ladder
#   F P5 Peers (H) | G P6a Abstraction | H P6b Polity/Endings (H) | I P7 Time
#   J Ch/H | K Total | ...
GAME_COL = 1
# Chassis pillars (non-(H) pillars):
CHASSIS_COLS = 2, 3, 4, 5, 7, 9   # P1, P2, P3, P4, P6a, P7
# House pillars (the ones marked "(H)"):
HOUSE_COLS = 6, 8                # P5, P6b
WS_COL = 14
WS_VALUES = "N/A", "0 %", "25 %", "50 %", "75 %", "100 %"

# Short, recognizable labels for the dots. Anything not listed falls back to a
# heuristic shortener below, so adding new games never breaks the plot.
ALIASES = {
    "Against the Storm (2023–current, +2 DLC)":             "AtS",
    "Age of Empires II: DE (2019–current)":                 "AoE2",
    "Age of Wonders 3 (2014)":                              "AoW3",
    "Age of Wonders 4 (2023–current)":                      "AoW4",
    "Age of Wonders: Shadow Magic (2003)":                  "AoW:SM",
    "Alpha Centauri (1999, SMAX)":                          "SMAC",
    "Anno 1800 (2019–2022)":                                "A1800",
    "Beyond Earth (2014, +RT)":                             "BE",
    "Black & White (2001, +Creature Isle)":                 "B&W",
    "CK3 (2020–2026)":                                      "CK3",
    "Call to Power (1999, v1.2)":                           "CtP1",
    "Call to Power II (2000, retail final)":                "CtP2",
    "Civ1 (1991)":                                          "Civ1",
    "Civ2 (1996)":                                          "Civ2",
    "Civ3 (2001, Conquests)":                               "Civ3",
    "Civ4 (2005, BtS)":                                     "Civ4",
    "Civ5 (2010, BNW)":                                     "Civ5",
    "Civ6 (2016, GS)":                                      "Civ6",
    "Civ7 (2025, current)":                                 "Civ7",
    "Colonization (1994)":                                  "Col",
    "Conquest of Elysium 5 (2021–current)":                 "CoE5",
    "Dominions 6 (2024–current, 6.37)":                     "Dom6",
    "Doom (1993, Ultimate)":                                "Doom",
    "Eador: Genesis (2009–13, final)":                      "EG",
    "EU4 (2013–2025, final)":                               "EU4",
    "EVE Online (current)":                                 "EVE",
    "Endless Legend (2014, Symbiosis)":                     "EL",
    "Fantasy General (1996, +Island Challenge)":            "FG",
    "Football Manager 26 (final)":                          "FM26",
    "Frostpunk (2018–20, +Season Pass)":                    "FP",
    "Galactic Civilizations 2 (2006–08, Twilight)":         "GalCiv2",
    "Galactic Civilizations 3 (2014–19, Retribution)":      "GalCiv3",
    "Galactic Civilizations 4 (2022–current, Supernova)":   "GalCiv4",
    "Heroes of Might & Magic III (1999, Complete)":         "HoMM3",
    "Humankind (2021, TWR)":                                "HK",
    "Into the Breach (2018, AE)":                           "ItB",
    "MOO1 (1993)":                                          "MOO1",
    "MOO2 (1996)":                                          "MOO2",
    "Master of Magic (1994, CoM)":                          "MoM",
    "Offworld Trading Company (2016, final)":               "OTC",
    "Old World (2021, 7 DLC)":                              "OW",
    "Populous (1989, +The Promised Lands)":                 "Populous",
    "SimCity 4 (2003, Deluxe+641)":                         "SC4",
    "Songs of Conquest (2022–current, v1.4 + faction DLC)": "SoC",
    "SpellForce: Conquest of Eo (2023–25, +3 DLC)":         "Eo",
    "StarCraft II (melee, vs-AI)":                          "SC2",
    "Stellaris (2016–current)":                             "Stellaris",
    "The Sims 4 (live)":                                    "Sims4",
    "Total War: Warhammer III (2022, IE + 9.0)":            "TWW3",
    "Warcraft III: TFT (2003, melee vs-AI)":                "WC3",
    "Warhammer 40,000: Gladius (2018–26, 11 factions)":     "Gladius",
    "Warlords III (1997–98, + DLR)":                        "WL3",
    "XCOM 2 (2017, WotC)":                                  "XCOM2",
    "ZEPHON (2024–current, 1.3)":                           "ZEPHON",
}


def alias(full_name: str) -> str:
    """Return a short label for a game, with a heuristic fallback.
    """
    if sn := ALIASES.get(full_name):
        return sn
    # Fallback: take everything before the first '(' and trim.
    base = full_name.split("(", 1)[0].strip() if "(" in full_name else full_name.strip()
    if len(base) <= 14:
        return base
    # Otherwise keep initials of the first few words.
    words = base.split()
    if len(words) >= 2:
        return "".join(w[0] for w in words[:4]) or base[:12]
    return base[:12]


class WarShare(enum.Enum):
    N_A = "N/A"
    ABSENT = 0
    PERIPHERAL = 25
    SUBSTANTIAL = 50
    CORE = 75
    PURE = 100

    @property
    def is_na(self) -> bool:
        return self is self.N_A

    @classmethod
    def from_data(cls, data: str) -> Self:
        match data:
            case "N/A":
                return cls.N_A
            case "0 %":
                return cls.ABSENT
            case "25 %":
                return cls.PERIPHERAL
            case "50 %":
                return cls.SUBSTANTIAL
            case "75 %":
                return cls.CORE
            case "100 %":
                return cls.PURE
            case _:
                raise ValueError(f"Unknown WarShare value: '{data}', must be either of {WS_VALUES}")


class DataRow(NamedTuple):
    game: str
    chassis: float
    house: float
    warshare: WarShare


def viridis_red() -> LinearSegmentedColormap:
    # Custom colormap: viridis extended past its yellow endpoint to red, so the
    # very top of the scale (a perfect 40) is red instead of stopping at yellow.
    viridis_rgb = plt.cm.viridis(np.linspace(0, 1, 200))[:, :3]
    yellow = plt.cm.viridis(1.0)[:3]
    red = np.array([1.0, 0.0, 0.0])
    n_ext = 60
    ext_rgb = np.array([np.linspace(yellow[i], red[i], n_ext) for i in range(3)]).T
    return LinearSegmentedColormap.from_list(
        "viridis_red", np.vstack([viridis_rgb, ext_rgb[1:]])
    )


def gist_rainbow() -> LinearSegmentedColormap:
    """Sampled gist_rainbow: dark blue -> cyan -> green -> yellow -> red.
    """
    colors = plt.cm.gist_rainbow_r(np.linspace(0.25, 1, 256))
    return LinearSegmentedColormap.from_list("ws_rainbow", colors)


class Plotter:
    DPI = 200
    LABELS_FONTSIZE = 6
    REFLINES_FONTSIZE = 7
    DIAGONAL_FONTSIZE = 8

    def __init__(self, input_xlsx: str | Path = "", output_pgn: str | Path = "") -> None:
        self._input_xlsx = Path(input_xlsx) if input_xlsx else INPUT_XLSX
        self._output_pgn = Path(output_pgn) if output_pgn else OUTPUT_PNG
        self._rows: list[DataRow] = []
        self._cmap: LinearSegmentedColormap | None = None
        self._fig: Figure | None = None
        self._ax: Axes | None = None
        self._sc: PathCollection | None = None

    @staticmethod
    def _col_value(row: list, col: int, idx: int) -> float:
        v = row[col - 1]
        if v is None:
            raise ValueError(f"Column {col}, row {idx + FIRST_ROW} has no value")
        return float(v)

    def _load_rows(self) -> None:
        """Return list of (short_name, chassis, house, warshare) recomputed from pillars.
        """
        wb = openpyxl.load_workbook(self._input_xlsx, data_only=True)
        ws = wb[SHEET]

        for i, row in enumerate(ws.iter_rows(min_row=FIRST_ROW, values_only=True)):
            game = row[GAME_COL - 1]
            if game is None or str(game).strip() == "":
                continue
            vals = list(row)

            game = alias(str(game).strip())
            chassis = sum(self._col_value(vals, c, i) for c in CHASSIS_COLS)   # /30
            house = sum(self._col_value(vals, c, i) for c in HOUSE_COLS)       # /10
            war_share = WarShare.from_data(vals[WS_COL - 1])

            self._rows.append(DataRow(game, chassis, house, war_share))

    def _annotate(
        self,
        iter_lim: int = 5000,
        time_lim: float | None = None
    ) -> None:
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
                self._ax.text(xi, yi, n, fontsize=self.LABELS_FONTSIZE, zorder=6, bbox=_LABEL_BBOX)
                for n, xi, yi, *_ in self._rows
            ]
            adjust_text(
                texts,
                objects=self._sc,
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
                iter_lim=iter_lim,
                time_lim=time_lim,
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
            for n, xi, yi, *_ in self._rows:
                self._ax.annotate(
                    n, (xi, yi), textcoords="offset points", xytext=(7, 5),
                    fontsize=self.LABELS_FONTSIZE, zorder=5)

    def _plot_diagonal(self) -> None:
        # Reference line: the "3:1 ratio" diagonal from (0,0) to (30,10), i.e. a
        # House share of 1/4 of the total. Games above it are House-heavy
        # (relationship/ends-leaning); games below are Chassis-heavy (operational).
        # Dashed, semi-transparent, drawn under the points so it never crowds them.
        self._ax.plot(
            [0, 30],
            [0, 10],
            linestyle=(0, (5, 4)),
            color="#444444",
            alpha=0.6,
            linewidth=1.3,
            zorder=1
        )
        self._ax.text(
            1.2,
            0.55,
            "3:1 ratio line  (House = ¼ of total)",
            fontsize=self.DIAGONAL_FONTSIZE,
            color="#555555",
            rotation=math.degrees(math.atan2(10, 30)),
            zorder=1
        )

    def _plot_total_ref_lines(self, use_cmap=True) -> None:
        for total in (5, 10, 15, 20, 25, 30, 35):
            # For y = total - x, constrain the segment to:
            # 0 <= x <= 30 and 0 <= y <= 10.
            x0 = max(0, total - 10)
            x1 = min(30, total)
            if x0 > x1:
                continue

            y0 = total - x0
            y1 = total - x1
            color = self._cmap(total / 40) if use_cmap else "#777777"

            self._ax.plot(
                [x0, x1],
                [y0, y1],
                linestyle=":",
                color=color,
                linewidth=1.2,
                alpha=0.8,
                zorder=1,
            )

            if total == 35:
                # Label at the right-hand endpoint of the segment.
                self._ax.annotate(
                    f"total={total}",
                    xy=(x1, y1),
                    xytext=(0, 8),  # points: slightly above the endpoint
                    textcoords="offset points",
                    ha="left",
                    va="center",
                    rotation=-45,
                    rotation_mode="anchor",
                    fontsize=self.REFLINES_FONTSIZE,
                    color=color,
                    bbox=dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1.5),
                    annotation_clip=False,
                    zorder=2,
                )
            else:
                # Label at the left-hand endpoint of the segment.
                self._ax.annotate(
                    f"total={total}",
                    xy=(x0, y0),
                    xytext=(0, 8),  # points: slightly above the endpoint
                    textcoords="offset points",
                    ha="left",
                    va="center",
                    rotation=-45,
                    rotation_mode="anchor",
                    fontsize=self.REFLINES_FONTSIZE,
                    color=color,
                    bbox=dict(facecolor="white", edgecolor="none", alpha=0.8, pad=1.5),
                    annotation_clip=False,
                    zorder=2,
                )

    def _prepare_plot(
        self,
        name=""
    ) -> tuple[ndarray[tuple[Any, ...], dtype[Any]], ndarray[tuple[Any, ...], dtype[Any]]]:
        if not self._rows:
            print("Loading data...")
            self._load_rows()

        msg = f"Preparing {name} plot..." if name else "Preparing plot..."
        print(msg)

        self._fig, self._ax = plt.subplots(figsize=(16, 9))

        x = np.array([row.chassis for row in self._rows], dtype=float)
        y = np.array([row.house for row in self._rows], dtype=float)

        return x, y

    def plot(self) -> None:
        x, y = self._prepare_plot()

        self._cmap = viridis_red()

        # Color by total score (Chassis + House) on the nominal 0-40 scale.
        totals = x + y

        # Uniform units: equal aspect ratio so 1 House unit == 1 Chassis unit in
        # physical length. The axes' limits then define the visible region.
        self._sc = self._ax.scatter(
            x,
            y,
            c=totals,
            cmap=self._cmap,
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
        self._ax.set_xlim(-1, 37)
        self._ax.set_ylim(-0.5, 11.5)
        self._ax.set_aspect("equal", adjustable="box")   # <- the "uniform units" rule

        self._annotate()

        # Axes limits are set above (before adjust_text) so expand_axes can
        # push them further if needed. The aspect ratio is already set to equal.
        self._ax.set_xlabel("Chassis  (P1+P2+P3+P4+P6a+P7)   / 30")
        self._ax.set_ylabel("House  (P5+P6b)   / 10")
        self._ax.set_title("Civ-likeness: Chassis vs. House")

        # Light reference grid at every integer point.
        self._ax.set_xticks(np.arange(0, 31, 2))
        self._ax.set_yticks(np.arange(0, 11, 1))
        self._ax.grid(True, which="both", linestyle=":", linewidth=0.5, color="0.8", zorder=0)

        cbar = self._fig.colorbar(self._sc, ax=self._ax, shrink=0.7, pad=0.02)
        cbar.set_label("Total Ch+H")
        # Nominal 0-40 scale with 5-unit ticks (8 segments).
        cbar.set_ticks(np.arange(0, 41, 5))

        self._plot_diagonal()
        self._plot_total_ref_lines()

        self._fig.tight_layout()

        self._fig.savefig(self._output_pgn, dpi=self.DPI)
        print(f"Plotted {len(self._rows)} games -> {self._output_pgn.resolve()}")

    def plot_ws(self) -> None:
        """Plot Chassis vs. House, coloring games by WarShare.

        WarShare values 0/25/50/75/100 use a sampled gist rainbow colormap.
        Games with WarShare == "N/A" are shown in grey.
        """
        x, y = self._prepare_plot(name="Warshare")

        # Map the five numeric WarShare levels to 0, 25, 50, 75, 100.
        ws_values = np.array(
            [
                row.warshare.value if not row.warshare.is_na else np.nan
                for row in self._rows
            ],
            dtype=float,
        )

        self._cmap = gist_rainbow()

        # This is deliberately the only collection passed to adjustText.
        # It contains only finite color values and is safe as an obstacle source.
        numeric_mask = ~np.isnan(ws_values)
        self._sc = self._ax.scatter(
            x[numeric_mask],
            y[numeric_mask],
            c=ws_values[numeric_mask],
            cmap=self._cmap,
            vmin=0,
            vmax=100,
            s=70,
            edgecolors="white",
            linewidths=0.6,
            zorder=3,
        )

        # Plot N/A points separately, with literal grey.
        na_mask = ~numeric_mask
        if np.any(na_mask):
            self._ax.scatter(
                x[na_mask],
                y[na_mask],
                color="#808080",
                s=70,
                edgecolors="white",
                linewidths=0.6,
                zorder=3,
            )

        self._ax.set_xlim(-1, 37)
        self._ax.set_ylim(-0.5, 11.5)
        self._ax.set_aspect("equal", adjustable="box")

        self._annotate(iter_lim=800, time_lim=20)

        self._ax.set_xlabel("Chassis (P1+P2+P3+P4+P6a+P7) / 30")
        self._ax.set_ylabel("House (P5+P6b) / 10")
        self._ax.set_title(
            "Civ-likeness: Chassis vs. House, colored by WarShare\n"
            "Absent: 0% / Peripheral: 25% / Substantial: 50% / Core: 75% / Pure: 100%"
        )

        self._ax.set_xticks(np.arange(0, 31, 2))
        self._ax.set_yticks(np.arange(0, 11, 1))
        self._ax.grid(True, which="both", linestyle=":", linewidth=0.5, color="0.8", zorder=0)

        # Colorbar on the right, explicitly showing the five WarShare levels.
        cbar = self._fig.colorbar(
            self._sc,
            ax=self._ax,
            location="right",
            shrink=0.7,
            pad=0.02,
        )
        cbar.set_label("WarShare")
        cbar.set_ticks([0, 25, 50, 75, 100])
        cbar.set_ticklabels(["0 %", "25 %", "50 %", "75 %", "100 %"])
        cbar.set_label("WarShare (grey = N/A)")

        self._plot_diagonal()
        self._plot_total_ref_lines(use_cmap=False)

        self._fig.tight_layout()

        output_path = self._output_pgn.with_name(
            self._output_pgn.stem + "_ws" + self._output_pgn.suffix
        )
        self._fig.savefig(output_path, dpi=self.DPI)
        print(f"Plotted {len(self._rows)} games -> {output_path.resolve()}")


if __name__ == "__main__":
    plotter = Plotter()
    plotter.plot()
    plotter.plot_ws()
