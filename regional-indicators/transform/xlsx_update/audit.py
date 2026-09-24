"""
regional-indicators/transform/xlsx_update/audit.py -- pre-write auditing
for a workbook that's been maintained by hand for years and is known to
contain "crud": stray values left over from ad-hoc analysis done
directly in the sheet, formulas nobody meant to leave behind, and
historical data points that may be a one-off miscopy rather than a
genuine update.

Three checks, meant to run BEFORE any write:

1. CELL AUDIT -- what's currently in the exact cell we're about to
   write into? Empty, a formula, something that already matches the
   new value, or genuinely unexpected content.

2. SERIES AUDIT -- does the new value fit the EXISTING historical row
   in a way consistent with a genuine update? Two sub-checks: scale
   mismatch (wrong row / wrong geography level) and, when no ground-
   truth history is available, a shape-based guess at whether a
   nearby point looks like an isolated miscopy.

3. HISTORICAL AUDIT -- when a full freshly re-fetched source series is
   available, compare every existing year against the real source
   value directly (ground truth), rather than guessing from shape
   alone. This is strictly more reliable than the series audit's
   shape-based guess, and supersedes it when available -- confirmed on
   real data: Isaac's 2012 NRW figure looked like an isolated miscopy
   by shape alone, but exactly matches the real source; the ground-
   truth check correctly clears it.

None of these fix anything automatically -- all three produce a
structured report for a human to read. Same principle throughout: flag
it loudly, don't guess.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


# ── Cell audit ─────────────────────────────────────────────────────────────

class CellState(Enum):
    EMPTY = "empty"
    FORMULA = "formula"
    MATCHES_NEW_VALUE = "matches_new_value"
    UNEXPECTED_CONTENT = "unexpected_content"


@dataclass
class CellAuditResult:
    state: CellState
    coordinate: str
    existing_value: object = None
    existing_formula: str | None = None
    new_value: object = None
    note: str = ""

    @property
    def needs_review(self) -> bool:
        """FORMULA and UNEXPECTED_CONTENT should stop a human, not get
        silently overwritten. EMPTY and MATCHES_NEW_VALUE are safe to
        write through without interrupting anyone."""
        return self.state in (CellState.FORMULA, CellState.UNEXPECTED_CONTENT)


def audit_cell(cell, new_value, tolerance: float = 0.01) -> CellAuditResult:
    """Inspect a single xlwings Range/cell before writing into it.

    tolerance: relative difference allowed for "matches new value"
    (default 1% -- covers minor rounding/revision differences without
    masking a genuinely different figure).
    """
    formula = cell.formula
    value = cell.value
    coord = cell.address

    is_formula = isinstance(formula, str) and formula.startswith("=")

    if value is None and not is_formula:
        return CellAuditResult(CellState.EMPTY, coord, value, formula, new_value)

    if is_formula:
        return CellAuditResult(
            CellState.FORMULA, coord, value, formula, new_value,
            note=(
                f"Cell contains a formula ({formula!r}), not a plain value. "
                f"Overwriting it with a hardcoded number is a bigger change "
                f"than replacing a stray value -- this formula may be "
                f"referenced by other cells elsewhere in the workbook, or "
                f"be someone's in-progress analysis."
            ),
        )

    if isinstance(value, (int, float)) and isinstance(new_value, (int, float)):
        if value == 0 and new_value == 0:
            close = True
        elif value == 0 or new_value == 0:
            close = False
        else:
            close = abs(value - new_value) / abs(value) <= tolerance
        if close:
            return CellAuditResult(
                CellState.MATCHES_NEW_VALUE, coord, value, formula, new_value,
                note="Existing value already matches the new value within tolerance.",
            )

    return CellAuditResult(
        CellState.UNEXPECTED_CONTENT, coord, value, formula, new_value,
        note=(
            f"Cell already contains {value!r}, which doesn't match the new "
            f"value {new_value!r} and isn't empty. Could be a stray note, "
            f"leftover ad-hoc analysis someone did directly in the sheet, "
            f"or data that's actually correct and the NEW value is what's "
            f"wrong -- needs a human to look, this function will not guess."
        ),
    )


# ── Series audit ─────────────────────────────────────────────────────────

class SeriesFlag(Enum):
    OK = "ok"
    SCALE_MISMATCH = "scale_mismatch"        # possible wrong-row/wrong-geography-level
    ISOLATED_OUTLIER = "isolated_outlier"    # likely single miscopy, series reverts after
    STEP_CHANGE = "step_change"              # new value breaks from trend, no reversion to judge yet
    TOO_SHORT = "too_short"                  # not enough history to judge anything


@dataclass
class SeriesAuditResult:
    flag: SeriesFlag
    note: str = ""
    outlier_years: list = field(default_factory=list)


def audit_series(
    existing_series: dict,
    new_year: int,
    new_value,
    scale_bounds: tuple = (0.3, 3.0),
    spike_threshold: float = 0.25,
) -> SeriesAuditResult:
    """Check an existing {year: value} time series against a new
    incoming value for two different failure modes:

    1. SCALE_MISMATCH -- the new value is wildly different in
       magnitude from the historical series (outside scale_bounds x
       the series median). A wrong-row / wrong-geography-level red
       flag about the WRITE TARGET, not a comment on whether the new
       value itself is correct.

    2. ISOLATED_OUTLIER vs STEP_CHANGE, within the EXISTING series --
       a single miscopied point spikes away from both neighbours and
       the series reverts afterward; a genuine revision or real-world
       change tends to persist rather than revert. This function
       flags the PATTERN only -- see audit_historical_series() for a
       ground-truth version of this same question, which supersedes
       this shape-based guess when a source series is available.
    """
    values = sorted(existing_series.items())
    if len(values) < 3:
        return SeriesAuditResult(
            SeriesFlag.TOO_SHORT,
            note="Fewer than 3 existing data points -- not enough history to judge.",
        )

    nums_sorted = sorted(v for _, v in values)
    n = len(nums_sorted)
    median = (
        nums_sorted[n // 2] if n % 2
        else (nums_sorted[n // 2 - 1] + nums_sorted[n // 2]) / 2
    )

    if median != 0:
        ratio = new_value / median
        if not (scale_bounds[0] <= ratio <= scale_bounds[1]):
            return SeriesAuditResult(
                SeriesFlag.SCALE_MISMATCH,
                note=(
                    f"New value {new_value:,} is {ratio:.1f}x the existing "
                    f"series median ({median:,.0f}) -- outside the expected "
                    f"{scale_bounds[0]}x-{scale_bounds[1]}x range. Possible "
                    f"wrong row or wrong geography level, not just an "
                    f"unusual year."
                ),
            )

    outlier_years = []
    for i in range(1, len(values) - 1):
        year, val = values[i]
        prev_val = values[i - 1][1]
        next_val = values[i + 1][1]
        if prev_val == 0 or val == 0:
            continue
        jump_from_prev = abs(val - prev_val) / abs(prev_val)
        revert_to_next = abs(next_val - prev_val) / abs(prev_val)
        if jump_from_prev > spike_threshold and revert_to_next < spike_threshold / 2:
            outlier_years.append(year)

    if outlier_years:
        return SeriesAuditResult(
            SeriesFlag.ISOLATED_OUTLIER,
            note=(
                f"Year(s) {outlier_years} jump away from both neighbours "
                f"and the series reverts afterward -- pattern consistent "
                f"with a single miscopied/typo'd point rather than a "
                f"genuine change. Worth checking against the original "
                f"source for that specific year before trusting it."
            ),
            outlier_years=outlier_years,
        )

    last_year, last_val = values[-1]
    if last_val != 0:
        jump = abs(new_value - last_val) / abs(last_val)
        if jump > spike_threshold:
            return SeriesAuditResult(
                SeriesFlag.STEP_CHANGE,
                note=(
                    f"New value for {new_year} ({new_value:,}) differs "
                    f"from {last_year}'s value ({last_val:,}) by "
                    f"{jump:.0%}. Could be a genuine change (population "
                    f"growth, a revised rainfall/crime figure) or an "
                    f"error -- this check can't tell the difference, only "
                    f"flag it for a human to check against the source."
                ),
            )

    return SeriesAuditResult(SeriesFlag.OK)


# ── Historical cross-check against freshly re-fetched source data ─────────

class HistoricalMatchFlag(Enum):
    MINOR_DISCREPANCY = "minor_discrepancy"  # differs, but not by an order of
                                              # magnitude -- noted, not assumed
                                              # wrong (could be a legitimate
                                              # published revision)
    MAJOR_DISCREPANCY = "major_discrepancy"  # differs by roughly an order of
                                              # magnitude or more -- a clear
                                              # error (e.g. a mistyped digit),
                                              # blocks the write


@dataclass
class YearComparison:
    year: int
    existing_value: float
    source_value: float
    flag: HistoricalMatchFlag
    ratio: float


@dataclass
class HistoricalAuditResult:
    comparisons: list  # list[YearComparison], only years with a real discrepancy

    @property
    def has_major_discrepancy(self) -> bool:
        return any(c.flag == HistoricalMatchFlag.MAJOR_DISCREPANCY for c in self.comparisons)

    @property
    def notes(self) -> list:
        return [c for c in self.comparisons if c.flag == HistoricalMatchFlag.MINOR_DISCREPANCY]


def audit_historical_series(
    existing_series: dict,
    source_series: dict,
    major_discrepancy_ratio: float = 10.0,   # "an order of magnitude", taken literally
    negligible_tolerance: float = 0.02,      # within 2% isn't worth reporting at all
) -> HistoricalAuditResult:
    """Compare every year where we have BOTH an existing workbook value
    AND a freshly re-fetched source value -- ground truth, not a
    statistical guess.

    Real result confirmed on real data: Isaac's LGA NRW series (2011:
    13,590 / 2012: 17,125 / 2013: 14,950) exactly matches a fresh
    Bowen Basin download -- the shape-based series audit flagged 2012
    as a likely miscopy purely because it spikes away from its
    neighbours, with no way to know it's genuine workforce volatility.
    This function correctly finds zero discrepancy for all three
    years, because it checks against the real number instead of
    guessing from shape alone.

    Years present in only one of the two series aren't compared. A
    ratio >= major_discrepancy_ratio is treated as a clear error and
    blocks the write; anything smaller (but outside
    negligible_tolerance) is noted but does NOT block -- published
    sources do sometimes revise historical figures by a modest amount.
    """
    comparisons = []
    for year, existing_val in existing_series.items():
        if year not in source_series:
            continue
        source_val = source_series[year]
        if existing_val == source_val:
            continue

        larger = max(abs(existing_val), abs(source_val))
        smaller = min(abs(existing_val), abs(source_val))
        if larger == 0:
            continue
        if abs(existing_val - source_val) / larger <= negligible_tolerance:
            continue

        ratio = larger / smaller if smaller != 0 else float("inf")
        flag = (
            HistoricalMatchFlag.MAJOR_DISCREPANCY
            if ratio >= major_discrepancy_ratio
            else HistoricalMatchFlag.MINOR_DISCREPANCY
        )
        comparisons.append(YearComparison(year, existing_val, source_val, flag, ratio))

    return HistoricalAuditResult(comparisons)


# ── Combined report ────────────────────────────────────────────────────────

@dataclass
class WriteAuditReport:
    """Combined result of all audits for one (town, indicator, year)
    write. safe_to_write is the single thing calling code should check
    -- everything else is context for the human reviewing a flagged
    entry."""
    cell_audit: CellAuditResult
    series_audit: SeriesAuditResult
    historical_audit: "HistoricalAuditResult | None" = None
    override_reason: str | None = None

    @property
    def safe_to_write(self) -> bool:
        if self.override_reason is not None:
            # A verified override always wins -- it exists specifically
            # to write through a flag that's been independently checked
            # and confirmed correct. See verified_overrides.toml.
            return True

        cell_ok = not self.cell_audit.needs_review

        if self.historical_audit is not None:
            # Ground-truth comparison supersedes the shape-based
            # ISOLATED_OUTLIER/STEP_CHANGE guesses when available --
            # SCALE_MISMATCH (wrong-row detection) is a different,
            # still-useful check and keeps blocking regardless.
            series_ok = self.series_audit.flag != SeriesFlag.SCALE_MISMATCH
            historical_ok = not self.historical_audit.has_major_discrepancy
            return cell_ok and series_ok and historical_ok

        series_ok = self.series_audit.flag in (SeriesFlag.OK, SeriesFlag.TOO_SHORT)
        return cell_ok and series_ok

    @property
    def should_write_with_flag(self) -> bool:
        """True when the ONLY reason safe_to_write is False is a
        series-shape or historical-discrepancy concern -- NEVER a
        cell-level block (a formula or unexpected existing content).
        Distinguishes "this number looks unusual, write it but flag it
        visually for review" from "this cell isn't safe to touch at
        all, don't write anything" -- overwriting a live formula or
        someone's stray note is a different, higher-stakes risk than a
        suspicious-but-plausible number, and must never be silently
        written over regardless of formatting. Callers that support
        visual flagging (e.g. writing the value in bold red rather than
        leaving the cell untouched) should check this after
        safe_to_write comes back False; callers that don't support it
        can ignore this and keep the existing block-everything
        behaviour, which remains correct either way.
        """
        if self.safe_to_write:
            return False
        if self.cell_audit.needs_review:
            return False
        return True

    def summary_line(self) -> str:
        parts = []
        if self.cell_audit.needs_review:
            parts.append(f"CELL: {self.cell_audit.note}")
        if self.series_audit.flag not in (SeriesFlag.OK, SeriesFlag.TOO_SHORT):
            parts.append(f"SERIES: {self.series_audit.note}")
        if self.historical_audit:
            for c in self.historical_audit.comparisons:
                label = "MAJOR DISCREPANCY" if c.flag == HistoricalMatchFlag.MAJOR_DISCREPANCY else "note"
                parts.append(
                    f"HISTORICAL {label}: {c.year} existing={c.existing_value:,} vs "
                    f"source={c.source_value:,} ({c.ratio:.1f}x)"
                )
        if self.override_reason is not None:
            # Show the override AND whatever it's overriding -- full
            # transparency about why this was flagged in the first
            # place, not just that it got waved through.
            if parts:
                return f"OVERRIDE APPLIED ({self.override_reason}) — originally flagged: " + " | ".join(parts)
            return f"OVERRIDE APPLIED ({self.override_reason})"
        if not parts:
            return f"OK ({self.cell_audit.state.value}, {self.series_audit.flag.value})"
        return " | ".join(parts)


def apply_write_formatting(cell, safe_to_write: bool, should_write_with_flag: bool) -> None:
    """Shared font formatting for the write-with-flag pattern (2026-09-24,
    promoted here once a second indicator -- Business -- needed the same
    "flag with colour, don't block" behaviour Income already had, per
    Steve's stated preference: small-count data especially benefits,
    since the shape-based series checks trip constantly on trivial
    absolute changes when the underlying numbers are small, and on
    pre-existing historical anomalies unrelated to the new write at
    all -- flagging visually rather than blocking means neither ever
    withholds real data, just marks it for optional review).

    A clean write resets to plain black, not bold -- a cell that was
    once flagged and is now writing cleanly shouldn't stay visually
    flagged forever. CONFIRMED via a real xlwings crash (2026-09-23):
    Font.color has no "None means automatic" pathway -- must be an
    explicit RGB tuple, black (0,0,0) here, not None.
    """
    if safe_to_write:
        cell.number_format = "General"
        cell.font.bold = False
        cell.font.color = (0, 0, 0)
    elif should_write_with_flag:
        cell.number_format = "General"
        cell.font.bold = True
        cell.font.color = (255, 0, 0)


def describe_write_outcome(label: str, report: "WriteAuditReport", coord: str, extra_note: str = "") -> tuple[str, bool, bool]:
    """Shared three-way outcome description for a write attempt --
    written cleanly, written but visually flagged (see
    should_write_with_flag), or genuinely not written at all (a
    cell-level block, never overridden regardless of formatting).
    Returns (result_line, counts_as_written, counts_as_flagged).
    """
    suffix = f" ({extra_note})" if extra_note else ""
    if report.safe_to_write and not report.should_write_with_flag:
        return f"{label}: WRITTEN{suffix} -> {coord}", True, False
    if report.should_write_with_flag:
        return (
            f"{label}: WRITTEN but FLAGGED FOR REVIEW (bold red in sheet){suffix} "
            f"-> {coord} — {report.summary_line()}",
            True, True,
        )
    return f"{label}: FLAGGED, not written — {report.summary_line()}", False, True