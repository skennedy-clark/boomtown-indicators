"""
regional-indicators/transform/xlsx_update/audit.py

Pre-write audits for the indicators workbook.

The workbook has been maintained by hand for many years. Target cells
can contain leftover formulas or stray values, and historical rows can
contain transcription errors. Every write is therefore preceded by up to
three checks:

1. Cell audit: what the target cell currently holds -- empty, a
   formula, a value that already matches, or unexpected content.

2. Series audit: whether the new value is consistent with the existing
   row. Detects a scale mismatch (a sign of the wrong row or geography
   level) and, from the shape of the series alone, isolated outliers and
   step changes.

3. Historical audit: when the full source series is available, each
   existing year is compared with the source value. This supersedes the
   shape-based outlier check, which cannot distinguish a transcription
   error from genuine volatility.

The audits never modify data. They return structured results that the
writers use to decide whether to write, and that are reported for
review.
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
        """True for FORMULA and UNEXPECTED_CONTENT, which must not be
        overwritten without review. EMPTY and MATCHES_NEW_VALUE are safe.
        """
        return self.state in (CellState.FORMULA, CellState.UNEXPECTED_CONTENT)


def audit_cell(cell, new_value, tolerance: float = 0.01) -> CellAuditResult:
    """Classify the current content of a target cell.

    tolerance: relative difference within which an existing value counts
        as matching the new value (default 1%, to absorb rounding and
        minor revisions).
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

    # xlwings can return numeric cell values as decimal.Decimal, so Decimal
    # is accepted alongside int and float; otherwise matching values would
    # be reported as unexpected content.
    from decimal import Decimal
    numeric_types = (int, float, Decimal)
    if isinstance(value, numeric_types) and isinstance(new_value, numeric_types):
        value_f, new_value_f = float(value), float(new_value)
        if value_f == 0 and new_value_f == 0:
            close = True
        elif value_f == 0 or new_value_f == 0:
            close = False
        else:
            close = abs(value_f - new_value_f) / abs(value_f) <= tolerance
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
    SCALE_MISMATCH = "scale_mismatch"        # possible wrong row or geography level
    ISOLATED_OUTLIER = "isolated_outlier"    # single point that departs from, then returns to, the series
    STEP_CHANGE = "step_change"              # new value departs from the latest existing value
    TOO_SHORT = "too_short"                  # not enough history to assess


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
    """Check a new value against the existing {year: value} series.

    SCALE_MISMATCH: the new value lies outside `scale_bounds` times the
        series median. Indicates a wrong target row or geography level.
    ISOLATED_OUTLIER: an existing point departs from both neighbours by
        more than `spike_threshold` and the series then reverts.
    STEP_CHANGE: the new value differs from the latest existing value by
        more than `spike_threshold`.

    The outlier and step checks infer from shape only; prefer
    audit_historical_series() when a source series is available.
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
    MINOR_DISCREPANCY = "minor_discrepancy"  # differs by less than an order of
                                              # magnitude; reported, does not block
                                              # (sources do revise published
                                              # figures)
    MAJOR_DISCREPANCY = "major_discrepancy"  # differs by an order of magnitude
                                              # or more; treated as an error and
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
    comparisons: list  # list[YearComparison]; only years that differ

    @property
    def has_major_discrepancy(self) -> bool:
        return any(c.flag == HistoricalMatchFlag.MAJOR_DISCREPANCY for c in self.comparisons)

    @property
    def notes(self) -> list:
        return [c for c in self.comparisons if c.flag == HistoricalMatchFlag.MINOR_DISCREPANCY]


def audit_historical_series(
    existing_series: dict,
    source_series: dict,
    major_discrepancy_ratio: float = 10.0,   # ratio at or above which a difference blocks
    negligible_tolerance: float = 0.02,      # relative difference below which it is ignored
) -> HistoricalAuditResult:
    """Compare existing workbook values with the source series.

    Only years present in both series are compared. Differences within
    `negligible_tolerance` are ignored. A ratio at or above
    `major_discrepancy_ratio` is a MAJOR_DISCREPANCY and blocks the
    write; smaller differences are MINOR_DISCREPANCY and are reported
    only, since published sources revise historical figures.
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
    """Combined audit result for one write.

    Callers check `safe_to_write`; the remaining fields describe why a
    write was flagged.
    """
    cell_audit: CellAuditResult
    series_audit: SeriesAuditResult
    historical_audit: "HistoricalAuditResult | None" = None
    override_reason: str | None = None

    @property
    def safe_to_write(self) -> bool:
        if self.override_reason is not None:
            # An exact verified override takes precedence over every audit.
            return True

        cell_ok = not self.cell_audit.needs_review

        if self.historical_audit is not None:
            # With a source series, the historical comparison replaces the
            # shape-based outlier and step checks. SCALE_MISMATCH still blocks.
            series_ok = self.series_audit.flag != SeriesFlag.SCALE_MISMATCH
            historical_ok = not self.historical_audit.has_major_discrepancy
            return cell_ok and series_ok and historical_ok

        series_ok = self.series_audit.flag in (SeriesFlag.OK, SeriesFlag.TOO_SHORT)
        return cell_ok and series_ok

    @property
    def should_write_with_flag(self) -> bool:
        """True when the write is blocked only by a series or historical
        concern, not by the content of the target cell.

        Writers that support visual flagging may then write the value
        and mark it (bold red) for review. A cell-level block (a formula
        or unexpected content) is never written through.
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
            # Report the override together with the flags it overrides.
            if parts:
                return f"OVERRIDE APPLIED ({self.override_reason}) — originally flagged: " + " | ".join(parts)
            return f"OVERRIDE APPLIED ({self.override_reason})"
        if not parts:
            return f"OK ({self.cell_audit.state.value}, {self.series_audit.flag.value})"
        return " | ".join(parts)


def apply_write_formatting(cell, safe_to_write: bool, should_write_with_flag: bool) -> None:
    """Apply font formatting for a write.

    A clean write is set to regular black, clearing any earlier flag.
    A flagged write is set to bold red. xlwings requires an explicit RGB
    tuple for Font.color; None is not accepted.
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
    """Describe the outcome of a write attempt.

    Returns (result_line, counts_as_written, counts_as_flagged) for the
    three cases: written, written and flagged for review, or not
    written.
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