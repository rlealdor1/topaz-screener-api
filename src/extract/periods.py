"""Period axis utilities — mapping fiscal periods to column labels."""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Optional


@dataclass(frozen=True)
class Period:
    """A reporting period. `quarter` is 1-4, or None for full-year."""
    fiscal_year: int
    quarter: Optional[int]  # None for FY
    start: date
    end: date

    @property
    def is_annual(self) -> bool:
        return self.quarter is None

    @property
    def label(self) -> str:
        if self.is_annual:
            return str(self.fiscal_year)
        return f"{self.quarter}Q{str(self.fiscal_year)[-2:]}"

    @property
    def days(self) -> int:
        return (self.end - self.start).days + 1

    def __lt__(self, other: "Period") -> bool:
        # Chronological order: older first. Annual FY sorts AFTER its Q4 (visual order in template).
        if self.fiscal_year != other.fiscal_year:
            return self.fiscal_year < other.fiscal_year
        # Within same FY: Q1, Q2, Q3, Q4, FY
        self_key = 5 if self.is_annual else self.quarter
        other_key = 5 if other.is_annual else other.quarter
        return self_key < other_key


def parse_date(s: str) -> date:
    return date.fromisoformat(s)


def classify_duration(start: date, end: date) -> Optional[str]:
    """Return 'quarter' for ~90-day periods, 'annual' for ~365-day, None otherwise."""
    days = (end - start).days + 1
    if 80 <= days <= 100:
        return "quarter"
    if 350 <= days <= 380:
        return "annual"
    return None


def quarter_from_end_date(end: date) -> Optional[int]:
    """Infer quarter number from end date. Most US issuers use calendar-year quarters."""
    # Simple: Q1=Jan-Mar, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Oct-Dec
    if end.month in (3, 4) and end.day >= 28:
        return 1
    if end.month in (6, 7) and end.day >= 28:
        return 2
    if end.month in (9, 10) and end.day >= 28:
        return 3
    if end.month in (12, 1) and (end.day >= 28 or end.month == 1):
        return 4
    # Fallback: map by month number
    return ((end.month - 1) // 3) + 1


def build_period_axis(latest_fy: int, quarters_back: int = 12) -> list[Period]:
    """Build the period axis for the Excel sheet.

    Returns periods in display order:
      [{Y-2} FY, Q1 Y-1, Q2 Y-1, Q3 Y-1, Q4 Y-1, {Y-1} FY, Q1 Y, Q2 Y, Q3 Y, Q4 Y, {Y} FY]

    where Y is latest_fy and the number of quarters is approximately `quarters_back`.
    """
    periods: list[Period] = []
    years_back = max(2, (quarters_back // 4) - 1)
    start_year = latest_fy - years_back

    # Earliest FY
    periods.append(Period(
        fiscal_year=start_year,
        quarter=None,
        start=date(start_year, 1, 1),
        end=date(start_year, 12, 31),
    ))

    for y in range(start_year + 1, latest_fy + 1):
        for q in range(1, 5):
            # Calendar-quarter boundaries; exact boundary matching handled by data filter.
            q_start_month = (q - 1) * 3 + 1
            q_end_month = q * 3
            q_start = date(y, q_start_month, 1)
            # last day of quarter
            if q_end_month == 12:
                q_end = date(y, 12, 31)
            else:
                q_end = date(y, q_end_month + 1, 1) - timedelta(days=1)
            periods.append(Period(y, q, q_start, q_end))
        periods.append(Period(y, None, date(y, 1, 1), date(y, 12, 31)))

    return periods
