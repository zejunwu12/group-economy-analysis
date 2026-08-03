"""Quarter parsing and derived reporting-period metadata."""

from __future__ import annotations

from dataclasses import dataclass
import re


_QUARTER_CODE = re.compile(r"^(?P<year>\d{4})Q(?P<quarter>[1-4])$")
_CHINESE_QUARTERS = {
    1: "第一季度",
    2: "第二季度",
    3: "第三季度",
    4: "第四季度",
}
_QUARTER_END_DATES = {
    1: (3, 31),
    2: (6, 30),
    3: (9, 30),
    4: (12, 31),
}


class QuarterError(ValueError):
    """The runtime quarter code is missing or invalid."""


@dataclass(frozen=True)
class QuarterContext:
    """Normalized quarter values shared by runtime consumers."""

    code: str
    year: int
    quarter: int
    label: str
    start_date: str
    end_date: str

    @classmethod
    def parse(cls, code: str) -> "QuarterContext":
        """Parse a strict ``YYYYQn`` code and derive Chinese period labels."""
        if not isinstance(code, str):
            raise QuarterError("季度参数必须是 YYYYQn 格式字符串，例如 2026Q2")

        normalized = code.strip().upper()
        match = _QUARTER_CODE.fullmatch(normalized)
        if match is None:
            raise QuarterError(
                f"季度参数格式错误: {code!r}；应使用 YYYYQ1 至 YYYYQ4，例如 2026Q2"
            )

        year = int(match.group("year"))
        quarter = int(match.group("quarter"))
        end_month, end_day = _QUARTER_END_DATES[quarter]
        return cls(
            code=f"{year:04d}Q{quarter}",
            year=year,
            quarter=quarter,
            label=f"{year}年{_CHINESE_QUARTERS[quarter]}",
            start_date=f"{year}年1月1日",
            end_date=f"{year}年{end_month}月{end_day}日",
        )

    def next(self) -> "QuarterContext":
        """Return the immediately following calendar quarter."""
        if self.quarter == 4:
            return self.parse(f"{self.year + 1}Q1")
        return self.parse(f"{self.year}Q{self.quarter + 1}")

    def previous(self) -> "QuarterContext":
        """Return the immediately preceding calendar quarter."""
        if self.quarter == 1:
            return self.parse(f"{self.year - 1}Q4")
        return self.parse(f"{self.year}Q{self.quarter - 1}")

    def previous_year(self) -> "QuarterContext":
        """Return the same quarter in the preceding year."""
        return self.parse(f"{self.year - 1}Q{self.quarter}")

    def as_config(self) -> dict[str, str | int]:
        """Expose the legacy in-memory shape used by existing modules."""
        return {
            "code": self.code,
            "year": self.year,
            "quarter": self.quarter,
            "label": self.label,
            "start_date": self.start_date,
            "end_date": self.end_date,
        }
