from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta


@dataclass(frozen=True)
class LaunchDay:
    date: date
    day_number: int
    week_number: int

    @property
    def day_label(self) -> str:
        return f"D{self.day_number:02d}"

    @property
    def week_label(self) -> str:
        return f"第{self.week_number}周"

    @property
    def is_week_start(self) -> bool:
        return (self.day_number - 1) % 7 == 0


def build_timeline(d01: date, latest_covered_date: date) -> list[LaunchDay]:
    """Build a gap-free natural-day timeline; it does not create metric values.

    Weeks follow the uploaded template's seven-day blocks, anchored at D01.
    Missing source dates remain missing in the later metric join.
    """
    if latest_covered_date < d01:
        raise ValueError("最新覆盖日期早于 D01")
    count = (latest_covered_date - d01).days + 1
    return [
        LaunchDay(d01 + timedelta(days=offset), offset + 1, offset // 7 + 1)
        for offset in range(count)
    ]
