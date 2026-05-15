"""Area chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class AreaChart(ChartBuilder):
    """An area chart: series are drawn as filled, stacked regions over categories."""

    chart_type = "area"
    data_shape = (
        "List of row objects ordered along the x-axis. Each row has one "
        "category key ('category', 'label' or 'name') and one or more numeric "
        "series keys. At least two rows are required to form an area."
    )
    example = {
        "chart_type": "area",
        "title": "Cumulative Signups",
        "data": [
            {"category": "Week 1", "Free": 200, "Paid": 50},
            {"category": "Week 2", "Free": 320, "Paid": 90},
            {"category": "Week 3", "Free": 410, "Paid": 150},
        ],
    }

    def _validate(self) -> list[str]:
        return self.validate_categorical(min_categories=2, min_series=1)

    def build_table(self) -> Table:
        return self.categorical_table()
