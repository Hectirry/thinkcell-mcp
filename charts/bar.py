"""Bar / column chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class BarChart(ChartBuilder):
    """A simple bar/column chart: one value per category per series.

    A single series renders as plain bars; multiple series render as grouped
    (clustered) bars.
    """

    chart_type = "bar"
    data_shape = (
        "List of row objects. Each row has one category key ('category', "
        "'label' or 'name') and one or more numeric series keys. Multiple "
        "series render as grouped bars."
    )
    example = {
        "chart_type": "bar",
        "title": "Revenue by Region",
        "data": [
            {"category": "North", "Revenue": 120},
            {"category": "South", "Revenue": 95},
            {"category": "West", "Revenue": 140},
        ],
    }

    def _validate(self) -> list[str]:
        return self.validate_categorical(min_categories=1, min_series=1)

    def build_table(self) -> Table:
        return self.categorical_table()
