"""Stacked bar / column chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class StackedBarChart(ChartBuilder):
    """A stacked bar/column chart: series stack on top of each other per category."""

    chart_type = "stacked_bar"
    data_shape = (
        "List of row objects. Each row has one category key ('category', "
        "'label' or 'name') and one or more numeric series keys; the series "
        "are stacked within each category's bar. Two or more series make the "
        "stacking meaningful."
    )
    example = {
        "chart_type": "stacked_bar",
        "title": "Revenue Mix by Quarter",
        "data": [
            {"category": "Q1", "Product A": 40, "Product B": 30, "Product C": 20},
            {"category": "Q2", "Product A": 50, "Product B": 35, "Product C": 25},
        ],
    }

    def _validate(self) -> list[str]:
        return self.validate_categorical(min_categories=1, min_series=1)

    def build_table(self) -> Table:
        return self.categorical_table()
