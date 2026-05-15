"""Waterfall (bridge) chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class WaterfallChart(ChartBuilder):
    """A waterfall / bridge chart showing how a value builds up step by step.

    Each row contributes one bar. Positive values step up and negative values
    step down. Whether a bar is a connector or a calculated total is configured
    on the think-cell waterfall inside the PowerPoint template.
    """

    chart_type = "waterfall"
    data_shape = (
        "Ordered list of row objects. Each row has one category key "
        "('category', 'label' or 'name') and a single numeric series key "
        "holding the step delta (use negative numbers to step down). At "
        "least two rows are required."
    )
    example = {
        "chart_type": "waterfall",
        "title": "Profit Bridge 2025",
        "data": [
            {"category": "Opening", "Value": 100},
            {"category": "Sales", "Value": 45},
            {"category": "Costs", "Value": -30},
            {"category": "Closing", "Value": 115},
        ],
    }

    def _validate(self) -> list[str]:
        return self.validate_categorical(min_categories=2, min_series=1)

    def build_table(self) -> Table:
        return self.categorical_table()
