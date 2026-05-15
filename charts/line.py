"""Line chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class LineChart(ChartBuilder):
    """A line chart: each series is drawn as a connected line across categories."""

    chart_type = "line"
    data_shape = (
        "List of row objects ordered along the x-axis. Each row has one "
        "category key ('category', 'label' or 'name') and one or more numeric "
        "series keys. At least two rows are required so each line has points "
        "to connect."
    )
    example = {
        "chart_type": "line",
        "title": "Monthly Active Users",
        "data": [
            {"category": "Jan", "2024": 100, "2025": 130},
            {"category": "Feb", "2024": 110, "2025": 145},
            {"category": "Mar", "2024": 115, "2025": 160},
        ],
    }

    def _validate(self) -> list[str]:
        return self.validate_categorical(min_categories=2, min_series=1)

    def build_table(self) -> Table:
        return self.categorical_table()
