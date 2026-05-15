"""Marimekko (Mekko) chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class MekkoChart(ChartBuilder):
    """A Marimekko chart: stacked bars whose column widths vary by category.

    Segment heights come from the numeric series, exactly like a stacked bar.
    Column widths are configured on the think-cell Mekko inside the PowerPoint
    template (commonly driven by a dedicated 'Width' series in the data).
    """

    chart_type = "mekko"
    data_shape = (
        "List of row objects. Each row has one category key ('category', "
        "'label' or 'name') and one or more numeric series keys for the "
        "stacked segments. Optionally include a 'Width' series to drive the "
        "per-category column widths."
    )
    example = {
        "chart_type": "mekko",
        "title": "Market Share by Segment",
        "data": [
            {"category": "Segment A", "Company X": 60, "Company Y": 40, "Width": 70},
            {"category": "Segment B", "Company X": 35, "Company Y": 65, "Width": 30},
        ],
    }

    def _validate(self) -> list[str]:
        return self.validate_categorical(min_categories=1, min_series=1)

    def build_table(self) -> Table:
        return self.categorical_table()
