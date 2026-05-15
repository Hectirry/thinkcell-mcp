"""Combination (bar + line) chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table


class ComboChart(ChartBuilder):
    """A combination chart pairing bar series with line series.

    Whether a given series renders as a bar or a line is configured on the
    think-cell chart inside the PowerPoint template. Pass the optional
    ``line_series`` list of series names to document which series are intended
    as lines; those names are validated against the data.
    """

    chart_type = "combo"
    data_shape = (
        "List of row objects with one category key ('category', 'label' or "
        "'name') and at least two numeric series keys -- typically a volume "
        "series shown as bars and a rate/percentage series shown as a line. "
        "Optional 'line_series' option lists which series are lines."
    )
    example = {
        "chart_type": "combo",
        "title": "Revenue and Margin",
        "data": [
            {"category": "Q1", "Revenue": 120, "Margin %": 18},
            {"category": "Q2", "Revenue": 150, "Margin %": 22},
            {"category": "Q3", "Revenue": 165, "Margin %": 25},
        ],
    }

    known_option_keys = ChartBuilder.known_option_keys | {"line_series"}

    def _validate(self) -> list[str]:
        errors = self.validate_categorical(min_categories=1, min_series=2)
        line_series = self.options.get("line_series")
        if line_series is not None:
            if not isinstance(line_series, list):
                errors.append(
                    "`line_series` option must be a list of series names"
                )
            else:
                available = set(self.series_names(self.category_key()))
                for name in line_series:
                    if name not in available:
                        errors.append(
                            f"`line_series` entry '{name}' is not a series "
                            f"present in the data"
                        )
        return errors

    def build_table(self) -> Table:
        return self.categorical_table()
