"""Chart type mapping from DSL names to openpyxl chart classes."""

from __future__ import annotations

from openpyxl.chart import (
    AreaChart,
    AreaChart3D,
    BarChart,
    BarChart3D,
    BubbleChart,
    DoughnutChart,
    LineChart,
    LineChart3D,
    PieChart,
    PieChart3D,
    RadarChart,
    ScatterChart,
    StockChart,
    SurfaceChart,
    SurfaceChart3D,
)

# Mapping from DSL chart type name to (class, grouping, type_override)
CHART_TYPES: dict[str, tuple[type, str | None, str | None]] = {
    "bar": (BarChart, "clustered", "bar"),
    "column": (BarChart, "clustered", "col"),
    "line": (LineChart, None, None),
    "pie": (PieChart, None, None),
    "scatter": (ScatterChart, None, None),
    "area": (AreaChart, None, None),
    "doughnut": (DoughnutChart, None, None),
    "radar": (RadarChart, None, None),
    "bubble": (BubbleChart, None, None),
    "stock": (StockChart, None, None),
    "surface": (SurfaceChart, None, None),
    # Stacked variants
    "stacked-bar": (BarChart, "stacked", "bar"),
    "stacked-column": (BarChart, "stacked", "col"),
    "stacked-area": (AreaChart, "stacked", None),
    # 100% stacked
    "100-bar": (BarChart, "percentStacked", "bar"),
    "100-column": (BarChart, "percentStacked", "col"),
    "100-area": (AreaChart, "percentStacked", None),
    # 3D variants
    "bar-3d": (BarChart3D, "clustered", "bar"),
    "column-3d": (BarChart3D, "clustered", "col"),
    "line-3d": (LineChart3D, None, None),
    "pie-3d": (PieChart3D, None, None),
    "area-3d": (AreaChart3D, None, None),
    "surface-3d": (SurfaceChart3D, None, None),
}


def get_chart_class(type_name: str) -> tuple[type, str | None, str | None]:
    """Look up chart class and config for a DSL type name.

    Returns (chart_class, grouping, type_override).
    Raises ValueError if type not found.
    """
    key = type_name.lower()
    if key not in CHART_TYPES:
        available = ", ".join(sorted(CHART_TYPES.keys()))
        raise ValueError(f"Unknown chart type: {type_name!r}. Available: {available}")
    return CHART_TYPES[key]
