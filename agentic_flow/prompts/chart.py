"""
Prompt for the Chart Extractor — vision-based chart description.
"""

from __future__ import annotations

from typing import Optional


def get_chart_description_prompt(
    chart_title: Optional[str] = None,
    chart_type: Optional[str] = None,
    series_names: Optional[list] = None,
) -> str:
    context_parts = []
    if chart_title:
        context_parts.append(f"Title: {chart_title}")
    if chart_type:
        context_parts.append(f"Chart type: {chart_type}")
    if series_names:
        context_parts.append(f"Series: {', '.join(str(s) for s in series_names)}")
    context = "\n".join(context_parts) if context_parts else "(no metadata available)"

    return f"""You are a data analyst.  You are given a screenshot of a chart from an Excel worksheet.

Known metadata:
{context}

Describe this chart in 2-4 sentences.  Include:
- What the chart shows (topic / metric)
- The chart type (bar, line, pie, etc.)
- Key trends, comparisons, or notable data points visible
- Axis labels if readable

Output ONLY the description text, no JSON or formatting.
"""
