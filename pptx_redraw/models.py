"""Data models for AegisOps slide content."""

from __future__ import annotations

from pydantic import BaseModel, Field


class SectionCard(BaseModel):
    """Section card content in the central MVP container.

    Args:
        title: Card title.
        bullets: Card bullet rows.

    Returns:
        None.

    Raises:
        ValueError: If title or bullets violate constraints.
    """

    title: str = Field(min_length=1)
    bullets: list[str] = Field(min_length=1)


class ObjectiveItem(BaseModel):
    """Right panel MVP focus objective row.

    Args:
        name: Objective name.
        description: Objective description.

    Returns:
        None.

    Raises:
        ValueError: If fields are empty.
    """

    name: str = Field(min_length=1)
    description: str = Field(min_length=1)


class RoadmapQuarter(BaseModel):
    """Quarter roadmap block.

    Args:
        quarter: Quarter title.
        summary: Quarter summary.
        bullets: Quarter tasks.

    Returns:
        None.

    Raises:
        ValueError: If required fields are empty.
    """

    quarter: str = Field(min_length=1)
    summary: str = Field(min_length=1)
    bullets: list[str] = Field(min_length=1)


class SlideContent(BaseModel):
    """Whole single-slide content schema.

    Args:
        title: Main title.
        subtitle: Subtitle.
        badge: Goal badge text.
        control_sections: Core MVP section cards.
        objectives: MVP focus objectives.
        roadmap: Quarterly roadmap data.

    Returns:
        None.

    Raises:
        ValueError: If content lists are invalid.
    """

    title: str
    subtitle: str
    badge: str
    control_sections: list[SectionCard] = Field(min_length=6, max_length=6)
    objectives: list[ObjectiveItem] = Field(min_length=7, max_length=7)
    roadmap: list[RoadmapQuarter] = Field(min_length=4, max_length=4)
