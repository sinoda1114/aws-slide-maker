"""Pydantic models for slide specification (YAML/JSON schema)."""
from __future__ import annotations

from typing import Annotated, List, Literal, Optional, Union

from pydantic import BaseModel, Field


class TitleSlide(BaseModel):
    type: Literal["title"]
    title: str
    subtitle: Optional[str] = None


class AgendaSlide(BaseModel):
    type: Literal["agenda"]
    title: str = "アジェンダ"
    items: List[str]


class ContentSlide(BaseModel):
    type: Literal["content"]
    title: str
    body: Optional[str] = None
    bullets: Optional[List[str]] = None


class ColumnContent(BaseModel):
    title: Optional[str] = None
    body: Optional[str] = None
    bullets: Optional[List[str]] = None


class TwoColumnSlide(BaseModel):
    type: Literal["two_column"]
    title: str
    left: ColumnContent
    right: ColumnContent


class ClosingSlide(BaseModel):
    type: Literal["closing"]
    title: str = "ご清聴ありがとうございました"
    message: Optional[str] = None


AnySlide = Annotated[
    Union[TitleSlide, AgendaSlide, ContentSlide, TwoColumnSlide, ClosingSlide],
    Field(discriminator="type"),
]


class PresentationSpec(BaseModel):
    title: str
    subtitle: Optional[str] = None
    author: Optional[str] = None
    theme: str = "aws"
    slides: List[AnySlide]
