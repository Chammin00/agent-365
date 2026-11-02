"""
Data models for Office Documents, Excel Workbooks, PowerPoint Presentations, and OneNote.
"""

from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


class ExcelCell(BaseModel):
    coordinate: str  # e.g., "A1", "B4"
    value: Any
    formula: Optional[str] = None


class ExcelTable(BaseModel):
    sheet_name: str
    headers: List[str]
    rows: List[List[Any]]
    summary: Optional[Dict[str, Any]] = Field(default_factory=dict)


class DocumentParagraph(BaseModel):
    heading_level: int = 0  # 0 for normal text, 1 for Title, 2 for H2, etc.
    text: str


class DocumentContent(BaseModel):
    title: str
    author: str
    raw_content: Optional[str] = ""
    sections: List[DocumentParagraph] = Field(default_factory=list)


class PowerPointSlide(BaseModel):
    slide_number: int
    title: str
    subtitle: Optional[str] = ""
    bullets: List[str] = Field(default_factory=list)
    speaker_notes: Optional[str] = ""


class PresentationContent(BaseModel):
    title: str
    author: str
    slides: List[PowerPointSlide] = Field(default_factory=list)


class OneNotePage(BaseModel):
    id: str
    title: str
    created_at: str
    content: str


class OneNoteSection(BaseModel):
    id: str
    name: str
    color: str = "#7719aa"
    pages: List[OneNotePage] = Field(default_factory=list)


class OneNoteNotebook(BaseModel):
    id: str
    name: str
    sections: List[OneNoteSection] = Field(default_factory=list)
