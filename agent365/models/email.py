"""
Data models for Outlook Mail entities.
"""

from datetime import datetime
from typing import List, Optional
from pydantic import BaseModel, Field


class EmailAddress(BaseModel):
    name: str
    address: str


class EmailMessage(BaseModel):
    id: str
    subject: str
    sender: EmailAddress
    to_recipients: List[EmailAddress] = Field(default_factory=list)
    cc_recipients: List[EmailAddress] = Field(default_factory=list)
    body_preview: str = ""
    body_content: str = ""
    received_at: datetime
    is_read: bool = False
    is_flagged: bool = False
    importance: str = "normal"  # low, normal, high
    has_attachments: bool = False
    categories: List[str] = Field(default_factory=list)


class DraftEmailRequest(BaseModel):
    to: List[str]
    subject: str
    body: str
    cc: List[str] = Field(default_factory=list)
    importance: str = "normal"


class SendEmailResponse(BaseModel):
    success: bool
    message_id: str
    detail: str
