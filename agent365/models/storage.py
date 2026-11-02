"""
Data models for OneDrive and SharePoint storage items.
"""

from datetime import datetime
from typing import Optional
from pydantic import BaseModel


class DriveItem(BaseModel):
    id: str
    name: str
    web_url: str
    size_bytes: int
    created_at: datetime
    modified_at: datetime
    is_folder: bool = False
    parent_path: str = "/"
    mime_type: Optional[str] = "application/octet-stream"
    summary_preview: Optional[str] = None


class UploadFileRequest(BaseModel):
    name: str
    content: str
    parent_path: str = "/"


class ShareLinkResponse(BaseModel):
    success: bool
    item_id: str
    share_url: str
    permission: str = "view"
