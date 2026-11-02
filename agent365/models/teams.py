"""
Data models for Microsoft Teams entities.
"""

from datetime import datetime
from typing import List, Optional
from pydantic import BaseModel, Field


class TeamsChannel(BaseModel):
    id: str
    team_name: str
    channel_name: str
    description: Optional[str] = ""


class TeamsMessage(BaseModel):
    id: str
    channel_id: Optional[str] = None
    chat_id: Optional[str] = None
    sender_name: str
    content: str
    timestamp: datetime
    importance: str = "normal"  # normal, urgent
    mentions: List[str] = Field(default_factory=list)


class PostMessageRequest(BaseModel):
    target_type: str = "channel"  # 'channel' or 'chat'
    target_id: str
    content: str
    importance: str = "normal"
