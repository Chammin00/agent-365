"""
Data models for Calendar entities.
"""

from datetime import datetime
from typing import List, Optional
from pydantic import BaseModel, Field


class Attendee(BaseModel):
    name: str
    email: str
    status: str = "needsAction"  # accepted, declined, tentativelyAccepted, needsAction


class CalendarEvent(BaseModel):
    id: str
    subject: str
    start_time: datetime
    end_time: datetime
    location: Optional[str] = "Microsoft Teams Meeting"
    is_online_meeting: bool = True
    online_meeting_url: Optional[str] = None
    organizer: str
    attendees: List[Attendee] = Field(default_factory=list)
    body_preview: str = ""
    is_all_day: bool = False
    importance: str = "normal"


class CreateMeetingRequest(BaseModel):
    subject: str
    start_time: datetime
    end_time: datetime
    attendees: List[str]
    location: Optional[str] = "Microsoft Teams Meeting"
    description: Optional[str] = ""


class ConflictReport(BaseModel):
    has_conflict: bool
    conflicting_events: List[CalendarEvent] = Field(default_factory=list)
    suggested_alternative_slots: List[str] = Field(default_factory=list)
    summary: str
