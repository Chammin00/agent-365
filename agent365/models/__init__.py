"""
Models package for Agent 365.
"""

from .email import EmailAddress, EmailMessage, DraftEmailRequest, SendEmailResponse
from .calendar import CalendarEvent, Attendee, CreateMeetingRequest, ConflictReport
from .storage import DriveItem, UploadFileRequest, ShareLinkResponse
from .teams import TeamsChannel, TeamsMessage, PostMessageRequest
from .office import (
    ExcelTable,
    ExcelCell,
    DocumentContent,
    DocumentParagraph,
    PowerPointSlide,
    PresentationContent,
    OneNotePage,
    OneNoteSection,
    OneNoteNotebook,
)
from .agent_state import AgentRole, AgentStep, PlanStep, OrchestratorResponse

__all__ = [
    "EmailAddress",
    "EmailMessage",
    "DraftEmailRequest",
    "SendEmailResponse",
    "CalendarEvent",
    "Attendee",
    "CreateMeetingRequest",
    "ConflictReport",
    "DriveItem",
    "UploadFileRequest",
    "ShareLinkResponse",
    "TeamsChannel",
    "TeamsMessage",
    "PostMessageRequest",
    "ExcelTable",
    "ExcelCell",
    "DocumentContent",
    "DocumentParagraph",
    "PowerPointSlide",
    "PresentationContent",
    "OneNotePage",
    "OneNoteSection",
    "OneNoteNotebook",
    "AgentRole",
    "AgentStep",
    "PlanStep",
    "OrchestratorResponse",
]
