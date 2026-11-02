"""
Abstract base client for Microsoft 365 services.
Provides unified signatures for Outlook Mail, Calendar, OneDrive, Teams, and Excel.
"""

from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any, Dict, List, Optional
from ..models import (
    EmailMessage,
    DraftEmailRequest,
    SendEmailResponse,
    CalendarEvent,
    CreateMeetingRequest,
    ConflictReport,
    DriveItem,
    UploadFileRequest,
    ShareLinkResponse,
    TeamsChannel,
    TeamsMessage,
    PostMessageRequest,
    ExcelTable,
)


class M365ClientBase(ABC):
    """Abstract interface for Microsoft 365 operations."""

    # --- Mail Operations ---
    @abstractmethod
    def get_inbox_messages(self, limit: int = 10, unread_only: bool = False) -> List[EmailMessage]:
        pass

    @abstractmethod
    def get_message_by_id(self, message_id: str) -> Optional[EmailMessage]:
        pass

    @abstractmethod
    def send_email(self, request: DraftEmailRequest) -> SendEmailResponse:
        pass

    @abstractmethod
    def mark_as_read(self, message_id: str) -> bool:
        pass

    @abstractmethod
    def flag_message(self, message_id: str) -> bool:
        pass

    # --- Calendar Operations ---
    @abstractmethod
    def get_events(self, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None) -> List[CalendarEvent]:
        pass

    @abstractmethod
    def create_event(self, request: CreateMeetingRequest) -> CalendarEvent:
        pass

    @abstractmethod
    def check_conflicts(self, start_time: datetime, end_time: datetime) -> ConflictReport:
        pass

    # --- OneDrive / SharePoint Storage ---
    @abstractmethod
    def list_files(self, folder_path: str = "/") -> List[DriveItem]:
        pass

    @abstractmethod
    def get_file_content(self, file_id: str) -> str:
        pass

    @abstractmethod
    def upload_file(self, request: UploadFileRequest) -> DriveItem:
        pass

    @abstractmethod
    def create_share_link(self, file_id: str) -> ShareLinkResponse:
        pass

    # --- Microsoft Teams ---
    @abstractmethod
    def list_channels(self) -> List[TeamsChannel]:
        pass

    @abstractmethod
    def get_channel_messages(self, channel_id: str, limit: int = 10) -> List[TeamsMessage]:
        pass

    @abstractmethod
    def post_channel_message(self, request: PostMessageRequest) -> TeamsMessage:
        pass

    # --- Office / Excel / Word / PowerPoint ---
    @abstractmethod
    def get_spreadsheet_data(self, file_id: str, sheet_name: Optional[str] = None) -> ExcelTable:
        pass

    @abstractmethod
    def update_cell_value(self, file_id: str, sheet_name: str, coordinate: str, value: Any) -> bool:
        pass

    @abstractmethod
    def get_word_document(self, file_id: str = "file-302") -> Any:
        pass

    @abstractmethod
    def save_word_document(self, file_id: str, content: str) -> bool:
        pass

    # --- Mail Folders & Actions ---
    @abstractmethod
    def get_messages_by_folder(self, folder_name: str = "inbox", limit: int = 20) -> List[EmailMessage]:
        pass

    @abstractmethod
    def delete_message(self, message_id: str) -> bool:
        pass

    @abstractmethod
    def archive_message(self, message_id: str) -> bool:
        pass

    # --- OneNote ---
    @abstractmethod
    def get_onenote_notebooks(self) -> List[Any]:
        pass

    @abstractmethod
    def create_onenote_page(self, section_id: str, title: str, content: str) -> Any:
        pass

    @abstractmethod
    def create_onenote_section(self, notebook_id: str, section_name: str) -> Any:
        pass

    @abstractmethod
    def update_onenote_page(self, page_id: str, title: str, content: str) -> bool:
        pass

    # --- Teams Channels & Folders ---
    @abstractmethod
    def create_teams_channel(self, team_id: str, channel_name: str, description: Optional[str] = None) -> Any:
        pass

    @abstractmethod
    def create_folder(self, folder_name: str, parent_folder: str = "/") -> Any:
        pass

