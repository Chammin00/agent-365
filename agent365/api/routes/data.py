"""
M365 Data API routes for direct inspection and operations.
"""

from fastapi import APIRouter, Query
from typing import Any, List, Optional
from datetime import datetime
from pydantic import BaseModel
from ...m365_client import get_m365_client
from ...models import (
    EmailMessage,
    DraftEmailRequest,
    SendEmailResponse,
    CalendarEvent,
    CreateMeetingRequest,
    ConflictReport,
    DriveItem,
    TeamsChannel,
    TeamsMessage,
    PostMessageRequest,
    ExcelTable,
    DocumentContent,
    PresentationContent,
)

router = APIRouter(prefix="/m365", tags=["Microsoft 365 Direct Data Services"])


# --- Outlook Mail Endpoints ---
@router.get("/inbox", response_model=List[EmailMessage])
def get_inbox(limit: int = Query(10, ge=1, le=50), unread_only: bool = False):
    client = get_m365_client()
    return client.get_inbox_messages(limit=limit, unread_only=unread_only)


@router.post("/send-mail", response_model=SendEmailResponse)
def send_email(request: DraftEmailRequest):
    client = get_m365_client()
    return client.send_email(request)


# --- Calendar Endpoints ---
@router.get("/calendar", response_model=List[CalendarEvent])
def get_calendar():
    client = get_m365_client()
    return client.get_events()


@router.post("/check-conflicts", response_model=ConflictReport)
def check_conflicts(start_time: str, end_time: str):
    client = get_m365_client()
    st = datetime.fromisoformat(start_time)
    et = datetime.fromisoformat(end_time)
    return client.check_conflicts(st, et)


@router.post("/schedule", response_model=CalendarEvent)
def schedule_event(request: CreateMeetingRequest):
    client = get_m365_client()
    return client.create_event(request)


# --- OneDrive Endpoints ---
@router.get("/files", response_model=List[DriveItem])
def get_files(folder_path: str = "/"):
    client = get_m365_client()
    return client.list_files(folder_path=folder_path)


# --- Teams Endpoints ---
@router.get("/teams/channels", response_model=List[TeamsChannel])
def get_teams_channels():
    client = get_m365_client()
    return client.list_channels()


@router.get("/teams/messages", response_model=List[TeamsMessage])
def get_teams_messages(channel_id: str):
    client = get_m365_client()
    return client.get_channel_messages(channel_id)


@router.post("/teams/post", response_model=TeamsMessage)
def post_teams_message(request: PostMessageRequest):
    client = get_m365_client()
    return client.post_channel_message(request)


# --- Office & Excel Endpoints ---
@router.get("/office/spreadsheet", response_model=ExcelTable)
def get_spreadsheet(file_id: str = "file-301"):
    client = get_m365_client()
    return client.get_spreadsheet_data(file_id=file_id)


class UpdateCellRequest(BaseModel):
    file_id: str = "file-301"
    sheet_name: str = "Q4_Summary"
    coordinate: str
    value: Any


@router.post("/office/cell")
def update_spreadsheet_cell(request: UpdateCellRequest):
    client = get_m365_client()
    success = client.update_cell_value(request.file_id, request.sheet_name, request.coordinate, request.value)
    return {"success": success, "coordinate": request.coordinate, "value": request.value}


# --- Word Endpoints ---
@router.get("/word/document", response_model=DocumentContent)
def get_word_document(file_id: str = "file-302"):
    client = get_m365_client()
    return client.get_word_document(file_id=file_id)


class SaveDocumentRequest(BaseModel):
    file_id: str = "file-302"
    content: str


@router.post("/word/save")
def save_word_document(request: SaveDocumentRequest):
    client = get_m365_client()
    success = client.save_word_document(request.file_id, request.content)
    return {"success": success, "file_id": request.file_id}


# --- PowerPoint Endpoints ---
@router.get("/powerpoint/presentation", response_model=PresentationContent)
def get_presentation(file_id: str = "file-304"):
    client = get_m365_client()
    return client.get_presentation(file_id=file_id)


# --- Mail Folders & Actions ---
@router.get("/mail/folder", response_model=List[EmailMessage])
def get_mail_by_folder(folder: str = "inbox", limit: int = 20):
    client = get_m365_client()
    return client.get_messages_by_folder(folder_name=folder, limit=limit)


class MailActionRequest(BaseModel):
    message_id: str


@router.post("/mail/delete")
def delete_mail(request: MailActionRequest):
    client = get_m365_client()
    success = client.delete_message(request.message_id)
    return {"success": success, "message_id": request.message_id}


@router.post("/mail/archive")
def archive_mail(request: MailActionRequest):
    client = get_m365_client()
    success = client.archive_message(request.message_id)
    return {"success": success, "message_id": request.message_id}


# --- OneNote Endpoints ---
@router.get("/onenote/notebooks")
def get_onenote_notebooks():
    client = get_m365_client()
    return client.get_onenote_notebooks()


class CreateOneNotePageRequest(BaseModel):
    section_id: str = "sec-01"
    title: str
    content: str = "<p>New page notes</p>"


@router.post("/onenote/page")
def create_onenote_page(request: CreateOneNotePageRequest):
    client = get_m365_client()
    page = client.create_onenote_page(request.section_id, request.title, request.content)
    return page


class CreateOneNoteSectionRequest(BaseModel):
    notebook_id: str = "nb-01"
    section_name: str


@router.post("/onenote/section")
def create_onenote_section(request: CreateOneNoteSectionRequest):
    client = get_m365_client()
    sec = client.create_onenote_section(request.notebook_id, request.section_name)
    return sec


class UpdateOneNotePageRequest(BaseModel):
    page_id: str
    title: str
    content: str


@router.post("/onenote/page/update")
def update_onenote_page(request: UpdateOneNotePageRequest):
    client = get_m365_client()
    success = client.update_onenote_page(request.page_id, request.title, request.content)
    return {"success": success, "page_id": request.page_id}


class CreateChannelRequest(BaseModel):
    team_id: str = "team-101"
    channel_name: str
    description: Optional[str] = None


@router.post("/teams/channel")
def create_teams_channel(request: CreateChannelRequest):
    client = get_m365_client()
    return client.create_teams_channel(request.team_id, request.channel_name, request.description)


class CreateFolderRequest(BaseModel):
    folder_name: str
    parent_folder: str = "/"


@router.post("/files/folder")
def create_folder(request: CreateFolderRequest):
    client = get_m365_client()
    return client.create_folder(request.folder_name, request.parent_folder)

