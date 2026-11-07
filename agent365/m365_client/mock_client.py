"""
Rich simulated Microsoft 365 Sandbox Client.
Provides realistic enterprise data and stateful in-memory operations for testing and offline demos.
"""

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
import uuid

from .base import M365ClientBase
from ..models import (
    EmailAddress,
    EmailMessage,
    DraftEmailRequest,
    SendEmailResponse,
    CalendarEvent,
    Attendee,
    CreateMeetingRequest,
    ConflictReport,
    DriveItem,
    UploadFileRequest,
    ShareLinkResponse,
    TeamsChannel,
    TeamsMessage,
    PostMessageRequest,
    ExcelTable,
    DocumentContent,
    DocumentParagraph,
    PowerPointSlide,
    PresentationContent,
    OneNotePage,
    OneNoteSection,
    OneNoteNotebook,
)


class MockM365Client(M365ClientBase):
    """In-memory simulated M365 tenant with fresh out-of-the-box state and optional demo seed."""

    def __init__(self, fresh: bool = False):
        self.fresh = fresh
        now = datetime.now()
        if fresh:
            self.init_fresh_state(now)
        else:
            self.seed_demo_dataset(now)

    def init_fresh_state(self, now: Optional[datetime] = None):
        """Initializes a pristine, fresh out-of-the-box tenant with zero pre-written example text."""
        if now is None:
            now = datetime.now()
        self.emails: List[EmailMessage] = []
        self.events: List[CalendarEvent] = []
        self.files: List[DriveItem] = []
        self.channels: List[TeamsChannel] = [
            TeamsChannel(id="chan-401", team_name="Enterprise 365", channel_name="General", description="General discussions and updates")
        ]
        self.messages: List[TeamsMessage] = []
        self.spreadsheets: Dict[str, ExcelTable] = {
            "file-301": ExcelTable(
                sheet_name="Sheet1",
                headers=["A", "B", "C", "D", "E", "F", "G"],
                rows=[["" for _ in range(7)] for _ in range(25)],
                summary={"status": "Ready"}
            )
        }
        self.presentations: Dict[str, PresentationContent] = {
            "file-304": PresentationContent(
                title="Presentation1.pptx",
                author="User",
                slides=[
                    PowerPointSlide(
                        slide_number=1,
                        title="",
                        subtitle="",
                        bullets=[]
                    )
                ]
            )
        }
        self.word_documents: Dict[str, DocumentContent] = {
            "file-302": DocumentContent(
                title="Document1.docx",
                author="User",
                raw_content="",
                sections=[]
            )
        }
        self.folders: Dict[str, List[EmailMessage]] = {
            "inbox": self.emails,
            "sent": [],
            "drafts": [],
            "archive": [],
            "deleted": [],
        }
        self.onenote_notebooks: List[OneNoteNotebook] = [
            OneNoteNotebook(
                id="nb-01",
                name="My Notebook",
                sections=[
                    OneNoteSection(
                        id="sec-01",
                        name="Quick Notes",
                        color="#7719aa",
                        pages=[
                            OneNotePage(
                                id="pg-01",
                                title="Untitled page",
                                created_at=now.strftime("%Y-%m-%d"),
                                content=""
                            )
                        ]
                    )
                ]
            )
        ]

    def seed_demo_dataset(self, now: Optional[datetime] = None):
        """Seeds benchmark demo data for test suites."""
        if now is None:
            now = datetime.now()

        # Seed Emails
        self.emails: List[EmailMessage] = [
            EmailMessage(
                id="msg-101",
                subject="URGENT: Q4 Revenue Targets & Executive Briefing",
                sender=EmailAddress(name="Sarah Jenkins (CEO)", address="sarah.jenkins@enterprise365.com"),
                to_recipients=[EmailAddress(name="Executive Team", address="exec-team@enterprise365.com")],
                body_preview="Please prepare the Q4 variance analysis and budget reconciliations before our 2:00 PM review.",
                body_content="Team,\n\nPlease ensure the Q4 variance analysis and department budget reconciliations are finalized in the OneDrive financial workbook before our 2:00 PM review today. Key focus areas: Cloud infrastructure spend vs forecast and Enterprise Sales margin.\n\nBest,\nSarah",
                received_at=now - timedelta(hours=1, minutes=15),
                is_read=False,
                is_flagged=True,
                importance="high",
                categories=["Executive", "Financials"],
            ),
            EmailMessage(
                id="msg-102",
                subject="Project Sentinel: Sprint 42 Release Readiness",
                sender=EmailAddress(name="Marcus Chen (VP Eng)", address="marcus.chen@enterprise365.com"),
                to_recipients=[EmailAddress(name="Engineering Leads", address="eng-leads@enterprise365.com")],
                body_preview="All 14 integration test suites passed. Production deployment scheduled for 21:00 UTC.",
                body_content="Hi everyone,\n\nAll 14 integration test suites passed in staging. The rollback plan has been verified and documented. Deployment window starts tonight at 21:00 UTC.\n\nRegards,\nMarcus",
                received_at=now - timedelta(hours=3),
                is_read=False,
                is_flagged=False,
                importance="normal",
                categories=["Engineering"],
            ),
            EmailMessage(
                id="msg-103",
                subject="Action Required: Annual Cyber Security Compliance Audit",
                sender=EmailAddress(name="SecOps Governance", address="compliance@enterprise365.com"),
                to_recipients=[EmailAddress(name="All Staff", address="all-staff@enterprise365.com")],
                body_preview="All employees must complete the 2026 Entra ID zero-trust verification module by Friday.",
                body_content="Reminder to complete your annual Zero-Trust security review and acknowledge the updated acceptable use policies.",
                received_at=now - timedelta(hours=6),
                is_read=True,
                is_flagged=False,
                importance="normal",
                categories=["Compliance"],
            ),
            EmailMessage(
                id="msg-104",
                subject="Client Proposal: Northwind Traders Cloud Modernization",
                sender=EmailAddress(name="Elena Rostova (Account Dir)", address="elena.r@enterprise365.com"),
                to_recipients=[EmailAddress(name="Solutions Team", address="solutions@enterprise365.com")],
                body_preview="Northwind has requested an updated timeline and revised SLA agreement.",
                body_content="Northwind leadership has requested that we update our SLA terms from 99.9% to 99.99%. Let's review the revised pricing model.",
                received_at=now - timedelta(days=1),
                is_read=True,
                is_flagged=True,
                importance="high",
                categories=["Sales"],
            ),
        ]

        # Seed Calendar Events
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        self.events: List[CalendarEvent] = [
            CalendarEvent(
                id="evt-201",
                subject="Daily Engineering Standup",
                start_time=today_start + timedelta(hours=9, minutes=30),
                end_time=today_start + timedelta(hours=10, minutes=0),
                organizer="Marcus Chen",
                attendees=[
                    Attendee(name="Marcus Chen", email="marcus.chen@enterprise365.com", status="accepted"),
                    Attendee(name="You", email="user@enterprise365.com", status="accepted"),
                ],
                body_preview="Daily blocker check and sprint milestone synchronization.",
            ),
            CalendarEvent(
                id="evt-202",
                subject="Executive Q4 Budget & Revenue Review",
                start_time=today_start + timedelta(hours=14, minutes=0),
                end_time=today_start + timedelta(hours=15, minutes=0),
                organizer="Sarah Jenkins (CEO)",
                attendees=[
                    Attendee(name="Sarah Jenkins", email="sarah.jenkins@enterprise365.com", status="accepted"),
                    Attendee(name="Elena Rostova", email="elena.r@enterprise365.com", status="accepted"),
                    Attendee(name="You", email="user@enterprise365.com", status="accepted"),
                ],
                body_preview="Review Q4 department budgets and financial model projections.",
            ),
            CalendarEvent(
                id="evt-203",
                subject="Client Partner Sync: Northwind Traders",
                start_time=today_start + timedelta(hours=14, minutes=30),
                end_time=today_start + timedelta(hours=15, minutes=30),
                organizer="Elena Rostova",
                attendees=[
                    Attendee(name="Elena Rostova", email="elena.r@enterprise365.com", status="accepted"),
                    Attendee(name="You", email="user@enterprise365.com", status="needsAction"),
                ],
                body_preview="Conflicting meeting: Discussing updated SLA requirements.",
            ),
            CalendarEvent(
                id="evt-204",
                subject="Product Architecture Working Group",
                start_time=today_start + timedelta(hours=16, minutes=0),
                end_time=today_start + timedelta(hours=17, minutes=0),
                organizer="Architecture Board",
                attendees=[
                    Attendee(name="Architecture Board", email="arch@enterprise365.com", status="accepted"),
                ],
                body_preview="Agentic workflows and MCP integration deep-dive.",
            ),
        ]

        # Seed OneDrive Items
        self.files: List[DriveItem] = [
            DriveItem(
                id="file-301",
                name="Q4_Financial_Projections.xlsx",
                web_url="https://enterprise365.sharepoint.com/sites/Finance/Shared%20Documents/Q4_Financial_Projections.xlsx",
                size_bytes=425000,
                created_at=now - timedelta(days=12),
                modified_at=now - timedelta(hours=2),
                is_folder=False,
                mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                summary_preview="Department financial projections, headcount allocations, cloud operational expenditures, and revenue variances for FY26-Q4.",
            ),
            DriveItem(
                id="file-302",
                name="Executive_Strategic_Plan_2026.docx",
                web_url="https://enterprise365.sharepoint.com/sites/Executive/Shared%20Documents/Executive_Strategic_Plan_2026.docx",
                size_bytes=1150000,
                created_at=now - timedelta(days=30),
                modified_at=now - timedelta(days=3),
                is_folder=False,
                mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                summary_preview="Comprehensive company roadmap covering Agentic AI transformation, cloud infrastructure modernization, and enterprise margin expansion.",
            ),
            DriveItem(
                id="file-303",
                name="Enterprise_Security_Audit_Report.pdf",
                web_url="https://enterprise365.sharepoint.com/sites/Security/Shared%20Documents/Enterprise_Security_Audit_Report.pdf",
                size_bytes=2450000,
                created_at=now - timedelta(days=5),
                modified_at=now - timedelta(days=5),
                is_folder=False,
                mime_type="application/pdf",
                summary_preview="Zero-Trust posture audit, SOC 2 Type II compliance scorecard, and Entra ID governance audit findings.",
            ),
        ]

        # Seed Teams Channels & Messages
        self.channels: List[TeamsChannel] = [
            TeamsChannel(id="chan-401", team_name="Enterprise 365", channel_name="General", description="Company-wide announcements and updates"),
            TeamsChannel(id="chan-402", team_name="Engineering Core", channel_name="Releases", description="Automated build notifications and release status"),
            TeamsChannel(id="chan-403", team_name="Executive Leadership", channel_name="Strategy", description="High-priority strategic discussions"),
        ]

        self.messages: List[TeamsMessage] = [
            TeamsMessage(
                id="tmsg-501",
                channel_id="chan-401",
                sender_name="Sarah Jenkins (CEO)",
                content="Welcome everyone to the start of Q4! Please check the shared OneDrive folder for strategic priorities.",
                timestamp=now - timedelta(days=2),
            ),
            TeamsMessage(
                id="tmsg-502",
                channel_id="chan-402",
                sender_name="Marcus Chen (VP Eng)",
                content="Sprint 42 is in green status. Staging tests passed 100%. Ready for tonight's rollout.",
                timestamp=now - timedelta(hours=2),
                importance="normal",
            ),
            TeamsMessage(
                id="tmsg-503",
                channel_id="chan-403",
                sender_name="Elena Rostova",
                content="Reminder: Northwind Traders decision committee meets tomorrow. Ensure SLA terms are aligned.",
                timestamp=now - timedelta(hours=4),
                importance="urgent",
                mentions=["Executive Team"],
            ),
        ]

        # Seed Excel Spreadsheet Content
        self.spreadsheets: Dict[str, ExcelTable] = {
            "file-301": ExcelTable(
                sheet_name="Q4_Summary",
                headers=["Department", "Q4 Budget ($k)", "Q4 Actual ($k)", "Variance ($k)", "Status"],
                rows=[
                    ["Cloud Infrastructure", 1250, 1180, -70, "Under Budget (Favorable)"],
                    ["Engineering R&D", 3400, 3350, -50, "Under Budget (Favorable)"],
                    ["Sales & Marketing", 2100, 2220, +120, "Over Budget (Caution)"],
                    ["Customer Success", 850, 840, -10, "On Target"],
                    ["General & Administrative", 650, 645, -5, "On Target"],
                ],
                summary={
                    "total_budget_k": 8250,
                    "total_actual_k": 8235,
                    "net_variance_k": -15,
                    "overall_status": "Healthy / Under Budget by $15k",
                },
            )
        }

        # Seed PowerPoint Presentation
        self.presentations: Dict[str, PresentationContent] = {
            "file-304": PresentationContent(
                title="Enterprise 365 — Q4 Strategic Outlook",
                author="Sarah Jenkins (CEO)",
                slides=[
                    PowerPointSlide(
                        slide_number=1,
                        title="Enterprise 365 — Q4 Strategic Outlook",
                        subtitle="Accelerating Cloud Modernization & Autonomous Agentic Operations",
                        bullets=[
                            "FY26 Executive Briefing & Department Goals",
                            "Confidential — Internal Board Review",
                        ],
                        speaker_notes="Welcome board members. Today we review our Q4 growth vectors and margin targets.",
                    ),
                    PowerPointSlide(
                        slide_number=2,
                        title="Financial Performance & Variance",
                        subtitle="Q4 Operational Expenditure Under Budget by $15k",
                        bullets=[
                            "Cloud Infrastructure: $1,180k actual vs $1,250k budget (-$70k favorable)",
                            "Engineering R&D: $3,350k actual vs $3,400k budget (-$50k favorable)",
                            "Sales & Marketing: $2,220k actual vs $2,100k budget (+$120k expansion)",
                            "Overall tenant operating margin expanded by 2.4%",
                        ],
                        speaker_notes="Highlight that infrastructure optimization freed up budget for accelerated AI agent deployment.",
                    ),
                    PowerPointSlide(
                        slide_number=3,
                        title="Autonomous Multi-Agent AI Deployment",
                        subtitle="Agent 365 Architecture Roadmap",
                        bullets=[
                            "Specialized Agents for Mail, Calendar, OneDrive, Teams, and Excel",
                            "Model Context Protocol (MCP) Enterprise integration",
                            "Zero-trust security alignment with Microsoft Entra ID governance",
                            "Projected productivity uplift: 4.8 hours saved per executive weekly",
                        ],
                        speaker_notes="Demonstrate the live Chief of Staff orchestrator and response times.",
                    ),
                ],
            )
        }

        # Seed Word Document
        self.word_documents: Dict[str, DocumentContent] = {
            "file-302": DocumentContent(
                title="Executive_Strategic_Plan_2026.docx",
                author="Sarah Jenkins (CEO)",
                raw_content="Executive Strategic Plan 2026\n\n1. Vision & Core Mission\nEnterprise 365 leads the market in autonomous enterprise workflow transformation. In FY26, our primary objective is unifying cloud data pipelines with specialized AI agents across Microsoft 365.\n\n2. Key Priorities\n- Secure Entra ID Zero-Trust architecture across all cloud endpoints.\n- Automate routine financial variance calculations and calendar de-confliction.\n- Modernize collaboration channels in Microsoft Teams with real-time telemetry.\n\n3. Governance & Risk Management\nAll autonomous operations adhere strictly to SOC 2 Type II compliance, Purview audit logging, and role-based access control (RBAC).",
                sections=[
                    DocumentParagraph(heading_level=1, text="Executive Strategic Plan 2026"),
                    DocumentParagraph(heading_level=2, text="1. Vision & Core Mission"),
                    DocumentParagraph(heading_level=0, text="Enterprise 365 leads the market in autonomous enterprise workflow transformation. In FY26, our primary objective is unifying cloud data pipelines with specialized AI agents across Microsoft 365."),
                    DocumentParagraph(heading_level=2, text="2. Key Priorities"),
                    DocumentParagraph(heading_level=0, text="• Secure Entra ID Zero-Trust architecture across all cloud endpoints.\n• Automate routine financial variance calculations and calendar de-confliction.\n• Modernize collaboration channels in Microsoft Teams with real-time telemetry."),
                    DocumentParagraph(heading_level=2, text="3. Governance & Risk Management"),
                    DocumentParagraph(heading_level=0, text="All autonomous operations adhere strictly to SOC 2 Type II compliance, Purview audit logging, and role-based access control (RBAC)."),
                ],
            )
        }

        # Add PowerPoint to files list
        self.files.append(
            DriveItem(
                id="file-304",
                name="Enterprise_365_Strategy.pptx",
                web_url="https://enterprise365.sharepoint.com/sites/Executive/Shared%20Documents/Enterprise_365_Strategy.pptx",
                size_bytes=1840000,
                created_at=now - timedelta(days=8),
                modified_at=now - timedelta(hours=5),
                is_folder=False,
                mime_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                summary_preview="Q4 Executive board presentation covering financial performance, margin analysis, and Agent 365 deployment roadmap.",
            )
        )

        # Seed Outlook Folders (Inbox, Sent, Drafts, Archive, Deleted)
        self.folders: Dict[str, List[EmailMessage]] = {
            "inbox": self.emails,
            "sent": [
                EmailMessage(
                    id="msg-sent-01",
                    subject="Confirmed: Q4 Infrastructure Budget Approval",
                    sender=EmailAddress(name="You", address="user@enterprise365.com"),
                    to_recipients=[EmailAddress(name="Marcus Chen", address="marcus.chen@enterprise365.com")],
                    body_preview="Approved. The cloud spend reduction targets have been locked.",
                    body_content="Marcus,\n\nThe cloud spend optimization forecast has been approved. Proceed with the deployment window tonight.\n\nBest,\nYou",
                    received_at=now - timedelta(hours=4),
                    is_read=True,
                )
            ],
            "drafts": [
                EmailMessage(
                    id="msg-draft-01",
                    subject="Draft: FY27 AI Copilot Expansion Proposal",
                    sender=EmailAddress(name="You", address="user@enterprise365.com"),
                    to_recipients=[EmailAddress(name="Board of Directors", address="board@enterprise365.com")],
                    body_preview="Outlining our multi-agent autonomous enterprise governance model...",
                    body_content="Board of Directors,\n\nAttached is our FY27 strategy for scaling autonomous agent infrastructure across all tenant operations.",
                    received_at=now - timedelta(days=1),
                    is_read=True,
                )
            ],
            "archive": [],
            "deleted": [],
        }

        # Seed OneNote Notebooks, Sections, and Pages
        self.onenote_notebooks: List[OneNoteNotebook] = [
            OneNoteNotebook(
                id="nb-01",
                name="Executive Strategy Notebook",
                sections=[
                    OneNoteSection(
                        id="sec-01",
                        name="Q4 Directives",
                        color="#7719aa",
                        pages=[
                            OneNotePage(
                                id="pg-01",
                                title="FY26 Strategic OKRs",
                                created_at=(now - timedelta(days=10)).strftime("%Y-%m-%d"),
                                content="<h3>FY26 Strategic OKRs</h3><p><strong>Objective 1:</strong> Scale autonomous agent pipelines across all M365 tenants.</p><p><strong>Key Result 1.1:</strong> Reduce administrative triage latency to &lt; 5 minutes.</p><p><strong>Key Result 1.2:</strong> Maintain 100% Zero-Trust Entra ID compliance score.</p>",
                            ),
                            OneNotePage(
                                id="pg-02",
                                title="Infrastructure Cost Controls",
                                created_at=(now - timedelta(days=5)).strftime("%Y-%m-%d"),
                                content="<h3>Infrastructure Cost Controls</h3><p>Cloud spend optimization targets achieved. $70k favorable variance documented in Q4_Financial_Projections.xlsx.</p>",
                            ),
                        ],
                    ),
                    OneNoteSection(
                        id="sec-02",
                        name="Board Prep",
                        color="#0078d4",
                        pages=[
                            OneNotePage(
                                id="pg-03",
                                title="Board Review Talking Points",
                                created_at=(now - timedelta(days=2)).strftime("%Y-%m-%d"),
                                content="<h3>Board Review Talking Points</h3><p>&bull; Demonstrate Agent 365 Chief of Staff live.</p><p>&bull; Review SLA term adjustments for Northwind Traders.</p>",
                            )
                        ],
                    ),
                ],
            )
        ]

    # --- Mail Operations ---
    def get_inbox_messages(self, limit: int = 10, unread_only: bool = False) -> List[EmailMessage]:
        msgs = self.emails
        if unread_only:
            msgs = [m for m in msgs if not m.is_read]
        return sorted(msgs, key=lambda m: m.received_at, reverse=True)[:limit]

    def get_message_by_id(self, message_id: str) -> Optional[EmailMessage]:
        for m in self.emails:
            if m.id == message_id:
                return m
        return None

    def send_email(self, request: DraftEmailRequest) -> SendEmailResponse:
        new_msg = EmailMessage(
            id=f"msg-{uuid.uuid4().hex[:6]}",
            subject=request.subject,
            sender=EmailAddress(name="Agent 365 Executive Copilot", address="copilot@enterprise365.com"),
            to_recipients=[EmailAddress(name=addr.split("@")[0].capitalize(), address=addr) for addr in request.to],
            body_preview=request.body[:120],
            body_content=request.body,
            received_at=datetime.now(),
            is_read=True,
            is_flagged=False,
            importance=request.importance,
        )
        self.emails.insert(0, new_msg)
        if hasattr(self, "folders") and "sent" in self.folders:
            self.folders["sent"].insert(0, new_msg)
        return SendEmailResponse(success=True, message_id=new_msg.id, detail=f"Email successfully sent to {', '.join(request.to)}")

    def mark_as_read(self, message_id: str) -> bool:
        msg = self.get_message_by_id(message_id)
        if msg:
            msg.is_read = True
            return True
        return False

    def flag_message(self, message_id: str) -> bool:
        msg = self.get_message_by_id(message_id)
        if msg:
            msg.is_flagged = True
            return True
        return False

    # --- Calendar Operations ---
    def get_events(self, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None) -> List[CalendarEvent]:
        events = sorted(self.events, key=lambda e: e.start_time)
        if start_date:
            events = [e for e in events if e.end_time >= start_date]
        if end_date:
            events = [e for e in events if e.start_time <= end_date]
        return events

    def create_event(self, request: CreateMeetingRequest) -> CalendarEvent:
        event = CalendarEvent(
            id=f"evt-{uuid.uuid4().hex[:6]}",
            subject=request.subject,
            start_time=request.start_time,
            end_time=request.end_time,
            location=request.location,
            organizer="You (via Agent 365)",
            attendees=[Attendee(name=att.split("@")[0].capitalize(), email=att, status="accepted") for att in request.attendees],
            body_preview=request.description or "",
        )
        self.events.append(event)
        return event

    def check_conflicts(self, start_time: datetime, end_time: datetime) -> ConflictReport:
        conflicts: List[CalendarEvent] = []
        for e in self.events:
            # Overlap test: (StartA < EndB) and (EndA > StartB)
            if e.start_time < end_time and e.end_time > start_time:
                conflicts.append(e)

        if not conflicts:
            return ConflictReport(
                has_conflict=False,
                conflicting_events=[],
                suggested_alternative_slots=[],
                summary="No schedule conflicts detected for the specified time slot.",
            )

        # Generate smart alternative suggestions
        suggested_slots = [
            f"{(end_time + timedelta(hours=1)).strftime('%I:%M %p')} - {(end_time + timedelta(hours=2)).strftime('%I:%M %p')}",
            f"{(start_time + timedelta(days=1)).strftime('%Y-%m-%d %I:%M %p')}",
        ]

        conflict_names = ", ".join([f"'{c.subject}' ({c.start_time.strftime('%I:%M %p')})" for c in conflicts])
        return ConflictReport(
            has_conflict=True,
            conflicting_events=conflicts,
            suggested_alternative_slots=suggested_slots,
            summary=f"Detected conflict with {len(conflicts)} event(s): {conflict_names}.",
        )

    # --- Storage Operations ---
    def list_files(self, folder_path: str = "/") -> List[DriveItem]:
        return self.files

    def get_file_content(self, file_id: str) -> str:
        for f in self.files:
            if f.id == file_id or f.name.lower() == file_id.lower():
                if f.id in self.spreadsheets:
                    sheet = self.spreadsheets[f.id]
                    return f"Spreadsheet Data for {sheet.sheet_name}:\nHeaders: {sheet.headers}\nRows: {sheet.rows}\nSummary: {sheet.summary}"
                return f.summary_preview or f"Content of {f.name}"
        return f"File '{file_id}' not found."

    def upload_file(self, request: UploadFileRequest) -> DriveItem:
        now = datetime.now()
        item = DriveItem(
            id=f"file-{uuid.uuid4().hex[:6]}",
            name=request.name,
            web_url=f"https://enterprise365.sharepoint.com/sites/Documents/{request.name}",
            size_bytes=len(request.content.encode("utf-8")),
            created_at=now,
            modified_at=now,
            is_folder=False,
            summary_preview=request.content[:150],
        )
        self.files.append(item)
        return item

    def create_share_link(self, file_id: str) -> ShareLinkResponse:
        for f in self.files:
            if f.id == file_id or f.name == file_id:
                return ShareLinkResponse(
                    success=True,
                    item_id=f.id,
                    share_url=f"{f.web_url}?authkey=secure-token-{uuid.uuid4().hex[:8]}",
                    permission="view_edit",
                )
        return ShareLinkResponse(success=False, item_id=file_id, share_url="", permission="none")

    # --- Teams Operations ---
    def list_channels(self) -> List[TeamsChannel]:
        return self.channels

    def get_channel_messages(self, channel_id: str, limit: int = 10) -> List[TeamsMessage]:
        msgs = [m for m in self.messages if m.channel_id == channel_id or channel_id in (m.channel_id or "")]
        return sorted(msgs, key=lambda m: m.timestamp, reverse=True)[:limit]

    def post_channel_message(self, request: PostMessageRequest) -> TeamsMessage:
        new_msg = TeamsMessage(
            id=f"tmsg-{uuid.uuid4().hex[:6]}",
            channel_id=request.target_id,
            sender_name="Agent 365 Copilot",
            content=request.content,
            timestamp=datetime.now(),
            importance=request.importance,
        )
        self.messages.append(new_msg)
        return new_msg

    # --- Office / Excel Operations ---
    def get_spreadsheet_data(self, file_id: str, sheet_name: Optional[str] = None) -> ExcelTable:
        for fid, sheet in self.spreadsheets.items():
            if fid == file_id:
                return sheet
        for f in self.files:
            if (f.id == file_id or f.name.lower() == file_id.lower()) and f.id in self.spreadsheets:
                return self.spreadsheets[f.id]

        return ExcelTable(
            sheet_name=sheet_name or "Sheet1",
            headers=["Category", "Metric", "Value", "Status"],
            rows=[
                ["Operations", "SLA Compliance", "99.98%", "Optimal"],
                ["Security", "Zero Trust Identity", "100%", "Compliant"],
                ["Cost", "Cloud Efficiency Score", "94.2/100", "Top Decile"],
            ],
            summary={"status": "All key metrics healthy."},
        )

    def update_cell_value(self, file_id: str, sheet_name: str, coordinate: str, value: Any) -> bool:
        sheet = self.get_spreadsheet_data(file_id, sheet_name)
        # Parse coordinate e.g. B2 (col B=idx 1, row 2=idx 0)
        try:
            col_letter = coordinate[0].upper()
            row_idx = int(coordinate[1:]) - 2  # header is row 1
            col_idx = ord(col_letter) - ord('A')
            if 0 <= row_idx < len(sheet.rows) and 0 <= col_idx < len(sheet.headers):
                sheet.rows[row_idx][col_idx] = value
                # Recompute variance if col 1 (budget) or col 2 (actual)
                if col_idx in (1, 2):
                    try:
                        b = float(sheet.rows[row_idx][1])
                        a = float(sheet.rows[row_idx][2])
                        diff = round(a - b, 2)
                        sheet.rows[row_idx][3] = diff
                        sheet.rows[row_idx][4] = "Over Budget (Caution)" if diff > 0 else "Under Budget (Favorable)"
                    except Exception:
                        pass
                return True
        except Exception:
            pass
        return True

    def get_word_document(self, file_id: str = "file-302") -> DocumentContent:
        for fid, doc in self.word_documents.items():
            if fid == file_id:
                return doc
        for f in self.files:
            if (f.id == file_id or f.name.lower() == file_id.lower()) and f.id in self.word_documents:
                return self.word_documents[f.id]

        return DocumentContent(
            title="Executive_Strategic_Plan_2026.docx",
            author="Sarah Jenkins (CEO)",
            raw_content="Executive Strategic Plan 2026\n\n1. Overview\nDocument active in OneDrive.",
            sections=[DocumentParagraph(heading_level=1, text="Executive Strategic Plan 2026")],
        )

    def save_word_document(self, file_id: str, content: str) -> bool:
        doc = self.get_word_document(file_id)
        doc.raw_content = content
        return True

    def get_presentation(self, file_id: str = "file-304") -> PresentationContent:
        for fid, pres in self.presentations.items():
            if fid == file_id:
                return pres
        for f in self.files:
            if (f.id == file_id or f.name.lower() == file_id.lower()) and f.id in self.presentations:
                return self.presentations[f.id]

        # Return default presentation
        return self.presentations["file-304"]

    def get_messages_by_folder(self, folder_name: str = "inbox", limit: int = 20) -> List[EmailMessage]:
        f_lower = folder_name.lower().replace(" ", "").replace("_", "")
        for key in self.folders:
            if key in f_lower or f_lower in key:
                return sorted(self.folders[key], key=lambda m: m.received_at, reverse=True)[:limit]
        return self.get_inbox_messages(limit=limit)

    def delete_message(self, message_id: str) -> bool:
        for folder_name, msgs in self.folders.items():
            for m in msgs:
                if m.id == message_id:
                    msgs.remove(m)
                    self.folders["deleted"].append(m)
                    return True
        return False

    def archive_message(self, message_id: str) -> bool:
        for folder_name, msgs in self.folders.items():
            for m in msgs:
                if m.id == message_id:
                    msgs.remove(m)
                    self.folders["archive"].append(m)
                    return True
        return False

    def get_onenote_notebooks(self) -> List[OneNoteNotebook]:
        return self.onenote_notebooks

    def create_onenote_page(self, section_id: str, title: str, content: str) -> OneNotePage:
        new_page = OneNotePage(
            id=f"pg-{uuid.uuid4().hex[:6]}",
            title=title,
            created_at=datetime.now().strftime("%Y-%m-%d"),
            content=content,
        )
        for nb in self.onenote_notebooks:
            for sec in nb.sections:
                if sec.id == section_id or not section_id:
                    sec.pages.append(new_page)
                    return new_page
        if self.onenote_notebooks and self.onenote_notebooks[0].sections:
            self.onenote_notebooks[0].sections[0].pages.append(new_page)
        return new_page

    def create_onenote_section(self, notebook_id: str, section_name: str) -> OneNoteSection:
        new_section = OneNoteSection(
            id=f"sec-{uuid.uuid4().hex[:6]}",
            name=section_name,
            color="#0078d4",
            pages=[
                OneNotePage(
                    id=f"pg-{uuid.uuid4().hex[:6]}",
                    title="Untitled Page",
                    created_at=datetime.now().strftime("%Y-%m-%d"),
                    content=f"<h3>{section_name} Notes</h3><p>Start drafting your notes here.</p>",
                )
            ],
        )
        for nb in self.onenote_notebooks:
            if nb.id == notebook_id or not notebook_id:
                nb.sections.append(new_section)
                return new_section
        if self.onenote_notebooks:
            self.onenote_notebooks[0].sections.append(new_section)
        return new_section

    def update_onenote_page(self, page_id: str, title: str, content: str) -> bool:
        for nb in self.onenote_notebooks:
            for sec in nb.sections:
                for pg in sec.pages:
                    if pg.id == page_id:
                        pg.title = title
                        pg.content = content
                        return True
        return False

    def create_teams_channel(self, team_id: str, channel_name: str, description: Optional[str] = None) -> TeamsChannel:
        new_chan = TeamsChannel(
            id=f"chan-{uuid.uuid4().hex[:6]}",
            team_id=team_id or "team-101",
            team_name="Enterprise 365",
            channel_name=channel_name,
            description=description or f"Channel for {channel_name} discussions",
        )
        self.channels.append(new_chan)
        self.messages.append(
            TeamsMessage(
                id=f"tmsg-{uuid.uuid4().hex[:6]}",
                channel_id=new_chan.id,
                sender_name="System",
                content=f"Welcome to #{channel_name}! Start collaborating.",
                timestamp=datetime.now(),
            )
        )
        return new_chan


    def create_folder(self, folder_name: str, parent_folder: str = "/") -> DriveItem:
        new_item = DriveItem(
            id=f"folder-{uuid.uuid4().hex[:6]}",
            name=folder_name,
            web_url=f"https://enterprise365.sharepoint.com/documents/{folder_name}",
            size_bytes=0,
            created_at=datetime.now(),
            modified_at=datetime.now(),
            is_folder=True,
        )
        self.files.insert(0, new_item)
        return new_item

