"""
Unit tests for Mock Microsoft 365 Client.
"""

from datetime import datetime, timedelta
from agent365.m365_client import MockM365Client
from agent365.models import DraftEmailRequest, CreateMeetingRequest, PostMessageRequest, UploadFileRequest


def test_mock_mail_operations():
    client = MockM365Client()
    inbox = client.get_inbox_messages()
    assert len(inbox) >= 4

    unread = client.get_inbox_messages(unread_only=True)
    assert len(unread) > 0

    # Test sending email
    req = DraftEmailRequest(
        to=["sarah.jenkins@enterprise365.com"],
        subject="Test Briefing",
        body="This is an automated test email.",
    )
    resp = client.send_email(req)
    assert resp.success is True

    # Verify message exists in inbox
    new_inbox = client.get_inbox_messages()
    assert any(m.subject == "Test Briefing" for m in new_inbox)


def test_mock_calendar_conflicts():
    client = MockM365Client()
    today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    # 14:00 - 15:00 has the Executive Q4 Review and Northwind Traders sync
    conflict_start = today_start + timedelta(hours=14, minutes=15)
    conflict_end = today_start + timedelta(hours=14, minutes=45)

    report = client.check_conflicts(conflict_start, conflict_end)
    assert report.has_conflict is True
    assert len(report.conflicting_events) >= 1
    assert len(report.suggested_alternative_slots) > 0

    # 22:00 has no conflict
    free_start = today_start + timedelta(hours=22)
    free_end = today_start + timedelta(hours=23)
    free_report = client.check_conflicts(free_start, free_end)
    assert free_report.has_conflict is False


def test_mock_onedrive_operations():
    client = MockM365Client()
    files = client.list_files()
    assert len(files) >= 3
    assert any("Q4_Financial_Projections" in f.name for f in files)

    # Test file upload
    upload_req = UploadFileRequest(name="TestDoc.docx", content="Agent 365 test content")
    uploaded = client.upload_file(upload_req)
    assert uploaded.name == "TestDoc.docx"

    # Test share link
    link = client.create_share_link(uploaded.id)
    assert link.success is True
    assert "enterprise365" in link.share_url


def test_mock_teams_and_excel():
    client = MockM365Client()
    channels = client.list_channels()
    assert len(channels) >= 3

    # Post message
    post_req = PostMessageRequest(target_type="channel", target_id="chan-401", content="Agent test message")
    msg = client.post_channel_message(post_req)
    assert msg.content == "Agent test message"

    # Excel spreadsheet
    sheet = client.get_spreadsheet_data("file-301")
    assert sheet.sheet_name == "Q4_Summary"
    assert len(sheet.rows) > 0
