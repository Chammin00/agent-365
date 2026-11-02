"""
Unit tests for Agent 365 FastAPI endpoints.
"""

from fastapi.testclient import TestClient
from agent365.api import app

client = TestClient(app)


def test_health_endpoint():
    resp = client.get("/health")
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "healthy"


def test_get_agents():
    resp = client.get("/api/agents/")
    assert resp.status_code == 200
    agents = resp.json()
    assert len(agents) >= 6


def test_get_inbox():
    resp = client.get("/api/m365/inbox?limit=5")
    assert resp.status_code == 200
    emails = resp.json()
    assert len(emails) > 0


def test_get_calendar():
    resp = client.get("/api/m365/calendar")
    assert resp.status_code == 200
    events = resp.json()
    assert len(events) > 0


def test_chat_orchestrator():
    payload = {"message": "Check my inbox and review calendar conflicts for 2 PM"}
    resp = client.post("/api/chat/", json=payload)
    assert resp.status_code == 200
    data = resp.json()
    assert "final_response" in data
    assert len(data["steps_executed"]) >= 2


def test_onenote_endpoints():
    # Test get notebooks
    resp = client.get("/api/m365/onenote/notebooks")
    assert resp.status_code == 200
    notebooks = resp.json()
    assert len(notebooks) > 0
    assert len(notebooks[0]["sections"]) > 0

    # Test create section
    sec_resp = client.post("/api/m365/onenote/section", json={"notebook_id": "nb-01", "section_name": "Test Section"})
    assert sec_resp.status_code == 200
    sec_data = sec_resp.json()
    assert sec_data["name"] == "Test Section"

    # Test create page
    page_resp = client.post("/api/m365/onenote/page", json={
        "section_id": sec_data["id"],
        "title": "Automated Test Page",
        "content": "<p>Content for testing</p>"
    })
    assert page_resp.status_code == 200
    page_data = page_resp.json()
    assert page_data["title"] == "Automated Test Page"

    # Test update page
    update_resp = client.post("/api/m365/onenote/page/update", json={
        "page_id": page_data["id"],
        "title": "Updated Test Page",
        "content": "<p>Updated Content</p>"
    })
    assert update_resp.status_code == 200
    assert update_resp.json()["success"] is True


def test_teams_channel_creation():
    resp = client.post("/api/m365/teams/channel", json={"team_id": "team-101", "channel_name": "test-channel"})
    assert resp.status_code == 200
    assert resp.json()["channel_name"] == "test-channel"


def test_folder_creation():
    resp = client.post("/api/m365/files/folder", json={"folder_name": "Test Folder"})
    assert resp.status_code == 200
    assert resp.json()["name"] == "Test Folder"
    assert resp.json()["is_folder"] is True


def test_word_and_excel_endpoints():
    # Word save
    w_resp = client.post("/api/m365/word/save", json={"file_id": "file-302", "content": "<p>Updated text</p>"})
    assert w_resp.status_code == 200
    assert w_resp.json()["success"] is True

    # Excel cell update
    e_resp = client.post("/api/m365/office/cell", json={"file_id": "file-301", "coordinate": "B2", "value": 2500})
    assert e_resp.status_code == 200
    assert e_resp.json()["value"] == 2500

