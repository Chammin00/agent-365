"""
Unit tests for specialized M365 agents and Chief of Staff orchestrator.
"""

from agent365.agents import (
    OutlookMailAgent,
    CalendarAgent,
    StorageAgent,
    TeamsAgent,
    OfficeAgent,
    ChiefOfStaffOrchestrator,
)
from agent365.models import AgentRole


def test_mail_agent():
    agent = OutlookMailAgent()
    res = agent.triage_inbox()
    assert "total_messages" in res
    assert res["total_messages"] > 0


def test_calendar_agent():
    agent = CalendarAgent()
    agenda = agent.get_agenda()
    assert "events" in agenda
    assert len(agenda["events"]) > 0


def test_storage_agent():
    agent = StorageAgent()
    search_res = agent.search_file_by_name("Financial")
    assert search_res["matched_count"] >= 1


def test_teams_agent():
    agent = TeamsAgent()
    channels = agent.list_channels()
    assert channels["total_channels"] >= 3


def test_office_agent():
    agent = OfficeAgent()
    analysis = agent.analyze_spreadsheet("file-301")
    assert len(analysis["computed_insights"]) > 0


def test_chief_of_staff_orchestrator():
    orchestrator = ChiefOfStaffOrchestrator()
    query = "Prepare my morning briefing: check urgent emails, review calendar conflicts, find Q4 financials on OneDrive, and notify Teams."

    plan = orchestrator.plan_workflow(query)
    assert len(plan) >= 4

    response = orchestrator.run(query)
    assert response.success is True
    assert len(response.steps_executed) >= 4
    assert "Agent 365 Executive Report" in response.final_response
    assert all(step.status == "completed" for step in response.steps_executed)
