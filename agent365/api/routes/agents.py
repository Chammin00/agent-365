"""
Agent Registry & Status API routes.
"""

from fastapi import APIRouter
from typing import Any, Dict, List
from ...agents import (
    OutlookMailAgent,
    CalendarAgent,
    StorageAgent,
    TeamsAgent,
    OfficeAgent,
    ChiefOfStaffOrchestrator,
)
from ...models import AgentRole

router = APIRouter(prefix="/agents", tags=["Agent Registry & Status"])


@router.get("/")
def list_agents() -> List[Dict[str, Any]]:
    """Returns the registered agents and their available capabilities/tools."""
    orchestrator = ChiefOfStaffOrchestrator()
    agents = [
        {
            "role": AgentRole.ORCHESTRATOR.value,
            "description": "Decomposes requests, plans cross-app workflows, and coordinates specialist agents.",
            "tools": ["plan_workflow", "run_multi_agent_pipeline"],
            "status": "ready",
            "icon": "bi-diagram-3",
        },
        {
            "role": AgentRole.MAIL.value,
            "description": orchestrator.mail_agent.description,
            "tools": list(orchestrator.mail_agent.tools.keys()),
            "status": "ready",
            "icon": "bi-envelope",
        },
        {
            "role": AgentRole.CALENDAR.value,
            "description": orchestrator.calendar_agent.description,
            "tools": list(orchestrator.calendar_agent.tools.keys()),
            "status": "ready",
            "icon": "bi-calendar-event",
        },
        {
            "role": AgentRole.STORAGE.value,
            "description": orchestrator.storage_agent.description,
            "tools": list(orchestrator.storage_agent.tools.keys()),
            "status": "ready",
            "icon": "bi-folder2-open",
        },
        {
            "role": AgentRole.TEAMS.value,
            "description": orchestrator.teams_agent.description,
            "tools": list(orchestrator.teams_agent.tools.keys()),
            "status": "ready",
            "icon": "bi-chat-dots",
        },
        {
            "role": AgentRole.OFFICE.value,
            "description": orchestrator.office_agent.description,
            "tools": list(orchestrator.office_agent.tools.keys()),
            "status": "ready",
            "icon": "bi-file-earmark-spreadsheet",
        },
    ]
    return agents
