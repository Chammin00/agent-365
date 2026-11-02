"""
Calendar & Scheduling Agent specialist.
Handles agenda lookups, scheduling conflicts, and smart time slot proposals.
"""

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
from .base_agent import BaseAgent
from ..models import AgentRole, CreateMeetingRequest
from ..m365_client import M365ClientBase


class CalendarAgent(BaseAgent):
    """Specialist agent for Outlook Calendar and meeting scheduling."""

    def __init__(self, client: Optional[M365ClientBase] = None):
        super().__init__(
            role=AgentRole.CALENDAR,
            description="Manages Outlook Calendar, schedules meetings, and resolves calendar conflicts.",
            client=client,
        )
        self.register_tool("get_agenda", self.get_agenda)
        self.register_tool("check_conflicts", self.check_conflicts)
        self.register_tool("schedule_meeting", self.schedule_meeting)

    def get_agenda(self, days_ahead: int = 1) -> Dict[str, Any]:
        """Fetches upcoming calendar agenda."""
        now = datetime.now()
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=days_ahead, hours=23, minutes=59)

        events = self.client.get_events(start_date=start, end_date=end)
        return {
            "total_events": len(events),
            "events": [e.model_dump() for e in events],
            "start_window": start.isoformat(),
            "end_window": end.isoformat(),
        }

    def check_conflicts(self, start_iso: str, end_iso: str) -> Dict[str, Any]:
        """Evaluates scheduling conflicts for a target timeframe."""
        start = datetime.fromisoformat(start_iso)
        end = datetime.fromisoformat(end_iso)
        report = self.client.check_conflicts(start, end)
        return report.model_dump()

    def schedule_meeting(
        self,
        subject: str,
        start_iso: str,
        end_iso: str,
        attendees: List[str],
        location: str = "Microsoft Teams Meeting",
        description: str = "",
    ) -> Dict[str, Any]:
        """Creates a new calendar event with Teams meeting link."""
        req = CreateMeetingRequest(
            subject=subject,
            start_time=datetime.fromisoformat(start_iso),
            end_time=datetime.fromisoformat(end_iso),
            attendees=attendees,
            location=location,
            description=description,
        )
        event = self.client.create_event(req)
        return event.model_dump()

    def handle_task(self, task_instruction: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Executes calendar analysis or booking."""
        instruction_lower = task_instruction.lower()

        if "conflict" in instruction_lower:
            now = datetime.now()
            today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
            # Check afternoon slot 14:00 - 15:00 by default or from context
            start_iso = (today_start + timedelta(hours=14)).isoformat()
            end_iso = (today_start + timedelta(hours=15)).isoformat()
            step = self.execute_tool("check_conflicts", start_iso=start_iso, end_iso=end_iso)
            return {"action": "checked_conflicts", "step": step.model_dump()}

        if any(w in instruction_lower for w in ["schedule", "book", "create meeting"]):
            now = datetime.now()
            start_iso = (now + timedelta(hours=2)).isoformat()
            end_iso = (now + timedelta(hours=3)).isoformat()
            attendees = ["marcus.chen@enterprise365.com"]
            step = self.execute_tool(
                "schedule_meeting",
                subject="Agent 365 Review Sync",
                start_iso=start_iso,
                end_iso=end_iso,
                attendees=attendees,
            )
            return {"action": "scheduled_meeting", "step": step.model_dump()}

        # Default: Fetch daily agenda
        step = self.execute_tool("get_agenda", days_ahead=1)
        return {"action": "retrieved_agenda", "step": step.model_dump()}
