"""
Chief of Staff Orchestrator for Microsoft 365.
Coordinates multi-agent workflows across Outlook, Calendar, OneDrive, Teams, and Excel.
"""

from datetime import datetime
from typing import Any, Dict, List, Optional
from .base_agent import BaseAgent
from .mail_agent import OutlookMailAgent
from .calendar_agent import CalendarAgent
from .storage_agent import StorageAgent
from .teams_agent import TeamsAgent
from .office_agent import OfficeAgent
from ..models import (
    AgentRole,
    AgentStep,
    PlanStep,
    OrchestratorResponse,
)
from ..m365_client import M365ClientBase, get_m365_client
from ..config import settings


class ChiefOfStaffOrchestrator:
    """Master orchestrator decomposing user queries into multi-agent workflows."""

    def __init__(self, client: Optional[M365ClientBase] = None):
        self.client = client or get_m365_client()
        self.mail_agent = OutlookMailAgent(self.client)
        self.calendar_agent = CalendarAgent(self.client)
        self.storage_agent = StorageAgent(self.client)
        self.teams_agent = TeamsAgent(self.client)
        self.office_agent = OfficeAgent(self.client)

        self.agents_map: Dict[AgentRole, BaseAgent] = {
            AgentRole.MAIL: self.mail_agent,
            AgentRole.CALENDAR: self.calendar_agent,
            AgentRole.STORAGE: self.storage_agent,
            AgentRole.TEAMS: self.teams_agent,
            AgentRole.OFFICE: self.office_agent,
        }

    def plan_workflow(self, query: str) -> List[PlanStep]:
        """Analyzes query intent and generates a structured plan."""
        q = query.lower()
        plan: List[PlanStep] = []
        step_num = 1

        # Check if broad briefing or daily summary requested
        is_briefing = any(w in q for w in ["briefing", "daily", "morning", "status", "overview", "summary", "everything"])

        # 1. Mail intent
        if is_briefing or any(w in q for w in ["email", "mail", "inbox", "unread", "sent", "sarah", "message"]):
            plan.append(
                PlanStep(
                    step_number=step_num,
                    target_agent=AgentRole.MAIL,
                    description="Scan Outlook inbox for urgent or unread messages and extract action items.",
                    expected_outcome="Prioritized digest of incoming emails.",
                )
            )
            step_num += 1

        # 2. Calendar intent
        if is_briefing or any(w in q for w in ["calendar", "schedule", "meeting", "conflict", "agenda", "appointment"]):
            plan.append(
                PlanStep(
                    step_number=step_num,
                    target_agent=AgentRole.CALENDAR,
                    description="Inspect upcoming calendar agenda and scan for potential double-bookings or scheduling conflicts.",
                    expected_outcome="List of scheduled events and conflict report.",
                )
            )
            step_num += 1

        # 3. Storage & OneDrive intent
        if is_briefing or any(w in q for w in ["file", "onedrive", "sharepoint", "document", "financial", "projections", "search"]):
            plan.append(
                PlanStep(
                    step_number=step_num,
                    target_agent=AgentRole.STORAGE,
                    description="Search OneDrive & SharePoint repositories for relevant financial models and executive documents.",
                    expected_outcome="Matched documents with links and content summaries.",
                )
            )
            step_num += 1

        # 4. Office & Excel intent
        if is_briefing or any(w in q for w in ["excel", "spreadsheet", "budget", "variance", "sheet", "numbers", "margin"]):
            plan.append(
                PlanStep(
                    step_number=step_num,
                    target_agent=AgentRole.OFFICE,
                    description="Parse the Q4 Financial Projections workbook and compute department budget variances.",
                    expected_outcome="Tabular summary with variance calculations.",
                )
            )
            step_num += 1

        # 5. Teams intent
        if any(w in q for w in ["teams", "notify", "post", "channel", "announcement", "chat"]):
            plan.append(
                PlanStep(
                    step_number=step_num,
                    target_agent=AgentRole.TEAMS,
                    description="Post notification or briefing summary to the appropriate Microsoft Teams channel.",
                    expected_outcome="Teams channel dispatch confirmation.",
                )
            )
            step_num += 1

        # Fallback if no specific keyword matched
        if not plan:
            plan = [
                PlanStep(
                    step_number=1,
                    target_agent=AgentRole.MAIL,
                    description="Check Outlook inbox status.",
                    expected_outcome="Inbox summary.",
                ),
                PlanStep(
                    step_number=2,
                    target_agent=AgentRole.CALENDAR,
                    description="Review upcoming calendar schedule.",
                    expected_outcome="Today's agenda.",
                ),
            ]

        return plan

    def run(self, query: str) -> OrchestratorResponse:
        """Executes the complete multi-agent workflow for a user query."""
        plan = self.plan_workflow(query)
        executed_steps: List[AgentStep] = []
        collected_context: Dict[str, Any] = {}

        for plan_step in plan:
            agent = self.agents_map.get(plan_step.target_agent)
            if not agent:
                continue

            result = agent.handle_task(query, context=collected_context)
            step_dict = result.get("step")
            if step_dict:
                agent_step = AgentStep(**step_dict)
                executed_steps.append(agent_step)
                # Pass data forward to downstream agents
                collected_context[plan_step.target_agent.value] = agent_step.output_data

        # Synthesize final response
        final_text = self._synthesize_response(query, executed_steps)

        return OrchestratorResponse(
            user_query=query,
            plan=plan,
            steps_executed=executed_steps,
            final_response=final_text,
            m365_mode=settings.m365_mode,
            success=True,
        )

    def _synthesize_response(self, query: str, steps: List[AgentStep]) -> str:
        """Assembles structured output into a comprehensive executive briefing."""
        sections: List[str] = [f"### [Executive Report] Agent 365 Executive Report\n**Request:** *\"{query}\"*\n"]

        for step in steps:
            role_name = step.agent_role.value
            res = step.output_data.get("result", {}) if step.output_data else {}

            if step.agent_role == AgentRole.MAIL:
                unread = res.get("unread_count", 0)
                high_pri = res.get("high_priority_emails", [])
                sections.append(f"#### [Outlook Mail] ({unread} unread)")
                if high_pri:
                    for em in high_pri:
                        sections.append(f"- **[PRIORITY] {em.get('subject')}** from `{em.get('sender', {}).get('name')}`\n  > {em.get('body_preview')}")
                else:
                    sections.append("- No urgent email flags currently pending.")

            elif step.agent_role == AgentRole.CALENDAR:
                if "conflicting_events" in res:
                    has_conflict = res.get("has_conflict", False)
                    if has_conflict:
                        summary = res.get("summary", "")
                        sections.append(f"#### [Calendar Conflict Alert]\n- **[CONFLICT]** {summary}")
                        alts = res.get("suggested_alternative_slots", [])
                        if alts:
                            sections.append(f"- **Suggested Alternative Times:** {', '.join(alts)}")
                    else:
                        sections.append("#### [Calendar Status]\n- All clear: No conflicts detected.")
                elif "events" in res:
                    events = res.get("events", [])
                    sections.append(f"#### [Calendar Agenda] ({len(events)} events scheduled)")
                    for ev in events[:3]:
                        st = ev.get("start_time", "").split("T")[-1][:5]
                        sections.append(f"- **{st}** — {ev.get('subject')} (Organizer: {ev.get('organizer')})")

            elif step.agent_role == AgentRole.STORAGE:
                matched = res.get("matched_files", []) or res.get("files", [])
                sections.append(f"#### [OneDrive & SharePoint] ({len(matched)} files located)")
                for f in matched[:3]:
                    sections.append(f"- **[{f.get('name')}]({f.get('web_url')})** ({round(f.get('size_bytes', 0) / 1024, 1)} KB)\n  > {f.get('summary_preview')}")

            elif step.agent_role == AgentRole.OFFICE:
                insights = res.get("computed_insights", [])
                summary_metrics = res.get("summary_metrics", {})
                sections.append(f"#### [Excel Analysis] ({res.get('sheet_name', 'Summary')})")
                if summary_metrics:
                    sections.append(f"- **Net Variance:** `${summary_metrics.get('net_variance_k', 0)}k` | **Status:** {summary_metrics.get('overall_status', 'OK')}")
                for ins in insights[:4]:
                    sections.append(f"  - {ins}")

            elif step.agent_role == AgentRole.TEAMS:
                if "content" in res:
                    sections.append(f"#### [Microsoft Teams Notification]\n- **Status:** Message dispatched successfully.\n  > \"{res.get('content')}\"")
                else:
                    sections.append("#### [Microsoft Teams]\n- Channel activity refreshed.")

        sections.append("\n--- \n*Generated autonomously by Agent 365 Enterprise Suite*")
        return "\n\n".join(sections)
