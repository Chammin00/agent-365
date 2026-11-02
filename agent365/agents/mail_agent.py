"""
Outlook Mail Agent specialist.
Handles email triage, unread digest, priority detection, reply drafts, and sending.
"""

from typing import Any, Dict, List, Optional
from .base_agent import BaseAgent
from ..models import AgentRole, DraftEmailRequest, EmailMessage
from ..m365_client import M365ClientBase


class OutlookMailAgent(BaseAgent):
    """Specialist agent for Outlook Mail."""

    def __init__(self, client: Optional[M365ClientBase] = None):
        super().__init__(
            role=AgentRole.MAIL,
            description="Manages Outlook inbox, email summaries, priority triage, and draft dispatch.",
            client=client,
        )
        self.register_tool("triage_inbox", self.triage_inbox)
        self.register_tool("get_email_details", self.get_email_details)
        self.register_tool("send_email", self.send_email)
        self.register_tool("flag_urgent", self.flag_urgent)

    def triage_inbox(self, limit: int = 10, unread_only: bool = False) -> Dict[str, Any]:
        """Fetches and categorizes inbox messages by priority."""
        messages = self.client.get_inbox_messages(limit=limit, unread_only=unread_only)
        high_priority = [m.model_dump() for m in messages if m.importance == "high" or m.is_flagged]
        standard = [m.model_dump() for m in messages if m.importance != "high" and not m.is_flagged]

        return {
            "total_messages": len(messages),
            "unread_count": sum(1 for m in messages if not m.is_read),
            "high_priority_count": len(high_priority),
            "high_priority_emails": high_priority,
            "standard_emails": standard,
        }

    def get_email_details(self, message_id: str) -> Optional[Dict[str, Any]]:
        """Retrieves complete content of a specific email."""
        msg = self.client.get_message_by_id(message_id)
        return msg.model_dump() if msg else None

    def send_email(self, to: List[str], subject: str, body: str, importance: str = "normal") -> Dict[str, Any]:
        """Drafts and sends an email via Microsoft 365."""
        req = DraftEmailRequest(to=to, subject=subject, body=body, importance=importance)
        resp = self.client.send_email(req)
        return resp.model_dump()

    def flag_urgent(self, message_id: str) -> bool:
        """Flags an email as high-priority."""
        return self.client.flag_message(message_id)

    def handle_task(self, task_instruction: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Executes mail operations based on natural language task."""
        instruction_lower = task_instruction.lower()

        if any(w in instruction_lower for w in ["send", "draft", "reply"]):
            # Extract recipient and subject if provided in context
            to = context.get("to", ["sarah.jenkins@enterprise365.com"]) if context else ["sarah.jenkins@enterprise365.com"]
            subj = context.get("subject", "Follow-up: M365 Agent Briefing") if context else "Follow-up: M365 Agent Briefing"
            body = context.get("body", "Hi Sarah,\n\nI have reviewed the requested items and attached the summary.\n\nBest regards,\nExecutive Team") if context else "Automated update via Agent 365."
            step = self.execute_tool("send_email", to=to, subject=subj, body=body)
            return {"action": "sent_email", "step": step.model_dump()}

        # Default action: Triage inbox
        unread_only = "unread" in instruction_lower
        step = self.execute_tool("triage_inbox", limit=5, unread_only=unread_only)
        return {"action": "triaged_inbox", "step": step.model_dump()}
