"""
Microsoft Teams Agent specialist.
Handles channel announcements, message summaries, and team communication.
"""

from typing import Any, Dict, List, Optional
from .base_agent import BaseAgent
from ..models import AgentRole, PostMessageRequest
from ..m365_client import M365ClientBase


class TeamsAgent(BaseAgent):
    """Specialist agent for Microsoft Teams communication."""

    def __init__(self, client: Optional[M365ClientBase] = None):
        super().__init__(
            role=AgentRole.TEAMS,
            description="Monitors Microsoft Teams channels, fetches updates, and posts announcements.",
            client=client,
        )
        self.register_tool("list_channels", self.list_channels)
        self.register_tool("get_channel_messages", self.get_channel_messages)
        self.register_tool("post_message", self.post_message)

    def list_channels(self) -> Dict[str, Any]:
        """Lists accessible Teams channels."""
        channels = self.client.list_channels()
        return {
            "total_channels": len(channels),
            "channels": [c.model_dump() for c in channels],
        }

    def get_channel_messages(self, channel_id: str, limit: int = 5) -> Dict[str, Any]:
        """Retrieves recent discussions from a channel."""
        msgs = self.client.get_channel_messages(channel_id, limit=limit)
        return {
            "channel_id": channel_id,
            "message_count": len(msgs),
            "messages": [m.model_dump() for m in msgs],
        }

    def post_message(self, channel_id: str, content: str, importance: str = "normal") -> Dict[str, Any]:
        """Posts an update or announcement to a Teams channel."""
        req = PostMessageRequest(target_type="channel", target_id=channel_id, content=content, importance=importance)
        msg = self.client.post_channel_message(req)
        return msg.model_dump()

    def handle_task(self, task_instruction: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Processes Teams read or post tasks."""
        instruction_lower = task_instruction.lower()

        if any(w in instruction_lower for w in ["post", "notify", "announce", "send message", "teams"]):
            # Channel selection
            channel_id = "chan-401"
            if "eng" in instruction_lower:
                channel_id = "chan-402"
            elif "exec" in instruction_lower or "strat" in instruction_lower:
                channel_id = "chan-403"

            msg_text = (context.get("announcement_text") if context else None) or (
                "[Agent 365 Executive Update] Q4 financial figures and morning briefing items have been compiled and verified."
            )
            step = self.execute_tool("post_message", channel_id=channel_id, content=msg_text)
            return {"action": "posted_teams_message", "step": step.model_dump()}

        # Default: list channels or read recent
        step = self.execute_tool("list_channels")
        return {"action": "listed_channels", "step": step.model_dump()}
