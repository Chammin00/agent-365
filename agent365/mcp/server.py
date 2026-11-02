"""
Model Context Protocol (MCP) Server for Agent 365.
Exposes Microsoft 365 tools and multi-agent workflows over standard MCP JSON-RPC protocol.
"""

import sys
import json
from typing import Any, Dict, List
from ..m365_client import get_m365_client
from ..agents import ChiefOfStaffOrchestrator
from ..models import DraftEmailRequest, CreateMeetingRequest, PostMessageRequest
from datetime import datetime


class Agent365MCPServer:
    """Standard Model Context Protocol server exposing Microsoft 365 tools."""

    def __init__(self):
        self.client = get_m365_client()
        self.orchestrator = ChiefOfStaffOrchestrator(self.client)

    def list_tools(self) -> List[Dict[str, Any]]:
        """Returns standard list of available MCP tools."""
        return [
            {
                "name": "m365_run_orchestrator",
                "description": "Execute a full multi-agent workflow across Outlook, Calendar, OneDrive, Teams, and Excel using Chief of Staff Orchestrator.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "query": {"type": "string", "description": "The high-level user request or instructions"}
                    },
                    "required": ["query"],
                },
            },
            {
                "name": "m365_get_inbox",
                "description": "Fetch inbox messages from Outlook Mail with unread/urgent indicators.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "limit": {"type": "integer", "description": "Number of emails to fetch", "default": 5},
                        "unread_only": {"type": "boolean", "description": "Filter for unread messages", "default": False},
                    },
                },
            },
            {
                "name": "m365_send_email",
                "description": "Send an email via Microsoft 365 Outlook.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "to": {"type": "array", "items": {"type": "string"}, "description": "Recipient email addresses"},
                        "subject": {"type": "string", "description": "Email subject line"},
                        "body": {"type": "string", "description": "Body content of the email"},
                        "importance": {"type": "string", "enum": ["low", "normal", "high"], "default": "normal"},
                    },
                    "required": ["to", "subject", "body"],
                },
            },
            {
                "name": "m365_check_calendar_conflicts",
                "description": "Inspect calendar for meeting conflicts in a given time slot.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "start_time": {"type": "string", "description": "ISO format start timestamp"},
                        "end_time": {"type": "string", "description": "ISO format end timestamp"},
                    },
                    "required": ["start_time", "end_time"],
                },
            },
            {
                "name": "m365_list_files",
                "description": "List files and documents in OneDrive / SharePoint storage.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "folder_path": {"type": "string", "description": "Folder path", "default": "/"}
                    },
                },
            },
            {
                "name": "m365_post_teams_message",
                "description": "Post a message or announcement to a Microsoft Teams channel.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "channel_id": {"type": "string", "description": "ID of target Teams channel"},
                        "content": {"type": "string", "description": "Message content"},
                    },
                    "required": ["channel_id", "content"],
                },
            },
            {
                "name": "m365_analyze_excel",
                "description": "Retrieve and analyze financial spreadsheet data from OneDrive.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "file_id": {"type": "string", "description": "ID or name of Excel file", "default": "file-301"}
                    },
                },
            },
        ]

    def call_tool(self, name: str, arguments: Dict[str, Any]) -> Any:
        """Dispatches an MCP tool call."""
        if name == "m365_run_orchestrator":
            return self.orchestrator.run(arguments.get("query", "")).model_dump()

        elif name == "m365_get_inbox":
            limit = arguments.get("limit", 5)
            unread = arguments.get("unread_only", False)
            msgs = self.client.get_inbox_messages(limit=limit, unread_only=unread)
            return [m.model_dump() for m in msgs]

        elif name == "m365_send_email":
            req = DraftEmailRequest(
                to=arguments.get("to", []),
                subject=arguments.get("subject", ""),
                body=arguments.get("body", ""),
                importance=arguments.get("importance", "normal"),
            )
            return self.client.send_email(req).model_dump()

        elif name == "m365_check_calendar_conflicts":
            start = datetime.fromisoformat(arguments.get("start_time"))
            end = datetime.fromisoformat(arguments.get("end_time"))
            return self.client.check_conflicts(start, end).model_dump()

        elif name == "m365_list_files":
            files = self.client.list_files(arguments.get("folder_path", "/"))
            return [f.model_dump() for f in files]

        elif name == "m365_post_teams_message":
            req = PostMessageRequest(
                target_type="channel",
                target_id=arguments.get("channel_id", "chan-401"),
                content=arguments.get("content", ""),
            )
            return self.client.post_channel_message(req).model_dump()

        elif name == "m365_analyze_excel":
            return self.client.get_spreadsheet_data(arguments.get("file_id", "file-301")).model_dump()

        else:
            raise ValueError(f"Unknown MCP tool: {name}")

    def run_stdio(self):
        """Standard IO JSON-RPC loop for MCP clients."""
        for line in sys.stdin:
            if not line.strip():
                continue
            try:
                req = json.loads(line)
                req_id = req.get("id")
                method = req.get("method")
                params = req.get("params", {})

                if method == "tools/list":
                    resp = {"jsonrpc": "2.0", "id": req_id, "result": {"tools": self.list_tools()}}
                elif method == "tools/call":
                    tool_name = params.get("name")
                    args = params.get("arguments", {})
                    tool_result = self.call_tool(tool_name, args)
                    resp = {
                        "jsonrpc": "2.0",
                        "id": req_id,
                        "result": {"content": [{"type": "text", "text": json.dumps(tool_result, indent=2, default=str)}]},
                    }
                else:
                    resp = {"jsonrpc": "2.0", "id": req_id, "result": {}}

                sys.stdout.write(json.dumps(resp) + "\n")
                sys.stdout.flush()
            except Exception as e:
                err_resp = {"jsonrpc": "2.0", "id": req.get("id"), "error": {"code": -32000, "message": str(e)}}
                sys.stdout.write(json.dumps(err_resp) + "\n")
                sys.stdout.flush()


if __name__ == "__main__":
    server = Agent365MCPServer()
    server.run_stdio()
