"""
OneDrive & SharePoint Storage Agent specialist.
Handles file search, document content inspection, file uploads, and share links.
"""

from typing import Any, Dict, List, Optional
from .base_agent import BaseAgent
from ..models import AgentRole, UploadFileRequest
from ..m365_client import M365ClientBase


class StorageAgent(BaseAgent):
    """Specialist agent for OneDrive and SharePoint storage."""

    def __init__(self, client: Optional[M365ClientBase] = None):
        super().__init__(
            role=AgentRole.STORAGE,
            description="Searches, inspects, uploads, and generates share links for OneDrive & SharePoint files.",
            client=client,
        )
        self.register_tool("list_files", self.list_files)
        self.register_tool("search_file_by_name", self.search_file_by_name)
        self.register_tool("get_file_content", self.get_file_content)
        self.register_tool("create_share_link", self.create_share_link)

    def list_files(self, folder_path: str = "/") -> Dict[str, Any]:
        """Lists files and folders in OneDrive."""
        files = self.client.list_files(folder_path=folder_path)
        return {
            "total_items": len(files),
            "files": [f.model_dump() for f in files],
        }

    def search_file_by_name(self, keyword: str) -> Dict[str, Any]:
        """Finds files whose name or summary matches the keyword."""
        files = self.client.list_files()
        matched = [
            f.model_dump()
            for f in files
            if keyword.lower() in f.name.lower() or (f.summary_preview and keyword.lower() in f.summary_preview.lower())
        ]
        return {
            "query": keyword,
            "matched_count": len(matched),
            "matched_files": matched,
        }

    def get_file_content(self, file_id_or_name: str) -> Dict[str, Any]:
        """Retrieves textual summary or contents of a document."""
        content = self.client.get_file_content(file_id_or_name)
        return {"file": file_id_or_name, "content": content}

    def create_share_link(self, file_id: str) -> Dict[str, Any]:
        """Generates an organization share link for a file."""
        link_resp = self.client.create_share_link(file_id)
        return link_resp.model_dump()

    def handle_task(self, task_instruction: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Processes OneDrive search or inspect task."""
        instruction_lower = task_instruction.lower()

        # Check for specific search targets
        for kw in ["financial", "budget", "q4", "security", "audit", "plan", "roadmap"]:
            if kw in instruction_lower:
                step = self.execute_tool("search_file_by_name", keyword=kw)
                return {"action": "searched_files", "step": step.model_dump()}

        if any(w in instruction_lower for w in ["share", "link"]):
            step = self.execute_tool("create_share_link", file_id="file-301")
            return {"action": "created_share_link", "step": step.model_dump()}

        # Default: list files
        step = self.execute_tool("list_files")
        return {"action": "listed_files", "step": step.model_dump()}
