"""
Office & Excel Agent specialist.
Handles financial workbook analysis, metric computation, and document reports.
"""

from typing import Any, Dict, List, Optional
from .base_agent import BaseAgent
from ..models import AgentRole, DocumentContent, DocumentParagraph, ExcelTable
from ..m365_client import M365ClientBase


class OfficeAgent(BaseAgent):
    """Specialist agent for Excel workbooks and Office documents."""

    def __init__(self, client: Optional[M365ClientBase] = None):
        super().__init__(
            role=AgentRole.OFFICE,
            description="Analyzes Excel workbooks, calculates financial metrics, and drafts Office reports.",
            client=client,
        )
        self.register_tool("analyze_spreadsheet", self.analyze_spreadsheet)
        self.register_tool("generate_report_document", self.generate_report_document)

    def analyze_spreadsheet(self, file_id: str = "file-301", sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Parses spreadsheet tabular data and generates key financial insights."""
        table: ExcelTable = self.client.get_spreadsheet_data(file_id=file_id, sheet_name=sheet_name)

        # Compute department-level variances
        insights = []
        for row in table.rows:
            dept = row[0]
            budget = row[1]
            actual = row[2]
            status = row[4] if len(row) > 4 else "N/A"
            insights.append(f"{dept}: Budget ${budget}k vs Actual ${actual}k ({status})")

        return {
            "sheet_name": table.sheet_name,
            "headers": table.headers,
            "row_count": len(table.rows),
            "rows": table.rows,
            "summary_metrics": table.summary,
            "computed_insights": insights,
        }

    def generate_report_document(self, title: str, sections_dict: Dict[str, str]) -> Dict[str, Any]:
        """Drafts a structured executive document."""
        sections = []
        for heading, body in sections_dict.items():
            sections.append(DocumentParagraph(heading_level=2, text=heading))
            sections.append(DocumentParagraph(heading_level=0, text=body))

        doc = DocumentContent(
            title=title,
            author="Agent 365 Executive Suite",
            sections=sections,
        )
        return doc.model_dump()

    def handle_task(self, task_instruction: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Processes spreadsheet calculation or document generation."""
        instruction_lower = task_instruction.lower()

        if any(w in instruction_lower for w in ["excel", "spreadsheet", "budget", "variance", "financial"]):
            step = self.execute_tool("analyze_spreadsheet", file_id="file-301")
            return {"action": "analyzed_spreadsheet", "step": step.model_dump()}

        # Default: generate report summary
        step = self.execute_tool(
            "generate_report_document",
            title="Executive Operational Briefing",
            sections_dict={
                "Overview": "Consolidated operational analysis produced by Agent 365.",
                "Action Items": "1. Resolve 2 PM calendar overlap. 2. Acknowledge Q4 revenue memo.",
            },
        )
        return {"action": "generated_document", "step": step.model_dump()}
