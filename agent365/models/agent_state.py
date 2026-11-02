"""
Data models for Multi-Agent coordination, planning, and execution traces.
"""

from datetime import datetime
from enum import Enum
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


class AgentRole(str, Enum):
    ORCHESTRATOR = "Chief of Staff (Orchestrator)"
    MAIL = "Outlook Mail Agent"
    CALENDAR = "Calendar & Scheduling Agent"
    STORAGE = "OneDrive & Files Agent"
    TEAMS = "Microsoft Teams Agent"
    OFFICE = "Office & Excel Agent"


class AgentStep(BaseModel):
    step_id: str
    agent_role: AgentRole
    action: str
    input_data: Optional[Dict[str, Any]] = None
    output_data: Optional[Dict[str, Any]] = None
    status: str = "completed"  # pending, in_progress, completed, failed
    timestamp: datetime = Field(default_factory=datetime.now)
    duration_ms: float = 0.0


class PlanStep(BaseModel):
    step_number: int
    target_agent: AgentRole
    description: str
    expected_outcome: str


class OrchestratorResponse(BaseModel):
    user_query: str
    plan: List[PlanStep] = Field(default_factory=list)
    steps_executed: List[AgentStep] = Field(default_factory=list)
    final_response: str
    m365_mode: str
    success: bool = True
