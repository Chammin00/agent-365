"""
Chat API route for Agent 365 multi-agent interactions.
"""

from fastapi import APIRouter
from pydantic import BaseModel
from typing import Optional
from ...agents import ChiefOfStaffOrchestrator
from ...models import OrchestratorResponse

router = APIRouter(prefix="/chat", tags=["Agent Chat & Orchestration"])


class ChatRequest(BaseModel):
    message: str
    mode: Optional[str] = None


@router.post("/", response_model=OrchestratorResponse)
def execute_chat(request: ChatRequest):
    """Processes a user prompt through the Chief of Staff multi-agent orchestrator."""
    orchestrator = ChiefOfStaffOrchestrator()
    response = orchestrator.run(request.message)
    return response
