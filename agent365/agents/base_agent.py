"""
Base class for specialized Agent 365 agents.
Provides tool management, execution tracing, and LLM / heuristic reasoning hooks.
"""

from abc import ABC, abstractmethod
import time
import uuid
from typing import Any, Callable, Dict, List, Optional
from ..models import AgentRole, AgentStep
from ..m365_client import M365ClientBase, get_m365_client


class BaseAgent(ABC):
    """Abstract base class for domain-specific M365 agents."""

    def __init__(self, role: AgentRole, description: str, client: Optional[M365ClientBase] = None):
        self.role = role
        self.description = description
        self.client = client or get_m365_client()
        self.tools: Dict[str, Callable[..., Any]] = {}

    def register_tool(self, name: str, func: Callable[..., Any]):
        """Registers an executable tool for this agent."""
        self.tools[name] = func

    def execute_tool(self, tool_name: str, **kwargs) -> AgentStep:
        """Executes a registered tool with execution tracing and timing."""
        if tool_name not in self.tools:
            return AgentStep(
                step_id=f"step-{uuid.uuid4().hex[:6]}",
                agent_role=self.role,
                action=tool_name,
                input_data=kwargs,
                output_data={"error": f"Tool '{tool_name}' not registered for {self.role.value}"},
                status="failed",
            )

        start_time = time.time()
        try:
            result = self.tools[tool_name](**kwargs)
            duration_ms = round((time.time() - start_time) * 1000, 2)
            return AgentStep(
                step_id=f"step-{uuid.uuid4().hex[:6]}",
                agent_role=self.role,
                action=tool_name,
                input_data=kwargs,
                output_data={"result": result},
                status="completed",
                duration_ms=duration_ms,
            )
        except Exception as exc:
            duration_ms = round((time.time() - start_time) * 1000, 2)
            return AgentStep(
                step_id=f"step-{uuid.uuid4().hex[:6]}",
                agent_role=self.role,
                action=tool_name,
                input_data=kwargs,
                output_data={"error": str(exc)},
                status="failed",
                duration_ms=duration_ms,
            )

    @abstractmethod
    def handle_task(self, task_instruction: str, context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Processes a high-level task instruction assigned to this specialist agent."""
        pass
