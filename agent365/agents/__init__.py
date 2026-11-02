"""
Agent 365 agents package.
"""

from .base_agent import BaseAgent
from .mail_agent import OutlookMailAgent
from .calendar_agent import CalendarAgent
from .storage_agent import StorageAgent
from .teams_agent import TeamsAgent
from .office_agent import OfficeAgent
from .orchestrator import ChiefOfStaffOrchestrator

__all__ = [
    "BaseAgent",
    "OutlookMailAgent",
    "CalendarAgent",
    "StorageAgent",
    "TeamsAgent",
    "OfficeAgent",
    "ChiefOfStaffOrchestrator",
]
