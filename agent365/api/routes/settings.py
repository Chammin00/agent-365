"""
Settings & Governance API routes.
"""

from fastapi import APIRouter
from pydantic import BaseModel
from typing import Literal, Optional
from ...config import settings
from ...m365_client import reset_client

router = APIRouter(prefix="/settings", tags=["Governance & Tenant Configuration"])


class ModeUpdateRequest(BaseModel):
    mode: Literal["mock", "live"]
    azure_tenant_id: Optional[str] = None
    azure_client_id: Optional[str] = None
    azure_client_secret: Optional[str] = None
    m365_user_email: Optional[str] = None


@router.get("/")
def get_current_settings():
    """Returns active settings and connection mode."""
    return {
        "app_name": settings.app_name,
        "version": settings.version,
        "m365_mode": settings.m365_mode,
        "has_azure_tenant": bool(settings.azure_tenant_id),
        "has_azure_client": bool(settings.azure_client_id),
        "user_email": settings.m365_user_email,
        "llm_provider": settings.llm_provider,
    }


@router.post("/mode")
def update_mode(request: ModeUpdateRequest):
    """Updates runtime M365 execution mode and credentials."""
    settings.m365_mode = request.mode
    if request.azure_tenant_id:
        settings.azure_tenant_id = request.azure_tenant_id
    if request.azure_client_id:
        settings.azure_client_id = request.azure_client_id
    if request.azure_client_secret:
        settings.azure_client_secret = request.azure_client_secret
    if request.m365_user_email:
        settings.m365_user_email = request.m365_user_email

    reset_client()
    return {
        "success": True,
        "new_mode": settings.m365_mode,
        "detail": f"Switched to {settings.m365_mode.upper()} mode.",
    }
