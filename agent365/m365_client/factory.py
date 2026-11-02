"""
Client factory for Microsoft 365 services.
Provides seamless switching between mock sandbox and live Graph API client.
"""

from typing import Optional
from .base import M365ClientBase
from .mock_client import MockM365Client
from .graph_client import LiveM365GraphClient
from ..config import settings

_current_client: Optional[M365ClientBase] = None
_current_mode: Optional[str] = None


def get_m365_client(force_mode: Optional[str] = None) -> M365ClientBase:
    """Returns the active M365 client (cached singleton unless mode changes)."""
    global _current_client, _current_mode

    target_mode = (force_mode or settings.m365_mode).lower()

    if _current_client is not None and _current_mode == target_mode:
        return _current_client

    if target_mode == "live":
        if not (settings.azure_tenant_id and settings.azure_client_id and settings.azure_client_secret):
            # Fallback gracefully with a warning
            print("[Agent 365 Warning] Live mode requested but Azure credentials missing. Falling back to Mock Sandbox.")
            _current_client = MockM365Client()
            _current_mode = "mock"
            return _current_client

        _current_client = LiveM365GraphClient(
            tenant_id=settings.azure_tenant_id,
            client_id=settings.azure_client_id,
            client_secret=settings.azure_client_secret,
            user_email=settings.m365_user_email,
        )
        _current_mode = "live"
    else:
        _current_client = MockM365Client()
        _current_mode = "mock"

    return _current_client


def reset_client():
    """Resets client singleton to allow reload upon configuration changes."""
    global _current_client, _current_mode
    _current_client = None
    _current_mode = None
