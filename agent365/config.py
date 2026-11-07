"""
Configuration module for Agent 365.
Manages Microsoft Graph API credentials, execution mode, and LLM providers.
"""

import os
from typing import Literal, Optional
from pydantic import BaseModel, Field
from dotenv import load_dotenv

# Load .env if present
load_dotenv()


class Settings(BaseModel):
    """Application settings and runtime configurations."""
    app_name: str = "Agent 365"
    version: str = "1.0.0"

    # Microsoft 365 Mode: 'mock' (offline sandbox) or 'live' (Microsoft Graph API)
    m365_mode: Literal["mock", "live"] = Field(
        default_factory=lambda: os.getenv("M365_MODE", "mock").lower() # type: ignore
    )
    fresh_out_of_box: bool = Field(
        default_factory=lambda: os.getenv("FRESH_OUT_OF_BOX", "true").lower() in ("true", "1", "yes")
    )

    # Entra ID / Microsoft Graph Credentials
    azure_tenant_id: Optional[str] = Field(default_factory=lambda: os.getenv("AZURE_TENANT_ID"))
    azure_client_id: Optional[str] = Field(default_factory=lambda: os.getenv("AZURE_CLIENT_ID"))
    azure_client_secret: Optional[str] = Field(default_factory=lambda: os.getenv("AZURE_CLIENT_SECRET"))
    m365_user_email: Optional[str] = Field(default_factory=lambda: os.getenv("M365_USER_EMAIL", "executive@enterprise365.onmicrosoft.com"))

    # LLM Settings
    llm_provider: Literal["heuristic", "openai", "gemini", "azure_openai"] = Field(
        default_factory=lambda: os.getenv("LLM_PROVIDER", "heuristic").lower() # type: ignore
    )
    openai_api_key: Optional[str] = Field(default_factory=lambda: os.getenv("OPENAI_API_KEY"))
    gemini_api_key: Optional[str] = Field(default_factory=lambda: os.getenv("GEMINI_API_KEY"))
    llm_model: str = Field(default_factory=lambda: os.getenv("LLM_MODEL", "gpt-4o"))

    # Server settings
    host: str = Field(default_factory=lambda: os.getenv("HOST", "127.0.0.1"))
    port: int = Field(default_factory=lambda: int(os.getenv("PORT", "8000")))
    debug: bool = Field(default_factory=lambda: os.getenv("DEBUG", "True").lower() in ("true", "1"))

    def update_m365_mode(self, mode: Literal["mock", "live"]):
        self.m365_mode = mode


# Singleton configuration instance
settings = Settings()
