"""
FastAPI application factory for Agent 365.
Serves REST endpoints, Web UI dashboard, and static assets.
"""

import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from .routes import chat_router, data_router, agents_router, settings_router
from ..config import settings


def create_app() -> FastAPI:
    """Builds and configures the Agent 365 FastAPI application."""
    app = FastAPI(
        title="Agent 365 Enterprise Suite",
        description="Autonomous Multi-Agent AI Suite for Microsoft 365 (Outlook, Teams, OneDrive, Calendar, Excel).",
        version=settings.version,
    )

    # CORS configuration
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Include API Routers
    app.include_router(chat_router, prefix="/api")
    app.include_router(data_router, prefix="/api")
    app.include_router(agents_router, prefix="/api")
    app.include_router(settings_router, prefix="/api")

    # Static assets and Web UI
    current_dir = os.path.dirname(os.path.abspath(__file__))
    web_dir = os.path.abspath(os.path.join(current_dir, "..", "web"))
    static_dir = os.path.join(web_dir, "static")

    if os.path.exists(static_dir):
        app.mount("/static", StaticFiles(directory=static_dir), name="static")

    @app.get("/")
    def serve_index():
        index_file = os.path.join(web_dir, "index.html")
        if os.path.exists(index_file):
            return FileResponse(index_file)
        return {"status": "ok", "app": "Agent 365 API", "docs": "/docs"}

    @app.get("/health")
    def health_check():
        return {"status": "healthy", "mode": settings.m365_mode, "version": settings.version}

    return app


app = create_app()
