"""
Agent 365 Main Launcher.
Supports starting the Web Dashboard & API server, interactive CLI, one-shot queries, or MCP server.
"""

import sys
import argparse

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import uvicorn
from agent365.config import settings
from agent365.agents import ChiefOfStaffOrchestrator
from agent365.mcp import Agent365MCPServer


def run_interactive_cli():
    """Starts interactive terminal session with Chief of Staff."""
    print("=" * 65)
    print("  [Agent 365] Microsoft 365 Enterprise Copilot (CLI Mode)")
    print(f"  [Mode] {settings.m365_mode.upper()} | Type 'exit' to quit")
    print("=" * 65)

    orchestrator = ChiefOfStaffOrchestrator()
    while True:
        try:
            query = input("\n[Agent 365] Enter request > ").strip()
            if not query:
                continue
            if query.lower() in ("exit", "quit", "q"):
                print("Exiting Agent 365. Goodbye!")
                break

            print("\n[Orchestrating] Multi-agent workflow executing across M365...")
            resp = orchestrator.run(query)
            print("\n" + resp.final_response)
            print("\nExecution Trace:")
            for s in resp.steps_executed:
                print(f"  • {s.agent_role.value}: {s.action} ({s.duration_ms}ms) -> {s.status}")

        except (KeyboardInterrupt, EOFError):
            print("\nSession ended.")
            break


def run_one_shot_query(query: str):
    """Executes a single query and prints result."""
    orchestrator = ChiefOfStaffOrchestrator()
    resp = orchestrator.run(query)
    print(resp.final_response)


def main():
    parser = argparse.ArgumentParser(description="Agent 365: Multi-Agent Suite for Microsoft 365")
    parser.add_argument("--cli", action="store_true", help="Launch interactive CLI mode")
    parser.add_argument("--query", type=str, help="Execute a single one-shot query and exit")
    parser.add_argument("--mcp", action="store_true", help="Run as Model Context Protocol (MCP) stdio server")
    parser.add_argument("--port", type=int, default=settings.port, help=f"Web server port (default: {settings.port})")
    parser.add_argument("--host", type=str, default=settings.host, help=f"Web server host (default: {settings.host})")
    parser.add_argument("--mode", type=str, choices=["mock", "live"], default=settings.m365_mode, help="M365 execution mode")

    args = parser.parse_args()

    if args.mode:
        settings.m365_mode = args.mode

    if args.mcp:
        server = Agent365MCPServer()
        server.run_stdio()
    elif args.cli:
        run_interactive_cli()
    elif args.query:
        run_one_shot_query(args.query)
    else:
        print(f"[Startup] Starting Agent 365 Web Dashboard & API at http://{args.host}:{args.port}")
        uvicorn.run("agent365.api:app", host=args.host, port=args.port, reload=False)


if __name__ == "__main__":
    main()
