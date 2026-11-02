# 🏢 Microsoft 365 Enterprise Suite & Autonomous Agent Copilot

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.110%2B-009688.svg)](https://fastapi.tiangolo.com/)
[![Tests Passing](https://img.shields.io/badge/tests-15%20passed-brightgreen.svg)]()
[![Anti-Slop Certified](https://img.shields.io/badge/Anti--Slop-0%20Violations%20Passed-success.svg)]()
[![High-Gloss Lacquer](https://img.shields.io/badge/Baseline-High--Gloss%20Lacquer%20Optics-gold.svg)]()

**Agent 365** is a 1-to-1 authentic web replica of the **Microsoft 365 Suite** powered by a unified multi-agent AI copilot system. Grounded in official Microsoft documentation, the suite maps every tab directly to its corresponding functional application with authentic navigation rails, ribbon bars, interactive document canvases, spreadsheet formula engines, and presentation runners.

---

## 🌟 1-to-1 Application Mapping & Capabilities

### 1. 🏠 Home / Microsoft 365 Hub
- Overview dashboard with active executive briefing, quick-create buttons ("New Document", "New Workbook", "New Presentation", "New Email").
- Live cloud storage quota meter (OneDrive Business).
- Recent cloud documents table with one-click direct application launch.

### 2. 🔵 Microsoft Word for Web
- **Ribbon Tabs**: Home, Insert, Layout, Review, View.
- **Formatting Tools**: Font family, size, Bold, Italic, Underline, Strikethrough, Alignments, Bullet & Numbered lists, Styles (Title, H1, H2, Normal).
- **Document Canvas**: True 816px document page with drop shadow, live editing (`contenteditable`), live word counter, character counter, Save to OneDrive, Print, and Export (`.docx`/`.txt`).

### 3. 🟢 Microsoft Excel for Web
- **Ribbon Tabs**: Home, Insert, Formulas, Data.
- **Formula Bar**: Interactive Cell Name Box (`B4`), Formula icon (`fx`), formula input with live evaluation engine (`=SUM(...)`, `=AVERAGE(...)`, math operations, live recalculation).
- **Spreadsheet Grid**: Column letters (A-G), row numbers (1-20), cell selection, live in-place cell editing, and department variance calculation.
- **Sheet Navigation & Export**: `Q4_Summary`, `Department_Budgets`, CSV export, and status bar metrics (Average, Count, Sum).

### 4. 🔴 Microsoft PowerPoint for Web
- **Ribbon Tabs**: Home, Insert, Design, Slide Show.
- **Thumbnails Strip**: Left slide preview reel (Slide 1, Slide 2, Slide 3, "+ New Slide", Duplicate, Delete).
- **Slide Stage**: 16:9 widescreen canvas with editable Title and Subtitle/Bullet cards.
- **Slide Show Presentation Runner**: Fullscreen presentation runner with transition animations, Next/Previous controls, and keyboard arrow listeners (`LeftArrow`, `RightArrow`, `Escape`).

### 5. 🔵 Microsoft Outlook for Web
- **Top Command Bar**: "New mail", "Delete", "Archive", "Mark as read", "Flag", "Move".
- **Folder Navigation**: Favorites (Inbox, Sent Items, Drafts), Folders (Archive, Deleted Items).
- **Message List**: Filter tabs (`All` vs `Unread`), unread indicators, sender avatars, subjects, previews, and flag actions.
- **Reading Pane & Rich Composer**: Read full messages or switch into Composer mode to draft and dispatch emails that update tenant state live.

### 6. 🟣 Microsoft Teams for Web
- **Left Activity Rail**: Activity, Chat, Teams, Files.
- **Channel Tree**: Collapsible team trees (`Enterprise 365`, `Engineering Core`, `Executive Leadership`) with channels (`#General`, `#Releases`, `#Strategy`).
- **Discussion Stream**: Sender avatars, timestamps, message cards, reaction pills (Like, Heart), and an interactive message composer.

### 7. 🔵 Microsoft OneDrive for Web
- **Command Bar**: "+ New" file, Upload, Share link, Download.
- **Breadcrumb Path**: `My files > Documents > Q4`.
- **File Explorer**: File type vector icons, file name, modified date, size, and sharing status.

### 8. 🤖 Copilot / Chief of Staff Orchestrator
- Autonomous multi-agent coordinator that parses cross-app instructions, plans execution steps, invokes specialist agents, and displays real-time execution steps with millisecond timings.

---

## 🎨 Baseline Build & Anti-Slop Craftsman Integration

- **High-Gloss Lacquer System**:
  - Physical material optics: Deep Lacquer Navy (`#050B18`), Cobalt Glaze (`#0E1D3A`), Polished Camel (`#C5A059`), Sunlit Tan (`#DFBA73`).
  - Curved Specular Glare Shield (`::before`) and dynamic Pointer-Tracked Specular Flare (`::after` with `--mouse-x`, `--mouse-y`).
  - Still Webpage Background Enforcement with wet beveled drop shadows.
- **Zero-Tolerance Anti-Slop Standard**:
  - **Zero Emoji Asset Flips**: 100% custom SVG and Fluent/Bootstrap vector silhouettes with elevation shadows and contrast ratio $\ge 4.5:1$.
  - **Zero Purple Prose**: Pure, concise, mechanically-driven copy.
  - **Zero Shovelware Ad-Traps**: Clean enterprise workflow.
  - **Tactile Kinesthetics**: Minimum $48\times 48\text{ dp}$ hit targets, active press states, responsive layout.
  - Passed `Invoke-AntiSlopAudit.ps1` with 0 violations.

---

## 🚀 Quickstart

### 1. Launch the Web Suite
```powershell
Set-Location "G:\My Drive\Antigravity\Projects\agent 365"
.\start.ps1 -Interface web -Port 8000
```
Open **`http://localhost:8000`** in your browser.

### 2. Interactive Terminal CLI
```powershell
.\start.ps1 -Interface cli
```

### 3. One-Shot Command Execution
```powershell
.\start.ps1 -Query "Prepare my morning briefing: check urgent emails and review calendar conflicts"
```

### 4. Run Pytest Suite
```powershell
python -m pytest -v
```

---

## 📁 Project Structure

```
agent 365/
├── agent365/
│   ├── config.py                 # Configuration and environment settings
│   ├── models/                   # Pydantic schemas for M365 entities
│   │   ├── email.py              # Mail models
│   │   ├── calendar.py           # Calendar & conflict models
│   │   ├── storage.py            # OneDrive models
│   │   ├── teams.py              # Teams chat & channel models
│   │   ├── office.py             # Word, Excel, and PowerPoint models
│   │   └── agent_state.py        # Multi-agent trace and plan models
│   ├── m365_client/              # Unified client layer
│   │   ├── base.py               # Abstract base class
│   │   ├── mock_client.py        # Rich in-memory simulated tenant
│   │   ├── graph_client.py       # Live Graph API client (MSAL)
│   │   └── factory.py            # Client provider factory
│   ├── agents/                   # Specialized AI Agents
│   │   ├── base_agent.py         # Base agent with tool registry & tracing
│   │   ├── mail_agent.py         # Outlook Mail specialist
│   │   ├── calendar_agent.py     # Calendar & conflict specialist
│   │   ├── storage_agent.py      # OneDrive & files specialist
│   │   ├── teams_agent.py        # Microsoft Teams specialist
│   │   ├── office_agent.py       # Excel & document specialist
│   │   └── orchestrator.py       # Chief of Staff master orchestrator
│   ├── mcp/                      # Model Context Protocol
│   │   └── server.py             # MCP JSON-RPC stdio server
│   ├── api/                      # REST API endpoints (FastAPI)
│   │   ├── app.py                # FastAPI application factory
│   │   └── routes/               # Modular route controllers
│   └── web/                      # 1-to-1 Microsoft 365 Web Suite
│       ├── index.html            # Complete authentic web application
│       └── static/
│           ├── css/
│           │   ├── high-gloss-lacquer.css  # Baseline optics & material depth
│           │   └── style.css               # M365 authentic styling & ribbon layout
│           └── js/
│               ├── high-gloss-lacquer.js   # Pointer flare & sheen engine
│               └── app.js                  # M365 application controller
├── tests/                        # Pytest automated test suite (15 passed)
├── run.py                        # Unified CLI & server launcher
├── start.ps1                     # Modern PowerShell automation script
├── requirements.txt              # Python package dependencies
├── pyproject.toml                # Build & pytest configuration
└── .env.example                  # Template configuration file
```
