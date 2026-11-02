<#
.SYNOPSIS
    Launches the Agent 365 Multi-Agent Suite for Microsoft 365.

.DESCRIPTION
    Validates Python environment prerequisites, dependencies, and executes Agent 365
    in Web Dashboard, Interactive CLI, or One-Shot Query mode.

.PARAMETER Mode
    Execution mode: 'mock' for sandbox simulation or 'live' for Microsoft Graph API.

.PARAMETER Interface
    Interface to launch: 'web' for FastAPI web dashboard, 'cli' for terminal, or 'mcp' for MCP server.

.PARAMETER Port
    Port number for the web dashboard (default: 8000).

.PARAMETER Query
    Optional one-shot query to execute without starting interactive session.

.OUTPUTS
    [PSCustomObject] with execution metadata and startup status.

.EXAMPLE
    .\start.ps1 -Interface web -Port 8000
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('mock', 'live')]
    [string]$Mode = 'mock',

    [Parameter(Mandatory = $false)]
    [ValidateSet('web', 'cli', 'mcp')]
    [string]$Interface = 'web',

    [Parameter(Mandatory = $false)]
    [int]$Port = 8000,

    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [string]$Query
)

begin {
    $ErrorActionPreference = 'Stop'
}

process {
    try {
        # Check Python installation
        $pythonCommand = Get-Command -Name python -ErrorAction SilentlyContinue
        if ($null -eq $pythonCommand) {
            throw [System.IO.FileNotFoundException]::new("Python executable was not found on PATH.")
        }

        # Determine target arguments
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'run.py'
        if (-not (Test-Path -LiteralPath $scriptPath)) {
            throw [System.IO.FileNotFoundException]::new("Launcher script 'run.py' not found at: $scriptPath")
        }

        $arguments = [System.Collections.Generic.List[string]]::new()
        $arguments.Add($scriptPath)
        $arguments.Add("--mode")
        $arguments.Add($Mode)

        if ($null -ne $Query -and $Query.Trim() -ne '') {
            $arguments.Add("--query")
            $arguments.Add($Query)
        }
        elseif ($Interface -eq 'cli') {
            $arguments.Add("--cli")
        }
        elseif ($Interface -eq 'mcp') {
            $arguments.Add("--mcp")
        }
        else {
            $arguments.Add("--port")
            $arguments.Add($Port.ToString())
        }

        $statusObject = [PSCustomObject]@{
            PSTypeName = 'Agent365.ExecutionStatus'
            Timestamp  = (Get-Date)
            Interface  = $Interface
            Mode       = $Mode
            Port       = $Port
            ScriptPath = $scriptPath
            Status     = 'Launching'
        }
        Write-Output -InputObject $statusObject

        if ($PSCmdlet.ShouldProcess("Agent 365 ($Interface mode)", "Start Process")) {
            $argArray = $arguments.ToArray()
            & python.exe @argArray
        }
    }
    catch [System.IO.FileNotFoundException] {
        Write-Error -Message "[Agent 365 Error] Required file not found: $_"
        throw
    }
    catch {
        Write-Error -Message "[Agent 365 Error] Execution failed: $_"
        throw
    }
}

end {
    # Pipeline cleanup if needed
}
