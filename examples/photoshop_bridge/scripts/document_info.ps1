param(
    [bool]$ProbeCom = $true
)

$ErrorActionPreference = "Stop"

function Write-JsonResult {
    param([hashtable]$Result)
    $Result | ConvertTo-Json -Depth 12
}

if (-not $ProbeCom) {
    Write-JsonResult @{
        ok = $false
        bridge = "photoshop"
        task = "document_info"
        active_document = $false
        warnings = @("COM probing was skipped by request.")
        next_steps = @("Run without ProbeCom=false on a Windows machine with Photoshop available.")
    }
    exit 0
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$jsxPath = Join-Path $repoRoot "examples\photoshop_bridge\jsx\document_info.jsx"

try {
    $app = New-Object -ComObject Photoshop.Application
    $jsx = Get-Content -Raw -Path $jsxPath
    $raw = $app.DoJavaScript($jsx)
    $raw | ConvertFrom-Json | ConvertTo-Json -Depth 12
} catch {
    Write-JsonResult @{
        ok = $false
        bridge = "photoshop"
        task = "document_info"
        active_document = $false
        width = $null
        height = $null
        color_mode = $null
        layer_count = 0
        path = $null
        warnings = @("Could not connect to Photoshop.Application COM or run the document info JSX.")
        next_steps = @(
            "Start an authorized Photoshop desktop session.",
            "Open a document manually if you want current document metadata.",
            "Run npm.cmd run photoshop:info again."
        )
        error_type = $_.Exception.GetType().Name
    }
    exit 0
}
