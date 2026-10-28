[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectName,
    [string]$WorkspaceRoot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([string]::IsNullOrWhiteSpace($WorkspaceRoot)) {
    $WorkspaceRoot = Split-Path -Parent $PSScriptRoot
}

function Copy-Template {
    param([string]$Source, [string]$Destination)

    foreach ($item in Get-ChildItem -LiteralPath $Source -Force) {
        if ($item.Name -eq '.git') {
            continue
        }
        if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw "Template links are not supported: $($item.FullName)"
        }
        $target = Join-Path $Destination $item.Name
        if ($item.PSIsContainer) {
            New-Item -ItemType Directory -Path $target -ErrorAction Stop | Out-Null
            Copy-Template -Source $item.FullName -Destination $target
        } else {
            Copy-Item -LiteralPath $item.FullName -Destination $target
        }
    }
}

$projectName = $ProjectName.Trim()
if ([string]::IsNullOrWhiteSpace($projectName) -or
    $projectName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
    $projectName -match '[\x00-\x1f]' -or $projectName.EndsWith('.') -or
    $projectName -in @('.', '..')) {
    throw 'Enter a valid Windows folder name, without a path.'
}

$parent = Join-Path $WorkspaceRoot 'Projects'
$template = Join-Path $parent '_Template_Excel'
$destination = Join-Path $parent "Project $projectName"
if (Test-Path -LiteralPath $destination) {
    throw "Project already exists; nothing was changed: $destination"
}
if (-not (Test-Path -LiteralPath $template -PathType Container)) {
    throw "Template not found: $template"
}
Get-Command git -ErrorAction Stop | Out-Null
& git --version
if ($LASTEXITCODE -ne 0) { throw 'Git is not available.' }

$messagePath = Join-Path $PSScriptRoot 'New-Project.messages.json'
$gitViewMessages = Get-Content -LiteralPath $messagePath -Encoding UTF8 -Raw | ConvertFrom-Json

$created = $false
try {
    New-Item -ItemType Directory -Path $destination -ErrorAction Stop | Out-Null
    $created = $true
    Copy-Template -Source $template -Destination $destination

    & git -C $destination init -b main
    if ($LASTEXITCODE -ne 0) { throw 'Git initialization failed.' }

    & git -C $destination add -A
    if ($LASTEXITCODE -ne 0) { throw 'Git staging failed.' }

    Write-Host "Created: $destination"
    Write-Host 'Project files are ready to use in VS Code.'
    Write-Host 'Initial project files have been staged in Git; no commit, remote repository, or push was created.'
    Write-Host 'Review staged files before your first commit:'
    Write-Host "git -C `"$destination`" status"
    Write-Host "git -C `"$destination`" diff --cached"
    Write-Host "git -C `"$destination`" commit -m `"Initial project setup`""

    # Replace all the recently added Japanese Write-Host statements with this loop.
    foreach ($message in $gitViewMessages) {
        Write-Host ($message.Replace('{projectPath}', $destination))
    }
} catch {
    if ($created) {
        Write-Warning "Setup failed. Partial files were kept for inspection: $destination"
    }
    throw
}
