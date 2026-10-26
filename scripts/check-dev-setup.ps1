#Requires -Version 5.1
<#
.SYNOPSIS
    Checks that this machine is set up to develop the add-in. Changes nothing.

.DESCRIPTION
    Verifies Node.js, Corepack, pnpm, Git, the installed dependencies, the lefthook
    Git hooks, the trusted localhost certificate and the recommended VS Code
    extensions. Prints how to fix each problem and exits with code 1 if a required
    check fails.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\check-dev-setup.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$onWindows = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
$script:failed = 0
$script:warned = 0

function Write-Check {
    param(
        [Parameter(Mandatory)] [ValidateSet('OK', 'WARN', 'FAIL', 'SKIP')] [string] $Status,
        [Parameter(Mandatory)] [string] $Message,
        [string] $Fix
    )
    $color = @{ OK = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; SKIP = 'DarkGray' }[$Status]
    Write-Host ('[{0,-4}] ' -f $Status) -ForegroundColor $color -NoNewline
    Write-Host $Message
    if ($Fix) { Write-Host "       -> $Fix" -ForegroundColor DarkGray }
    if ($Status -eq 'FAIL') { $script:failed++ }
    if ($Status -eq 'WARN') { $script:warned++ }
}

# Runs a native command in the repository root and returns its first output line,
# or $null when the command is missing or fails.
function Invoke-Tool {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [string[]] $Arguments = @('--version')
    )
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) { return $null }

    # Windows PowerShell turns native stderr into terminating errors under 'Stop'.
    $ErrorActionPreference = 'Continue'
    Push-Location $repoRoot
    try {
        $output = & $Name @Arguments 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        return $output | Select-Object -First 1
    }
    finally {
        Pop-Location
    }
}

$package = Get-Content (Join-Path $repoRoot 'package.json') -Raw | ConvertFrom-Json
$expectedPnpm = ($package.packageManager -split '@')[1]
$minimumNode = [version]($package.engines.node -replace '[^\d.]', '')
$nodeMajor = [int](Get-Content (Join-Path $repoRoot '.nvmrc') -Raw).Trim()

Write-Host "Checking the development setup in $repoRoot"
Write-Host ''

# Node.js
$nodeVersion = Invoke-Tool 'node'
if (-not $nodeVersion) {
    Write-Check FAIL 'Node.js is not installed.' "Install Node.js $nodeMajor from https://nodejs.org (or with fnm: fnm install $nodeMajor)."
}
else {
    $node = [version]$nodeVersion.TrimStart('v')
    if ($node -lt $minimumNode) {
        Write-Check FAIL "Node.js $node is too old, the project needs $minimumNode or newer." "Install Node.js $nodeMajor from https://nodejs.org (or with fnm: fnm install $nodeMajor)."
    }
    elseif ($node.Major -ne $nodeMajor) {
        Write-Check WARN "Node.js $node is installed, the project targets $nodeMajor.x (.nvmrc)." "Switch to Node.js $nodeMajor, for example with fnm: fnm use $nodeMajor"
    }
    else {
        Write-Check OK "Node.js $node"
    }
}

# Corepack provides the pnpm version pinned in package.json#packageManager.
$corepackVersion = Invoke-Tool 'corepack'
if ($corepackVersion) {
    Write-Check OK "Corepack $corepackVersion"
}
elseif ($nodeVersion) {
    Write-Check FAIL 'Corepack is not available.' 'npm install --global corepack'
}
else {
    Write-Check SKIP 'Corepack: install Node.js first.'
}

# pnpm
$previousPrompt = $env:COREPACK_ENABLE_DOWNLOAD_PROMPT
$env:COREPACK_ENABLE_DOWNLOAD_PROMPT = '0'
try {
    $pnpmVersion = Invoke-Tool 'pnpm'
}
finally {
    $env:COREPACK_ENABLE_DOWNLOAD_PROMPT = $previousPrompt
}
if (-not $pnpmVersion) {
    Write-Check FAIL 'pnpm is not available.' 'corepack enable'
}
elseif ($pnpmVersion -ne $expectedPnpm) {
    Write-Check FAIL "pnpm $pnpmVersion is active, the project pins $expectedPnpm." 'corepack enable (Corepack then uses the pinned version)'
}
else {
    Write-Check OK "pnpm $pnpmVersion"
}

# Git
$gitVersion = Invoke-Tool 'git'
if ($gitVersion) {
    Write-Check OK ($gitVersion -replace '^git version ', 'Git ')
}
else {
    Write-Check FAIL 'Git is not installed.' 'winget install Git.Git'
}

# Dependencies: pnpm compares node_modules with package.json and pnpm-lock.yaml.
if (-not $pnpmVersion) {
    Write-Check SKIP 'Dependencies: install pnpm first.'
}
elseif (-not (Test-Path (Join-Path $repoRoot 'node_modules/.modules.yaml'))) {
    Write-Check FAIL 'Dependencies are not installed.' 'pnpm install'
}
elseif ($null -eq (Invoke-Tool 'pnpm' @('--config.verify-deps-before-run=error', 'exec', 'node', '--version'))) {
    Write-Check FAIL 'Dependencies are out of date with package.json or pnpm-lock.yaml.' 'pnpm install'
}
else {
    Write-Check OK 'Dependencies are installed and up to date'
}

# Git hooks (lefthook's postinstall installs them during pnpm install)
$hooksPath = Invoke-Tool 'git' @('rev-parse', '--git-path', 'hooks')
if (-not $gitVersion -or -not $hooksPath) {
    Write-Check SKIP 'Git hooks: not a Git checkout, or Git is missing.'
}
else {
    if (-not [System.IO.Path]::IsPathRooted($hooksPath)) { $hooksPath = Join-Path $repoRoot $hooksPath }
    $preCommit = Join-Path $hooksPath 'pre-commit'
    if ((Test-Path $preCommit) -and (Select-String -Path $preCommit -Pattern 'lefthook' -Quiet)) {
        Write-Check OK 'lefthook Git hooks are installed'
    }
    else {
        Write-Check FAIL 'lefthook Git hooks are not installed.' 'pnpm exec lefthook install'
    }
}

# Localhost certificate: Outlook only loads add-ins over trusted HTTPS.
$certDir = Join-Path $HOME '.office-addin-dev-certs'
$caPath = Join-Path $certDir 'ca.crt'
$certPath = Join-Path $certDir 'localhost.crt'
if (-not ((Test-Path $caPath) -and (Test-Path $certPath) -and (Test-Path (Join-Path $certDir 'localhost.key')))) {
    Write-Check FAIL 'The localhost development certificate is missing.' 'pnpm certs:install'
}
else {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $certPath
    $ca = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $caPath
    $trusted = -not $onWindows -or @(
        Get-ChildItem Cert:\CurrentUser\Root, Cert:\LocalMachine\Root |
            Where-Object { $_.Thumbprint -eq $ca.Thumbprint }
    ).Count -gt 0

    if (-not $trusted) {
        Write-Check FAIL 'The development CA certificate is not trusted.' 'pnpm certs:install'
    }
    elseif ($cert.NotAfter -lt (Get-Date)) {
        Write-Check WARN "The localhost certificate expired on $($cert.NotAfter.ToString('yyyy-MM-dd'))." 'pnpm certs:install (pnpm dev also renews it)'
    }
    else {
        Write-Check OK "Localhost certificate is trusted, valid until $($cert.NotAfter.ToString('yyyy-MM-dd'))"
    }
}

# Recommended VS Code extensions (optional)
$codeCli = @('code', 'code-insiders') | Where-Object { Get-Command $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
if (-not $codeCli) {
    Write-Check WARN 'VS Code command line (code) not found, extensions not checked.' 'Install VS Code with "Add to PATH" enabled.'
}
else {
    $recommended = @((Get-Content (Join-Path $repoRoot '.vscode/extensions.json') -Raw | ConvertFrom-Json).recommendations)
    $ErrorActionPreference = 'Continue'
    $installed = @(& $codeCli --list-extensions 2>$null)
    $ErrorActionPreference = 'Stop'
    $missing = @($recommended | Where-Object { $installed -notcontains $_ })
    if ($missing.Count -gt 0) {
        Write-Check WARN "Recommended VS Code extensions are missing: $($missing -join ', ')" "$codeCli --install-extension $($missing -join " --install-extension ")"
    }
    else {
        Write-Check OK 'Recommended VS Code extensions are installed'
    }
}

Write-Host ''
if ($script:failed -gt 0) {
    Write-Host "Setup incomplete: $($script:failed) check(s) failed." -ForegroundColor Red
    exit 1
}
if ($script:warned -gt 0) {
    Write-Host "Setup complete, with $($script:warned) warning(s)." -ForegroundColor Yellow
}
else {
    Write-Host 'Setup complete.' -ForegroundColor Green
}
exit 0
