<#
.SYNOPSIS
    Installs or updates Bob — the agent orchestrator — and its CLI agent prerequisites.

.DESCRIPTION
    Downloads and installs the latest (or specified) version of Bob from GitHub releases.
    Supports silent installation, architecture detection, version comparison, and checksum verification.
    After installing Bob, automatically installs/verifies CLI tools: Git, Node.js, GitHub CLI,
    Claude Code, Copilot CLI, Codex CLI, and Gemini CLI.

.PARAMETER Version
    Install a specific version (e.g., "1.2.0" or "v1.2.0"). Defaults to latest.

.PARAMETER Force
    Reinstall even if the current version is already up-to-date.

.PARAMETER SkipPrereqs
    Skip installation of CLI agent prerequisites (Git, Node.js, GitHub CLI, Claude Code, etc.).

.EXAMPLE
    # Install latest version:
    irm https://raw.githubusercontent.com/saketlunker/bob-releases/main/install.ps1 | iex

    # Install a specific version:
    & { param($Version) irm https://raw.githubusercontent.com/saketlunker/bob-releases/main/install.ps1 | iex } -Version "1.2.0"
#>
param(
    [string]$Version,
    [switch]$Force,
    [switch]$SkipPrereqs
)

$ErrorActionPreference = 'Stop'

# ─── Constants ────────────────────────────────────────────────────────────────

$GH_OWNER    = 'saketlunker'
$GH_REPO     = 'bob-releases'
$API_BASE    = "https://api.github.com/repos/$GH_OWNER/$GH_REPO"
$PRODUCT     = 'Bob'
$EXE_NAME    = 'Bob.exe'

# Candidate install paths (NSIS defaults)
$INSTALL_PATHS = @(
    [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Programs', $PRODUCT, $EXE_NAME)
    [System.IO.Path]::Combine($env:ProgramFiles, $PRODUCT, $EXE_NAME)
    [System.IO.Path]::Combine(${env:ProgramFiles(x86)}, $PRODUCT, $EXE_NAME)
)

# ─── Helpers ──────────────────────────────────────────────────────────────────

function Write-Status {
    param(
        [string]$Label,
        [string]$Value,
        [string]$Color = 'White'
    )
    $padding = 26 - $Label.Length
    if ($padding -lt 1) { $padding = 1 }
    $dots = '.' * $padding
    Write-Host "  $Label" -NoNewline
    Write-Host "$dots " -ForegroundColor DarkGray -NoNewline
    Write-Host $Value -ForegroundColor $Color
}

function Write-Step {
    param([string]$Message)
    Write-Host "  > " -ForegroundColor DarkCyan -NoNewline
    Write-Host $Message
}

function Write-Success {
    param([string]$Message)
    Write-Host ""
    Write-Host "  [OK] " -ForegroundColor Green -NoNewline
    Write-Host $Message
}

function Write-Fail {
    param([string]$Message)
    Write-Host ""
    Write-Host "  [ERROR] " -ForegroundColor Red -NoNewline
    Write-Host $Message
    Write-Host ""
}

function Write-Warn {
    param([string]$Message)
    Write-Host "  [WARN] " -ForegroundColor Yellow -NoNewline
    Write-Host $Message
}

function Show-Banner {
    Write-Host ""
    Write-Host "  +==============================+" -ForegroundColor Cyan
    Write-Host "  |     " -ForegroundColor Cyan -NoNewline
    Write-Host "Bob Installer" -ForegroundColor White -NoNewline
    Write-Host "            |" -ForegroundColor Cyan
    Write-Host "  +==============================+" -ForegroundColor Cyan
    Write-Host "  |  " -ForegroundColor Cyan -NoNewline
    Write-Host "Bob + CLI agents" -ForegroundColor DarkGray -NoNewline
    Write-Host "          |" -ForegroundColor Cyan
    Write-Host "  +==============================+" -ForegroundColor Cyan
    Write-Host ""
}

function Get-Architecture {
    $arch = $null

    # PowerShell 7+ exposes $env:PROCESSOR_ARCHITECTURE reliably, but we also
    # check the .NET runtime architecture for the most robust detection.
    if ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq [System.Runtime.InteropServices.Architecture]::Arm64) {
        $arch = 'arm64'
    }
    elseif ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq [System.Runtime.InteropServices.Architecture]::X64) {
        $arch = 'x64'
    }

    if (-not $arch) {
        # Fallback for older PowerShell
        $envArch = $env:PROCESSOR_ARCHITECTURE
        switch ($envArch) {
            'AMD64' { $arch = 'x64' }
            'ARM64' { $arch = 'arm64' }
            'x86'   { $arch = 'x64' }   # 32-bit PS on 64-bit OS
            default  { $arch = 'x64' }   # best guess
        }
    }

    return $arch
}

function Find-ExistingInstall {
    foreach ($path in $INSTALL_PATHS) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

function Get-InstalledVersion {
    param([string]$ExePath)
    if (-not $ExePath -or -not (Test-Path $ExePath)) {
        return $null
    }
    try {
        $info = (Get-Item $ExePath).VersionInfo
        $ver = $info.ProductVersion
        if (-not $ver) { $ver = $info.FileVersion }
        if ($ver) {
            # Strip any metadata suffix (e.g. "1.0.0+build123")
            return ($ver -split '[+]')[0].Trim()
        }
    }
    catch {
        return $null
    }
    return $null
}

function Normalize-Version {
    param([string]$Ver)
    if (-not $Ver) { return $null }
    $v = $Ver.Trim()
    if ($v.StartsWith('v')) { $v = $v.Substring(1) }
    return $v
}

function Compare-SemVer {
    <#
        Returns:
         -1 if Left < Right
          0 if Left == Right
          1 if Left > Right
    #>
    param([string]$Left, [string]$Right)
    try {
        $l = [System.Version]::new($Left)
        $r = [System.Version]::new($Right)
        return $l.CompareTo($r)
    }
    catch {
        # Fall back to string comparison
        return [string]::Compare($Left, $Right, [System.StringComparison]::Ordinal)
    }
}

function Get-LatestRelease {
    try {
        # Use TLS 1.2+; required for GitHub API on older PowerShell
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    $headers = @{
        'Accept'     = 'application/vnd.github+json'
        'User-Agent' = 'Bob-Installer/1.0'
    }

    $url = "$API_BASE/releases/latest"
    try {
        $release = Invoke-RestMethod -Uri $url -Headers $headers -UseBasicParsing
        return $release
    }
    catch {
        throw "Failed to fetch latest release from GitHub.`n  URL: $url`n  $_"
    }
}

function Get-SpecificRelease {
    param([string]$Tag)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    $headers = @{
        'Accept'     = 'application/vnd.github+json'
        'User-Agent' = 'Bob-Installer/1.0'
    }

    # Try with and without 'v' prefix
    $tags = @($Tag, "v$Tag")
    if ($Tag.StartsWith('v')) {
        $tags = @($Tag, $Tag.Substring(1))
    }

    foreach ($t in $tags) {
        $url = "$API_BASE/releases/tags/$t"
        try {
            $release = Invoke-RestMethod -Uri $url -Headers $headers -UseBasicParsing
            return $release
        }
        catch {
            continue
        }
    }

    throw "Release '$Tag' not found in $GH_OWNER/$GH_REPO. Check the version and try again."
}

function Find-InstallerAsset {
    param($Release, [string]$Arch)

    $assets = $Release.assets

    # Primary pattern: Bob-{version}-setup.exe (NSIS output)
    # Try architecture-specific first, then fall back to universal
    $patterns = @(
        "Bob-*-setup-$Arch.exe"
        "Bob-*-setup.exe"
        "Bob Setup *-$Arch.exe"
        "Bob Setup *.exe"
        "Bob-*-$Arch.exe"
    )

    foreach ($pattern in $patterns) {
        $match = $assets | Where-Object { $_.name -like $pattern } | Select-Object -First 1
        if ($match) { return $match }
    }

    # Last resort: any .exe that isn't a blockmap
    $match = $assets | Where-Object { $_.name -like '*.exe' -and $_.name -notlike '*.blockmap*' } | Select-Object -First 1
    if ($match) { return $match }

    throw "No installer asset found for architecture '$Arch' in release $($Release.tag_name).`n  Available assets: $(($assets | ForEach-Object { $_.name }) -join ', ')"
}

function Get-ChecksumFromRelease {
    <#
        Attempts to find a SHA512 checksum for the installer from the latest.yml
        asset or a .sha256 / .sha512 sidecar file in the release.
    #>
    param($Release, [string]$InstallerName)

    $headers = @{
        'Accept'     = 'application/octet-stream'
        'User-Agent' = 'Bob-Installer/1.0'
    }

    # Try latest.yml (electron-builder publishes this)
    $ymlAsset = $Release.assets | Where-Object { $_.name -eq 'latest.yml' } | Select-Object -First 1
    if ($ymlAsset) {
        try {
            $ymlContent = Invoke-RestMethod -Uri $ymlAsset.browser_download_url -Headers @{ 'User-Agent' = 'Bob-Installer/1.0' } -UseBasicParsing
            $ymlText = if ($ymlContent -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($ymlContent) } else { "$ymlContent" }

            # Parse the YAML-like structure for sha512 of our file
            # latest.yml has blocks like:
            #   - url: Bob-1.0.0-setup.exe
            #     sha512: <base64hash>
            #     size: 12345
            $lines = $ymlText -split "`n"
            $foundFile = $false
            foreach ($line in $lines) {
                if ($line -match '^\s*-?\s*url:\s*(.+)$') {
                    $url = $Matches[1].Trim()
                    $foundFile = ($url -eq $InstallerName)
                }
                if ($foundFile -and $line -match '^\s*sha512:\s*(.+)$') {
                    $sha512Base64 = $Matches[1].Trim()
                    $sha512Bytes = [System.Convert]::FromBase64String($sha512Base64)
                    $sha512Hex = ($sha512Bytes | ForEach-Object { $_.ToString('x2') }) -join ''
                    return @{ Algorithm = 'SHA512'; Hash = $sha512Hex.ToUpper() }
                }
            }
        }
        catch {
            # Checksum retrieval is best-effort
        }
    }

    # Try .sha256 sidecar
    $sha256Asset = $Release.assets | Where-Object { $_.name -eq "$InstallerName.sha256" } | Select-Object -First 1
    if ($sha256Asset) {
        try {
            $content = Invoke-RestMethod -Uri $sha256Asset.browser_download_url -Headers @{ 'User-Agent' = 'Bob-Installer/1.0' } -UseBasicParsing
            $hash = ("$content" -split '\s')[0].Trim().ToUpper()
            if ($hash.Length -eq 64) {
                return @{ Algorithm = 'SHA256'; Hash = $hash }
            }
        }
        catch { }
    }

    return $null
}

function Invoke-Download {
    param(
        [string]$Url,
        [string]$OutFile,
        [long]$ExpectedSize
    )

    $headers = @{ 'User-Agent' = 'Bob-Installer/1.0' }

    # Use .NET WebClient for progress on PS 5.1, Invoke-WebRequest on PS 7+
    $isPSCore = $PSVersionTable.PSVersion.Major -ge 7

    if ($isPSCore) {
        # PowerShell 7+ has native progress with Invoke-WebRequest
        $ProgressPreference = 'Continue'
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -Headers $headers -UseBasicParsing
    }
    else {
        # PowerShell 5.1: use System.Net.WebClient for better progress
        try {
            $wc = [System.Net.WebClient]::new()
            $wc.Headers.Add('User-Agent', 'Bob-Installer/1.0')

            $downloadComplete = $false
            $lastPercent = -1

            $progressHandler = {
                param($sender, $e)
                if ($e.ProgressPercentage -ne $script:lastPercent) {
                    $script:lastPercent = $e.ProgressPercentage
                    $filled = [math]::Floor($e.ProgressPercentage / 4)
                    $empty  = 25 - $filled
                    $bar    = ('#' * $filled) + ('-' * $empty)
                    $sizeMB = [math]::Round($e.TotalBytesToReceive / 1MB, 1)
                    $recvMB = [math]::Round($e.BytesReceived / 1MB, 1)
                    Write-Host "`r  [$bar] $($e.ProgressPercentage)%  ($recvMB / $sizeMB MB)" -NoNewline
                }
            }
            $completedHandler = {
                param($sender, $e)
                $script:downloadComplete = $true
            }

            $wc.add_DownloadProgressChanged($progressHandler)
            $wc.add_DownloadFileCompleted($completedHandler)

            $wc.DownloadFileAsync([uri]$Url, $OutFile)

            while (-not $downloadComplete) {
                Start-Sleep -Milliseconds 250
            }

            Write-Host ""  # newline after progress bar

            if ($wc.IsBusy) { $wc.CancelAsync() }
            $wc.Dispose()
        }
        catch {
            # Fallback to Invoke-WebRequest
            Write-Warn "WebClient download failed, falling back to Invoke-WebRequest..."
            $ProgressPreference = 'Continue'
            Invoke-WebRequest -Uri $Url -OutFile $OutFile -Headers $headers -UseBasicParsing
        }
    }
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N1} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# ─── Prerequisites ─────────────────────────────────────────────────────────────

function Refresh-Path {
    $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')
}

function Test-CommandExists {
    param([string]$Command)
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    $result = $false
    try {
        if (Get-Command $Command -ErrorAction SilentlyContinue) { $result = $true }
    } catch {}
    $ErrorActionPreference = $old
    return $result
}

function Get-CommandVersion {
    param([string]$Command, [string[]]$Args = @('--version'))
    try {
        $output = & $Command @Args 2>&1 | Out-String
        if ($output -match '(\d+\.\d+[\.\d]*)') {
            return $Matches[1]
        }
    } catch {}
    return $null
}

function Add-ToUserPath {
    <#
        Adds a directory to the user PATH permanently and for the current session.
        Returns $true if the path was added, $false if it was already present.
    #>
    param([string]$Directory)
    if (-not $Directory -or -not (Test-Path $Directory)) { return $false }
    # Check if already in current session PATH
    $pathDirs = $env:Path -split ';' | Where-Object { $_ }
    foreach ($d in $pathDirs) {
        try {
            if ([System.IO.Path]::GetFullPath($d) -eq [System.IO.Path]::GetFullPath($Directory)) {
                return $false
            }
        } catch { }
    }
    # Add to user PATH permanently
    $userPath = [System.Environment]::GetEnvironmentVariable('Path', 'User')
    if (-not $userPath) { $userPath = '' }
    $userDirs = $userPath -split ';' | Where-Object { $_ }
    $alreadyPersisted = $false
    foreach ($d in $userDirs) {
        try {
            if ([System.IO.Path]::GetFullPath($d) -eq [System.IO.Path]::GetFullPath($Directory)) {
                $alreadyPersisted = $true
                break
            }
        } catch { }
    }
    if (-not $alreadyPersisted) {
        $newUserPath = if ($userPath) { $userPath + ';' + $Directory } else { $Directory }
        [System.Environment]::SetEnvironmentVariable('Path', $newUserPath, 'User')
    }
    # Add to current session
    $env:Path += ';' + $Directory
    return $true
}

function Get-NpmGlobalBin {
    <# Returns the npm global bin directory, or $null. #>
    if (-not (Test-CommandExists 'npm')) { return $null }
    try {
        $prefix = (& npm prefix -g 2>&1 | Out-String).Trim()
        if ($prefix -and (Test-Path $prefix)) { return $prefix }
    } catch {}
    # Fallback: common Windows location
    $fallback = Join-Path $env:APPDATA 'npm'
    if (Test-Path $fallback) { return $fallback }
    return $null
}

function Find-InKnownLocations {
    <#
        Searches known install directories for a tool binary.
        Returns the directory containing the binary, or $null.
    #>
    param([string]$ToolName)

    $candidates = @()
    switch ($ToolName) {
        'git' {
            $candidates = @(
                (Join-Path $env:ProgramFiles 'Git\cmd')
                (Join-Path ${env:ProgramFiles(x86)} 'Git\cmd')
                (Join-Path $env:LOCALAPPDATA 'Programs\Git\cmd')
            )
        }
        'node' {
            $candidates = @(
                (Join-Path $env:ProgramFiles 'nodejs')
                (Join-Path ${env:ProgramFiles(x86)} 'nodejs')
                (Join-Path $env:LOCALAPPDATA 'Programs\nodejs')
            )
        }
        'gh' {
            $candidates = @(
                (Join-Path $env:ProgramFiles 'GitHub CLI')
                (Join-Path ${env:ProgramFiles(x86)} 'GitHub CLI')
                (Join-Path $env:LOCALAPPDATA 'Programs\GitHub CLI')
            )
        }
        'claude' {
            $candidates = @(
                (Join-Path $env:USERPROFILE '.local\bin')
                (Join-Path $env:LOCALAPPDATA 'Programs\claude')
                (Join-Path $env:APPDATA 'Claude\bin')
            )
        }
        { $_ -in @('copilot', 'codex', 'gemini') } {
            # npm global bin directories
            $npmBin = Get-NpmGlobalBin
            if ($npmBin) {
                $candidates = @($npmBin)
            }
            $roamingNpm = Join-Path $env:APPDATA 'npm'
            if ($roamingNpm -and ($candidates -notcontains $roamingNpm)) {
                $candidates += $roamingNpm
            }
        }
    }

    $exeName = "$ToolName.exe"
    # For node, npm ships as node.exe; for gh, binary is gh.exe; etc.
    foreach ($dir in $candidates) {
        if (-not $dir) { continue }
        $fullPath = Join-Path $dir $exeName
        if (Test-Path $fullPath) {
            return $dir
        }
        # Also check .cmd shims (npm installs .cmd wrappers)
        $cmdPath = Join-Path $dir "$ToolName.cmd"
        if (Test-Path $cmdPath) {
            return $dir
        }
    }
    return $null
}

function Resolve-Tool {
    <#
        Full resolution flow for a tool:
        1. Check PATH via Get-Command
        2. If not found, check known install locations
        3. If found off-PATH, fix PATH permanently + session
        Returns a hashtable: @{ Found = $bool; FixedPath = $bool; Dir = $string }
    #>
    param([string]$Command)

    # Step 1: already in PATH
    if (Test-CommandExists $Command) {
        return @{ Found = $true; FixedPath = $false; Dir = $null }
    }

    # Step 2: check known locations
    $dir = Find-InKnownLocations -ToolName $Command
    if ($dir) {
        # Step 3: fix PATH
        Add-ToUserPath -Directory $dir | Out-Null
        # Verify it works now
        if (Test-CommandExists $Command) {
            return @{ Found = $true; FixedPath = $true; Dir = $dir }
        }
    }

    return @{ Found = $false; FixedPath = $false; Dir = $null }
}

function Test-WingetAvailable {
    return (Test-CommandExists 'winget')
}

function Install-WithWinget {
    param([string]$PackageId, [string]$DisplayName)
    if (-not (Test-WingetAvailable)) {
        Write-Warn "$DisplayName not found. Install winget (aka App Installer) from the Microsoft Store, then install $DisplayName manually."
        return $false
    }
    try {
        Write-Step "Installing $DisplayName via winget..."
        $output = & winget install $PackageId --accept-source-agreements --accept-package-agreements 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0 -and $output -notmatch 'already installed') {
            Write-Warn "winget install $PackageId exited with code $LASTEXITCODE"
            return $false
        }
        Refresh-Path
        return $true
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Warn "Failed to install $DisplayName`: $errMsg"
        return $false
    }
}

function Install-WithNpm {
    param([string]$PackageName, [string]$DisplayName)
    if (-not (Test-CommandExists 'npm')) {
        Write-Warn "npm not available - cannot install $DisplayName. Install Node.js first."
        return $false
    }
    try {
        Write-Step "Installing $DisplayName via npm..."
        $output = & npm install -g $PackageName 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            Write-Warn "npm install -g $PackageName failed (exit code $LASTEXITCODE)"
            return $false
        }
        # Ensure npm global bin is in PATH permanently
        $npmBin = Get-NpmGlobalBin
        if ($npmBin) {
            Add-ToUserPath -Directory $npmBin | Out-Null
        }
        return $true
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Warn "Failed to install $DisplayName`: $errMsg"
        return $false
    }
}

function Install-Prerequisites {
    Write-Host ""
    Write-Host "  ── Prerequisites ──────────────────────────────" -ForegroundColor Cyan
    Write-Host ""

    $script:prereqResults = @()

    # ── 1. Git ────────────────────────────────────────────────────────────────
    $res = Resolve-Tool -Command 'git'
    if ($res.Found -and $res.FixedPath) {
        $ver = Get-CommandVersion 'git' @('--version')
        Write-Status "Git" "v$ver (found, fixed PATH)" -Color Yellow
        $script:prereqResults += @{ Name = 'Git'; Version = "v$ver"; Status = 'path-fixed' }
    }
    elseif ($res.Found) {
        $ver = Get-CommandVersion 'git' @('--version')
        Write-Status "Git" "v$ver (already installed)" -Color Green
        $script:prereqResults += @{ Name = 'Git'; Version = "v$ver"; Status = 'present' }
    }
    else {
        if (Install-WithWinget 'Git.Git' 'Git') {
            Refresh-Path
            # Post-install: resolve again in case winget didn't update PATH
            $post = Resolve-Tool -Command 'git'
            $ver = Get-CommandVersion 'git' @('--version')
            if ($ver) {
                Write-Status "Git" "v$ver (installed)" -Color Green
                $script:prereqResults += @{ Name = 'Git'; Version = "v$ver"; Status = 'installed' }
            }
            else {
                Write-Status "Git" "installed but not in PATH" -Color Yellow
                $script:prereqResults += @{ Name = 'Git'; Version = $null; Status = 'installed' }
            }
        }
        else {
            Write-Status "Git" "not installed" -Color Red
            $script:prereqResults += @{ Name = 'Git'; Version = $null; Status = 'failed' }
        }
    }

    # ── 2. Node.js ────────────────────────────────────────────────────────────
    $res = Resolve-Tool -Command 'node'
    if ($res.Found -and $res.FixedPath) {
        $ver = Get-CommandVersion 'node' @('--version')
        Write-Status "Node.js" "v$ver (found, fixed PATH)" -Color Yellow
        $script:prereqResults += @{ Name = 'Node.js'; Version = "v$ver"; Status = 'path-fixed' }
    }
    elseif ($res.Found) {
        $ver = Get-CommandVersion 'node' @('--version')
        Write-Status "Node.js" "v$ver (already installed)" -Color Green
        $script:prereqResults += @{ Name = 'Node.js'; Version = "v$ver"; Status = 'present' }
    }
    else {
        if (Install-WithWinget 'OpenJS.NodeJS.LTS' 'Node.js') {
            Refresh-Path
            $post = Resolve-Tool -Command 'node'
            $ver = Get-CommandVersion 'node' @('--version')
            if ($ver) {
                Write-Status "Node.js" "v$ver (installed)" -Color Green
                $script:prereqResults += @{ Name = 'Node.js'; Version = "v$ver"; Status = 'installed' }
            }
            else {
                Write-Status "Node.js" "installed but not in PATH" -Color Yellow
                $script:prereqResults += @{ Name = 'Node.js'; Version = $null; Status = 'installed' }
            }
        }
        else {
            Write-Status "Node.js" "not installed" -Color Red
            $script:prereqResults += @{ Name = 'Node.js'; Version = $null; Status = 'failed' }
        }
    }

    # ── 3. GitHub CLI ─────────────────────────────────────────────────────────
    $res = Resolve-Tool -Command 'gh'
    if ($res.Found -and $res.FixedPath) {
        $ver = Get-CommandVersion 'gh' @('--version')
        Write-Status "GitHub CLI" "v$ver (found, fixed PATH)" -Color Yellow
        $script:prereqResults += @{ Name = 'GitHub CLI'; Version = "v$ver"; Status = 'path-fixed' }
    }
    elseif ($res.Found) {
        $ver = Get-CommandVersion 'gh' @('--version')
        Write-Status "GitHub CLI" "v$ver (already installed)" -Color Green
        $script:prereqResults += @{ Name = 'GitHub CLI'; Version = "v$ver"; Status = 'present' }
    }
    else {
        if (Install-WithWinget 'GitHub.GitHubCLI' 'GitHub CLI') {
            Refresh-Path
            $post = Resolve-Tool -Command 'gh'
            $ver = Get-CommandVersion 'gh' @('--version')
            if ($ver) {
                Write-Status "GitHub CLI" "v$ver (installed)" -Color Green
                $script:prereqResults += @{ Name = 'GitHub CLI'; Version = "v$ver"; Status = 'installed' }
            }
            else {
                Write-Status "GitHub CLI" "installed but not in PATH" -Color Yellow
                $script:prereqResults += @{ Name = 'GitHub CLI'; Version = $null; Status = 'installed' }
            }
        }
        else {
            Write-Status "GitHub CLI" "not installed" -Color Red
            $script:prereqResults += @{ Name = 'GitHub CLI'; Version = $null; Status = 'failed' }
        }
    }

    # ── 4. Claude Code ────────────────────────────────────────────────────────
    $res = Resolve-Tool -Command 'claude'
    if ($res.Found -and $res.FixedPath) {
        $ver = Get-CommandVersion 'claude' @('--version')
        Write-Status "Claude Code" "v$ver (found, fixed PATH)" -Color Yellow
        $script:prereqResults += @{ Name = 'Claude Code'; Version = "v$ver"; Status = 'path-fixed' }
    }
    elseif ($res.Found) {
        $ver = Get-CommandVersion 'claude' @('--version')
        Write-Status "Claude Code" "v$ver (already installed)" -Color Green
        $script:prereqResults += @{ Name = 'Claude Code'; Version = "v$ver"; Status = 'present' }
    }
    else {
        try {
            Write-Step "Installing Claude Code via official installer..."
            $installerScript = Invoke-RestMethod -Uri 'https://claude.ai/install.ps1' -UseBasicParsing
            Invoke-Expression $installerScript
            Refresh-Path
            # The official installer puts claude in ~/.local/bin; resolve again
            $post = Resolve-Tool -Command 'claude'
            if ($post.Found) {
                $ver = Get-CommandVersion 'claude' @('--version')
                $msg = if ($post.FixedPath) { "v$ver (installed, fixed PATH)" } else { "v$ver (installed)" }
                Write-Status "Claude Code" $msg -Color Green
                $script:prereqResults += @{ Name = 'Claude Code'; Version = "v$ver"; Status = 'installed' }
            }
            else {
                Write-Status "Claude Code" "installer ran but command not found" -Color Yellow
                $script:prereqResults += @{ Name = 'Claude Code'; Version = $null; Status = 'failed' }
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Warn "Failed to install Claude Code: $errMsg"
            Write-Status "Claude Code" "not installed" -Color Red
            $script:prereqResults += @{ Name = 'Claude Code'; Version = $null; Status = 'failed' }
        }
    }

    # ── npm global bin PATH fix ───────────────────────────────────────────────
    # Before checking npm-based tools, ensure the npm global bin dir is in PATH.
    if (Test-CommandExists 'npm') {
        $npmBin = Get-NpmGlobalBin
        if ($npmBin -and (Test-Path $npmBin)) {
            $added = Add-ToUserPath -Directory $npmBin
            if ($added) {
                Write-Step "Added npm global bin to PATH: $npmBin"
            }
        }
    }

    # ── npm-based CLI tools (require Node.js) ─────────────────────────────────
    $hasNode = Test-CommandExists 'node'
    if (-not $hasNode) {
        Write-Warn "Node.js is not available - skipping npm-based CLI tools (Copilot, Codex, Gemini)."
        foreach ($name in @('Copilot CLI', 'Codex CLI', 'Gemini CLI')) {
            Write-Status $name "skipped (no Node.js)" -Color DarkGray
            $script:prereqResults += @{ Name = $name; Version = $null; Status = 'skipped' }
        }
    }
    else {
        # ── 5. GitHub Copilot CLI ─────────────────────────────────────────────
        $res = Resolve-Tool -Command 'copilot'
        if ($res.Found -and $res.FixedPath) {
            $ver = Get-CommandVersion 'copilot' @('--version')
            Write-Status "Copilot CLI" "v$ver (found, fixed PATH)" -Color Yellow
            $script:prereqResults += @{ Name = 'Copilot CLI'; Version = "v$ver"; Status = 'path-fixed' }
        }
        elseif ($res.Found) {
            $ver = Get-CommandVersion 'copilot' @('--version')
            Write-Status "Copilot CLI" "v$ver (already installed)" -Color Green
            $script:prereqResults += @{ Name = 'Copilot CLI'; Version = "v$ver"; Status = 'present' }
        }
        else {
            if (Install-WithNpm '@github/copilot' 'Copilot CLI') {
                Refresh-Path
                $post = Resolve-Tool -Command 'copilot'
                $ver = Get-CommandVersion 'copilot' @('--version')
                if ($ver) {
                    Write-Status "Copilot CLI" "v$ver (installed)" -Color Green
                    $script:prereqResults += @{ Name = 'Copilot CLI'; Version = "v$ver"; Status = 'installed' }
                }
                else {
                    Write-Status "Copilot CLI" "installed but not in PATH" -Color Yellow
                    $script:prereqResults += @{ Name = 'Copilot CLI'; Version = $null; Status = 'installed' }
                }
            }
            else {
                Write-Status "Copilot CLI" "not installed" -Color Red
                $script:prereqResults += @{ Name = 'Copilot CLI'; Version = $null; Status = 'failed' }
            }
        }

        # ── 6. Codex CLI ─────────────────────────────────────────────────────
        $res = Resolve-Tool -Command 'codex'
        if ($res.Found -and $res.FixedPath) {
            $ver = Get-CommandVersion 'codex' @('--version')
            Write-Status "Codex CLI" "v$ver (found, fixed PATH)" -Color Yellow
            $script:prereqResults += @{ Name = 'Codex CLI'; Version = "v$ver"; Status = 'path-fixed' }
        }
        elseif ($res.Found) {
            $ver = Get-CommandVersion 'codex' @('--version')
            Write-Status "Codex CLI" "v$ver (already installed)" -Color Green
            $script:prereqResults += @{ Name = 'Codex CLI'; Version = "v$ver"; Status = 'present' }
        }
        else {
            if (Install-WithNpm '@openai/codex' 'Codex CLI') {
                Refresh-Path
                $post = Resolve-Tool -Command 'codex'
                $ver = Get-CommandVersion 'codex' @('--version')
                if ($ver) {
                    Write-Status "Codex CLI" "v$ver (installed)" -Color Green
                    $script:prereqResults += @{ Name = 'Codex CLI'; Version = "v$ver"; Status = 'installed' }
                }
                else {
                    Write-Status "Codex CLI" "installed but not in PATH" -Color Yellow
                    $script:prereqResults += @{ Name = 'Codex CLI'; Version = $null; Status = 'installed' }
                }
            }
            else {
                Write-Status "Codex CLI" "not installed" -Color Red
                $script:prereqResults += @{ Name = 'Codex CLI'; Version = $null; Status = 'failed' }
            }
        }

        # ── 7. Gemini CLI ────────────────────────────────────────────────────
        $res = Resolve-Tool -Command 'gemini'
        if ($res.Found -and $res.FixedPath) {
            $ver = Get-CommandVersion 'gemini' @('--version')
            Write-Status "Gemini CLI" "v$ver (found, fixed PATH)" -Color Yellow
            $script:prereqResults += @{ Name = 'Gemini CLI'; Version = "v$ver"; Status = 'path-fixed' }
        }
        elseif ($res.Found) {
            $ver = Get-CommandVersion 'gemini' @('--version')
            Write-Status "Gemini CLI" "v$ver (already installed)" -Color Green
            $script:prereqResults += @{ Name = 'Gemini CLI'; Version = "v$ver"; Status = 'present' }
        }
        else {
            if (Install-WithNpm '@google/gemini-cli' 'Gemini CLI') {
                Refresh-Path
                $post = Resolve-Tool -Command 'gemini'
                $ver = Get-CommandVersion 'gemini' @('--version')
                if ($ver) {
                    Write-Status "Gemini CLI" "v$ver (installed)" -Color Green
                    $script:prereqResults += @{ Name = 'Gemini CLI'; Version = "v$ver"; Status = 'installed' }
                }
                else {
                    Write-Status "Gemini CLI" "installed but not in PATH" -Color Yellow
                    $script:prereqResults += @{ Name = 'Gemini CLI'; Version = $null; Status = 'installed' }
                }
            }
            else {
                Write-Status "Gemini CLI" "not installed" -Color Red
                $script:prereqResults += @{ Name = 'Gemini CLI'; Version = $null; Status = 'failed' }
            }
        }
    }

    Write-Host ""
}

# ─── Main ─────────────────────────────────────────────────────────────────────

function Install-Bob {
    Show-Banner

    # ── System detection ──────────────────────────────────────────────────────
    if ($env:OS -ne 'Windows_NT' -and -not ($PSVersionTable.Platform -eq 'Win32NT' -or -not $PSVersionTable.Platform)) {
        Write-Fail "Bob installer is only supported on Windows."
        return
    }

    $arch = Get-Architecture
    Write-Status "Detecting system" "Windows $arch" -Color Cyan

    # ── Check existing install ────────────────────────────────────────────────
    $existingPath = Find-ExistingInstall
    $installedVersion = $null

    if ($existingPath) {
        $installedVersion = Get-InstalledVersion -ExePath $existingPath
        if ($installedVersion) {
            Write-Status "Current version" "v$installedVersion" -Color Yellow
        }
        else {
            Write-Status "Current version" "installed (version unknown)" -Color Yellow
        }
    }
    else {
        Write-Status "Current version" "not installed" -Color DarkGray
    }

    # ── Resolve target release ────────────────────────────────────────────────
    Write-Step "Fetching release information..."

    $release = $null
    if ($Version) {
        $release = Get-SpecificRelease -Tag (Normalize-Version $Version)
    }
    else {
        $release = Get-LatestRelease
    }

    $targetVersion = Normalize-Version $release.tag_name
    Write-Status "Target version" "v$targetVersion" -Color Green

    # ── Version comparison ────────────────────────────────────────────────────
    if ($installedVersion -and -not $Force) {
        $cmp = Compare-SemVer -Left $installedVersion -Right $targetVersion
        if ($cmp -ge 0) {
            Write-Host ""
            Write-Host "  Bob v$installedVersion is already " -NoNewline
            if ($cmp -eq 0) {
                Write-Host "up-to-date" -ForegroundColor Green -NoNewline
            }
            else {
                Write-Host "newer than v$targetVersion" -ForegroundColor Green -NoNewline
            }
            Write-Host "."
            Write-Host "  Use " -NoNewline
            Write-Host "-Force" -ForegroundColor Yellow -NoNewline
            Write-Host " to reinstall."
            Write-Host ""
            return
        }
    }

    # ── Find installer asset ──────────────────────────────────────────────────
    $asset = Find-InstallerAsset -Release $release -Arch $arch
    $installerName = $asset.name
    $downloadUrl   = $asset.browser_download_url
    $fileSize      = $asset.size

    Write-Status "Downloading" "$installerName ($(Format-FileSize $fileSize))" -Color White

    # ── Download ──────────────────────────────────────────────────────────────
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "bob-install-$([System.Guid]::NewGuid().ToString('N').Substring(0,8))"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $tempInstaller = Join-Path $tempDir $installerName

    try {
        Invoke-Download -Url $downloadUrl -OutFile $tempInstaller -ExpectedSize $fileSize
    }
    catch {
        Write-Fail "Download failed: $_"
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    if (-not (Test-Path $tempInstaller)) {
        Write-Fail "Download failed: installer file not found."
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    $actualSize = (Get-Item $tempInstaller).Length
    if ($fileSize -gt 0 -and $actualSize -ne $fileSize) {
        Write-Fail "Download incomplete: expected $(Format-FileSize $fileSize), got $(Format-FileSize $actualSize)."
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    # ── Verify checksum ───────────────────────────────────────────────────────
    Write-Step "Verifying integrity..."
    $checksumInfo = Get-ChecksumFromRelease -Release $release -InstallerName $installerName

    if ($checksumInfo) {
        $algo = $checksumInfo.Algorithm
        $expected = $checksumInfo.Hash

        $actual = (Get-FileHash -Path $tempInstaller -Algorithm $algo).Hash.ToUpper()

        if ($actual -ne $expected) {
            Write-Fail "Checksum verification failed!`n    Expected ($algo): $expected`n    Actual   ($algo): $actual`n  The download may be corrupted. Please try again."
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
        Write-Status "Integrity check" "$algo verified" -Color Green
    }
    else {
        Write-Status "Integrity check" "skipped (no checksum available)" -Color DarkGray
    }

    # ── Install ───────────────────────────────────────────────────────────────
    Write-Step "Running installer..."

    try {
        $process = Start-Process -FilePath $tempInstaller -ArgumentList '/S' -PassThru -Wait
        if ($process.ExitCode -ne 0) {
            Write-Fail "Installer exited with code $($process.ExitCode)."
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
    }
    catch {
        Write-Fail "Failed to run installer: $_"
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        return
    }

    Write-Status "Installing" "done" -Color Green

    # ── Prerequisites (CLI agents) ────────────────────────────────────────────
    if (-not $SkipPrereqs) {
        Install-Prerequisites
    }
    else {
        Write-Host ""
        Write-Step "Skipping prerequisites (-SkipPrereqs specified)"
    }

    # ── Verify installation ───────────────────────────────────────────────────
    $finalPath = Find-ExistingInstall
    $finalVersion = if ($finalPath) { Get-InstalledVersion -ExePath $finalPath } else { $null }

    $displayVersion = if ($finalVersion) { "v$finalVersion" } else { "v$targetVersion" }
    $displayPath    = if ($finalPath) { Split-Path $finalPath } else { "$env:LOCALAPPDATA\Programs\Bob" }

    # ── Clean up ──────────────────────────────────────────────────────────────
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

    # ── Success ───────────────────────────────────────────────────────────────
    if ($installedVersion) {
        Write-Success "Bob upgraded from v$installedVersion to $displayVersion!"
    }
    else {
        Write-Success "Bob $displayVersion installed successfully!"
    }

    Write-Host ""
    Write-Host "  Install path:  " -NoNewline
    Write-Host $displayPath -ForegroundColor Cyan

    # Show summary of prerequisites
    if (-not $SkipPrereqs -and $script:prereqResults -and $script:prereqResults.Count -gt 0) {
        Write-Host ""
        Write-Host "  CLI agents:" -ForegroundColor Cyan
        foreach ($r in $script:prereqResults) {
            $icon = switch ($r.Status) {
                'present'    { '[=]'; break }
                'installed'  { '[+]'; break }
                'path-fixed' { '[~]'; break }
                'skipped'    { '[-]'; break }
                default      { '[!]'; break }
            }
            $color = switch ($r.Status) {
                'present'    { 'Green';     break }
                'installed'  { 'Green';     break }
                'path-fixed' { 'Yellow';    break }
                'skipped'    { 'DarkGray';  break }
                default      { 'Yellow';    break }
            }
            $detail = if ($r.Version) { $r.Version } else { $r.Status }
            Write-Host "    $icon " -ForegroundColor $color -NoNewline
            Write-Host "$($r.Name): $detail"
        }
    }

    Write-Host ""
    Write-Host "  Launch Bob from the Start Menu or run: " -NoNewline
    Write-Host "Bob" -ForegroundColor Green
    Write-Host ""
}

# ── Entry point ───────────────────────────────────────────────────────────────

try {
    Install-Bob
}
catch {
    Write-Host ""
    Write-Host "  [ERROR] " -ForegroundColor Red -NoNewline
    Write-Host "An unexpected error occurred:"
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ScriptStackTrace) {
        Write-Host ""
        Write-Host "  Stack trace:" -ForegroundColor DarkGray
        Write-Host "  $($_.ScriptStackTrace)" -ForegroundColor DarkGray
    }
    Write-Host ""
    exit 1
}
