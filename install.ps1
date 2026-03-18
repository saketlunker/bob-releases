<#
.SYNOPSIS
    Installs or updates Bob — the agent orchestrator.

.DESCRIPTION
    Downloads and installs the latest (or specified) version of Bob from GitHub releases.
    Supports silent installation, architecture detection, version comparison, and checksum verification.

.PARAMETER Version
    Install a specific version (e.g., "1.2.0" or "v1.2.0"). Defaults to latest.

.PARAMETER Force
    Reinstall even if the current version is already up-to-date.

.EXAMPLE
    # Install latest version:
    irm https://raw.githubusercontent.com/saketlunker/bob-releases/main/install.ps1 | iex

    # Install a specific version:
    & { param($Version) irm https://raw.githubusercontent.com/saketlunker/bob-releases/main/install.ps1 | iex } -Version "1.2.0"
#>
param(
    [string]$Version,
    [switch]$Force
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
    Write-Host "The agent orchestrator" -ForegroundColor DarkGray -NoNewline
    Write-Host "       |" -ForegroundColor Cyan
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
