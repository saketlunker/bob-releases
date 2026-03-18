# Bob Releases

> Installer binaries for [Bob](https://github.com/saketlunker/bob) — the agent orchestrator.

This repo contains **only compiled binaries**. No source code.

## Install Bob

```powershell
irm https://raw.githubusercontent.com/saketlunker/bob-releases/main/install.ps1 | iex
```

## Update

Bob auto-updates on launch. No action needed.

To manually update: re-run the install command above.

## Install specific version

```powershell
& { param($Version) irm https://raw.githubusercontent.com/saketlunker/bob-releases/main/install.ps1 | iex } -Version "1.2.0"
```
