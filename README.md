# Bell Boy BB-2 — Release1_Free (Sealed)

**Sealed (UTC):** `20260117_004752_347Z`  
**Build file:** `BellBoy-BB2-Release1_Free.ps1`  
**SHA256:** `030d8bfabb97fe0d33696ead39ea70d02b2f50f5d9933f1dcfbdf643a041670d`  
**Bytes:** `110746`

## What this is
This is a **sealed release pack** for an older Bell Boy BB-2 build.  
It contains the exact script, a machine-readable manifest, and integrity hashes.

## Requirements
- Windows 10/11
- Windows PowerShell **5.1**

## Run
From this folder:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\src\BellBoy-BB2-Release1_Free.ps1"
```

## Verify integrity (recommended)
```powershell
Get-FileHash ".\src\BellBoy-BB2-Release1_Free.ps1" -Algorithm SHA256
# Compare to docs\SHA256SUMS.txt
```

## Canon notes (design intent)
- Backups doctrine: every write creates a new timestamped backup; never prune
- Atomic writes: temp → replace/move
- Deletes go to Recycle Bin (no hidden undo)
- Deny-list / permission preflight behavior for protected system paths

**Detected Tab references (best-effort):** `1, 2, 3`

Bell Boy v2 — BB-2 (Release 1 — Free) — Sealed

Product: Bell Boy v2 — BB-2
Edition: Free (Tabs 1–3)
Platform: Windows 10/11
Shell: Windows PowerShell 5.1
Publisher: Trishula Software

Disclaimer

This release is provided as an early introduction to a Trishula Software project that is actively in development and subject to change.

We’re releasing this build for free to introduce Bell Boy to the public and to support a simple idea: helping people work faster and more confidently, without fear of losing work and without losing accountability for what changed.

Bell Boy began life under the internal name “Test-File-Deploy (TFD)”. The original motivation was practical: make it easy to drop and update large blocks of code/text without constantly opening new files, and without creating messy, repetitive manual backups that balloon over time. This release reflects that same intent: safer edits, visible previews, and an operator-friendly workflow.

Performance note: writing speed and overall responsiveness depend on your machine, storage, antivirus settings, and the files you are editing. Results will vary by environment.

What this is

This repository contains a sealed release pack for Bell Boy v2 — BB-2 (Free).

It includes:

the exact script used to run the app (src/)

a machine-readable manifest (docs/manifest.json)

integrity hashes (docs/SHA256SUMS.txt)

a verifier (VERIFY.ps1)

Why it’s useful

Bell Boy is built to help you create, append, and inject code/text safely with:

preview-first workflows (generate preview + diff before writing)

dry-run modes (no writes)

backup-first behavior (writes are intended to be recoverable)

quick “open/reveal” actions for operator speed

Features (this Release 1 Free build)
Tab 1 — Create

Choose a Save Folder (with Browse)

Set File Name (includes filename safety handling)

Templates included:

Blank (UTF-8)

README (txt)

PowerShell Script (.ps1)

JSON stub (.json)

Create / Overwrite (w/ Backup)

Optional scaffold assist:

“Scaffold default structure” toggle

“Scaffold Project Folders” button

Quick actions:

Reveal in Explorer

Open in Notepad

Open in VS Code

Status + “Last actions” log panel (operator visibility)

Tab 2 — Append / Inject

Target file selection (with Browse)

Insert modes:

Append to end

Insert at line (line assist + tip)

Original Preview

Generate Preview + Diff

Final Preview

Unified diff viewer (added/removed)

Dry-Run (no writes) option

Apply action: Append / Inject (Backups Enabled)

Quick actions:

Open in Notepad

Reveal in Explorer

Undo Last Write

Tab 3 — Code Inject

Target file selection (with Browse)

Injection modes:

Append to end

After Pattern

Before Pattern

Replace Pattern

Inside Region (Begin/End)

Pattern controls:

Regex toggle

Case-insensitive toggle

Original Preview

Generate Preview + Diff

Final Preview

Unified diff viewer (added/removed)

Dry-Run (no writes) option

Apply action: Inject (Backups Enabled)

Quick actions:

Reveal in Explorer

Open in Notepad

Open in VS Code

Undo Last Write

Run (GUI)

From repo root:

powershell -NoProfile -ExecutionPolicy Bypass -File ".\src\BellBoy-BB2-Release1_Free.ps1"
Verify integrity (recommended)
powershell -NoProfile -ExecutionPolicy Bypass -File .\VERIFY.ps1

Expected: ALL OK

Headless mode (CLI / automation)

This build includes a headless mode for injection-style operations:

powershell -NoProfile -ExecutionPolicy Bypass -File ".\src\BellBoy-BB2-Release1_Free.ps1" `
  -Headless `
  -Target "C:\Path\to\file.txt" `
  -Mode AfterPattern `
  -Pattern "needle" `
  -Snippet "# injected by BB-2" `
  -DryRun

Notes:

Use either -SnippetPath or -Snippet

Options include: -Encoding (UTF8|UTF8BOM|UTF16LE|ASCII), -EOL (CRLF|LF), -CaseInsensitive, -Literal, -DryRun

Repo layout

src/ — the Bell Boy build script

docs/manifest.json — metadata for this sealed release

docs/SHA256SUMS.txt — integrity hashes

VERIFY.ps1 — integrity verifier

License

TBD (private for now). Do not redistribute without permission.

“CI PR check”
