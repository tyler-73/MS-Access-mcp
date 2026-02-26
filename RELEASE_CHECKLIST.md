# Release Checklist

## 1. Repair + Verify Hardening (recommended pre-publish)

Run from repo root:

```powershell
# Full hardening run
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateConfigs

# Dry-run / WhatIf
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateConfigs `
  -WhatIf

# Config update toggle: Codex only
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateCodexConfig

# Config update toggle: Claude only
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateClaudeConfig

# Optional: override Claude Desktop config file path
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateClaudeConfig `
  -ClaudeConfigPath "C:\path\to\claude_desktop_config.json"

# Default: require validated x64 promoted binary and manifest git_commit matching HEAD
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb"

# Optional: diagnostics override for manifest/HEAD mismatch checks
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -AllowManifestHeadMismatch

# Strict mode: require regression-backed validation manifest for candidate selection
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -RequireRegressionBackedManifest

# Strict-mode override: keep strict mode visible but allow non-regression manifests
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -RequireRegressionBackedManifest `
  -AllowNonRegressionManifest

# Optional: diagnostics override for unvalidated binaries
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -AllowUnvalidatedBinary

# Optional: include x86 fallback candidate
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -AllowX86Fallback
```

## 2. Publish + Promote x64 (default smoke verification)

Run from repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1
```

Expected result:
- Promotion completes to `.\mcp-server-official-x64`
- Output contains `Smoke test passed.`
- Output contains `Validation manifest written: ...\mcp-server-official-x64\release-validation.json`
- Validation manifest contains `git_commit`, `regression_run`, and `regression_passed`
- A backup directory is preserved as `.\mcp-server-official-x64-backup-*` when a prior target existed
- If promotion fails after backup creation, the script attempts rollback to the previous target
- If `-RunRegression` is used and regression fails after promotion, the script archives the promoted target and restores the backup target when available

Optional backup pruning:

```powershell
# Keep newest 5 backups
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 -BackupRetentionCount 5
```

Optional artifact cleanup verification:

```powershell
# Preview stale run/smoke/backup cleanup behavior
powershell -ExecutionPolicy Bypass -File .\scripts\prune-release-artifacts.ps1 -WhatIf
```

## 3. Repo hygiene gate (required before release)

Run from repo root:

```powershell
# Prune stale release artifacts (run/smoke, plus backups when requested)
powershell -ExecutionPolicy Bypass -File .\scripts\prune-release-artifacts.ps1 -IncludeBackups -BackupRetentionCount 5

# Tree must be clean except known, intentional source edits for this release
git status --short

# Verify release scripts are versioned
git ls-files scripts
```

Pass criteria:
- No untracked release artifact directories remain (for example `mcp-server-official-x64-run-*`, `mcp-server-official-x64-smoke*`, `mcp-server-official-x64-next*`, `mcp-server-official-x64-backup-*`, `mcp-server-official-x64-regression-failed-*`)
- Output is clean or shows only known, intentional source file changes for the release
- `git ls-files scripts` returns the expected release scripts (`publish-and-promote-x64.ps1`, `repair-and-verify-access-mcp.ps1`, `prune-release-artifacts.ps1`, `bootstrap-github-actions.ps1`)

## 4. Full regression gate

Use the committed harness (recommended before release sign-off):

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1 `
  -ServerExe ".\mcp-server-official-x64\MS.Access.MCP.Official.exe" `
  -DatabasePath "C:\path\to\database.accdb"

powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_negative_paths.ps1 `
  -ServerExe ".\mcp-server-official-x64\MS.Access.MCP.Official.exe" `
  -DatabasePath "C:\path\to\database.accdb"
```

Note:
- Full regression runs headless by default to avoid interactive Access pop-ups.
- Use `-IncludeUiCoverage` only for explicit UI-tool validation windows.

Pass criteria:
- Script exits with code `0`
- Output contains `TOTAL_FAIL=0`
- Output contains `NEGATIVE_PATHS_PASS=1`
- Full regression output includes `database_file_tools_coverage: INFO` (or explicit `SKIP` only when `-AllowCoverageSkips` is intentionally used)
- Negative-path output includes `connect_access_secure_arg_detected=`

## 5. GitHub workflow bootstrap (self-hosted CI)

Run from repo root when onboarding a new machine/account:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bootstrap-github-actions.ps1 `
  -SetDatabaseSecret `
  -DatabasePath "C:\path\to\database.accdb" `
  -TriggerRegressionWorkflow
```

Expected result:
- `gh auth status` succeeds
- Repo access check succeeds
- `ACCESS_DATABASE_PATH` secret is set
- `windows-self-hosted-access-regression.yml` dispatch is accepted

## If gate fails

1. Kill stale processes:
   - `MSACCESS`
   - `MS.Access.MCP.Official`
2. Remove stale lock file (`.laccdb`) next to your test database.
3. Re-run the regression command.
4. If publish/promotion reports `Access denied`, rerun promotion from an elevated PowerShell shell so locked server processes can be terminated.
