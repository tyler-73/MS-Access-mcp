# Release Checklist

## Access MCP release gate

Run this command from repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1
```

## Pass criteria

- Script exits with code `0`
- Output contains `TOTAL_FAIL=0`

## If gate fails

1. Kill stale processes:
   - `MSACCESS`
   - `MS.Access.MCP.Official`
2. Remove stale lock file (`.laccdb`) next to your test database.
3. Re-run the regression command.
