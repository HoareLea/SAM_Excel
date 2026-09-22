# Project Progress

## Branch
`sow/2026-Q3`

## Last updated
2026-09-22 - app.config cleanup merged

## Current status
Part of the repo-family .NET Framework `app.config` cleanup: base [SAM#126](https://github.com/SAM-BIM/SAM/pull/126) plus 17 sibling PRs, all merged into `sow/2026-Q3` on 2026-09-22 (SAM first), with their branches deleted.

## Completed
- [SAM_Excel#5](https://github.com/SAM-BIM/SAM_Excel/pull/5) merged as `4960137c`: removed dead .NET Framework `app.config` files.

## Decisions / assumptions
- Every project here targets `netstandard2.0` or `net8.0(-windows)` and is an `OutputType Library`. Library `.dll.config` files are never read at runtime (only the host `Rhino.exe`/`Revit.exe` config is), so the net472-era binding redirects, `<supportedRuntime>` and `loadFromRemoteSources` were inert. They only emitted stale `.dll.config` files into `build/` and `%APPDATA%\SAM`.
- No `ConfigurationManager`/`AppSettings` use in the repo; deleted files held binding/runtime config only.

## Files changed
- `Grasshopper/SAM.Core.Grasshopper.Excel/app.config` (deleted)

## Validation
- Before merge, full `BuildAlls_v4.bat` (Debug Restore;Clean;Rebuild of every repo, starting from an emptied `%APPDATA%\SAM`): exit 0, 0 errors. The redeployed `%APPDATA%\SAM` has no `SAM.*.dll.config`.
- CI on the PR: build + spdx pass.

## Issues / blockers
- None known for this cleanup.

## Next step
- None for this cleanup. Continue with the next planned task on `sow/2026-Q3`.
