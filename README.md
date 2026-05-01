# MS Project Field Mapper

A VBA tool, hosted in Microsoft Project Desktop, that maps custom fields
across multiple `.mpp` files and extracts harmonized data for downstream use
in Power BI and other reporting tools.

## What it does (when complete)

- Sequentially scans up to 10 source `.mpp` files (open -> cache -> close).
- Auto-maps custom fields against a chosen master using a confidence-scored
  algorithm; only ambiguous cases prompt the user via a UserForm.
- Extracts harmonized data per-project and combined to CSV + multi-sheet
  XLSX (late-bound Excel) for Power BI ingestion.
- Optionally builds an integrated Master Schedule via inserted subprojects
  (originals untouched).
- Persists mappings as `.fmap` files so a project pair can be re-run without
  re-mapping.

## Current status: Phases 0-1

This branch ships the foundation plus the single-file scanner:

- **Environment detection** (`modEnvironment`) - host check, version check,
  AutomationSecurity, Trusted Locations, MOTW probe.
- **Logging infrastructure** (`clsLogger` + `clsLogEntry`).
- **Run-context object** (`clsRunContext`) - replaces global state.
- **Constants & enums** (`modConstants`).
- **File fingerprint hashing** (`modHash`) - FNV-1a 32-bit.
- **Field metadata + project scan classes** (`clsFieldMetadata`,
  `clsProjectScan`).
- **File scanner** (`modFileScanner`) - open/scan/close one `.mpp`,
  enumerate Task and Resource custom fields, sample values, detect
  duplicate aliases.
- **Phase 0 UI** (`frmMain`) - **Run Diagnostics** button.
- **Progress UI** (`frmProgress`) - subscribed-to-logger live log + Cancel
  button. Wired up for use by the Phase 2 multi-file orchestrator.

The full pipeline (multi-file orchestration, auto-mapper, mapping UI,
extraction, merge, persistence) lands in subsequent phases.

## Roadmap

| Phase | Goal | Status |
|---|---|---|
| 0 | Foundation: environment detection + diagnostics UI | **shipped** |
| 1 | File scanner: open/scan/close one `.mpp` | **shipped** |
| 2 | Multi-file orchestration with cancel + error recovery | next |
| 3 | Auto-mapper: confidence scoring + threshold partitioning | planned |
| 4 | Manual mapping UI (3-listbox `frmMapping`) | planned |
| 5 | Collision handling | planned |
| 6 | Extraction: CSV + XLSX | planned |
| 7 | Merge mode: subproject-based master schedule | planned |
| 8 | Persistence (`.fmap` save/load) | planned |
| 9 | Packaging & distribution (signed `.mpt` global template) | planned |

The full plan, including the auto-mapper scoring rules, persistence schema,
collision strategy, and risk register, lives in the planning doc that
generated this branch (saved by the planning workflow).

## Requirements

- Microsoft Project Desktop 2016 or later.
- Microsoft Scripting Runtime reference (added during install).

**Not supported:** Project for the Web, Project Online, or any browser-based
Project. Those hosts do not run VBA.

## Install

See [INSTALL.md](INSTALL.md) for step-by-step VBE import instructions and a
fallback for manual form construction.

## Repository layout

```
src/                          VBA source modules (.bas, .cls, .frm)
  modConstants.bas            enums, slot tables, threshold constants
  modHash.bas                 FNV-1a 32-bit + file fingerprint
  modEnvironment.bas          host/version/security detection
  modFileScanner.bas          open/scan/close one .mpp; enumerate fields
  modMain.bas                 entry points (ShowMain, RunDiagnostics, Phase1_TestScan)
  clsLogEntry.cls             single log record
  clsLogger.cls               buffered session logger
  clsFieldMetadata.cls        one custom-field record (alias, type, samples)
  clsProjectScan.cls          per-.mpp scan results
  clsRunContext.cls           per-session state container
  frmMain.frm                 Phase 0 UI (Diagnostics only)
  frmMain.code.bas            code-only fallback
  frmProgress.frm             live progress + log UI; Cancel button
  frmProgress.code.bas        code-only fallback
samples/                      sample .mpp files (added in later phases)
dist/                         built bundle / signed template (added in Phase 9)
docs/                         supplementary documentation (added as needed)
INSTALL.md                    install instructions
README.md                     this file
LICENSE                       project licence
```

## Licence

MIT - see [LICENSE](LICENSE).
