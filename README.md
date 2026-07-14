# VBA-KPI
# Stope Design KPI Tracker (VBA)

VBA port of an Office Script that tracks mining **stopes** as they move through a
design workflow. It builds a KPI table from source data, then on each subsequent
run grades every stope as a **progression**, **non-progression**, or unchanged by
comparing its current design stage against the previous run's saved baseline.

## Modules

| File | Contains | Purpose |
|------|----------|---------|
| `KPI_Common.bas` | Config constants + all shared helpers | Lookup formulas, grading logic, cache read/write, summary/breakdown output |
| `KPI_Build.bas` | `BuildKPITable` macro | Creates/refreshes the KPI table from the source sheet |
| `KPI_StatusCheck.bas` | `RunStatusCheck` macro | Grades each stope vs. the saved baseline and writes results |

## Sheets

- **`Stope Cadence`** (source) — raw stope data. Headers in row 5, data below.
- **`SchedulerData`** (target) — holds the `KPI` table plus the run summary
  (`Total Stopes` / `BLACK` / `RED` counts + last-updated time) at `A1`.
- **`_StageStateCache`** (hidden) — persistence between runs:
  - `A:C` — per-stope baseline: StopeID / DesignStage / SubProcess
  - `E` — the ordered list of stage keys (`StageOrder`)
  - `G:I` — the per-engineer Progression / Non-Progression breakdown

The cache sheet is `xlSheetHidden` (hidden, but users can unhide via right-click),
not very-hidden.

## Config (top of `KPI_Common.bas`)

| Constant | Value | Meaning |
|----------|-------|---------|
| `SRC_SHEET` | `Stope Cadence` | Source data sheet |
| `TGT_SHEET` | `SchedulerData` | KPI table + summary destination |
| `STATE_SHEET` | `_StageStateCache` | Hidden persistence sheet |
| `TBL_NAME` | `KPI` | Output ListObject name |
| `COL_ID` | 3 | Stope ID column in source |
| `COL_USER` | 8 | Assigned engineer |
| `COL_STAGE` | 22 | Design stage |
| `COL_SUB` | 23 | Sub-process |
| `COL_COMMENTS` | 24 | Comments |
| `COL_ZONE` | 35 | Zone (RED / BLACK / GREEN / YELLOW) |
| `STEPS_RED` | 1 | Stage steps required to count a RED stope as a progression |
| `STEPS_BLACK` | 2 | Stage steps required for a BLACK stope |

## How grading works

Each stope's position is a **stage key**: `CleanStr(stage) & "::" & CleanStr(sub)`
(e.g. `Draft_Design::25%`, `IFR::`). Keys are ranked by their index in the
`StageOrder` list on the cache sheet.

On a run, for each stope:

- **`new`** — the stope had no saved baseline (first time seen).
- **`?`** — the current or previous key isn't in `StageOrder` (unrecognised stage).
- **`Y`** (progression) — the stage advanced by at least the zone threshold
  (`STEPS_RED`=1 for RED, `STEPS_BLACK`=2 for BLACK).
- **`N`** (non-progression) — seen before but did not advance enough.

GREEN / YELLOW stopes are filtered out of the build and skipped during grading.

## Usage

1. Enable **Trust access to the VBA project object model** (needed only for
   re-importing modules programmatically).
2. Import the three `.bas` files into the workbook's VBA project.
3. Run **`BuildKPITable`** once to create the `KPI` table on `SchedulerData`.
4. Run **`RunStatusCheck`** to grade. The first run baselines every stope as
   `new`; each later run grades against the previous baseline and refreshes the
   summary and per-engineer breakdown.

Each macro ends with a `MsgBox` summary — click **OK** to finish.

## Notes / gotchas

- **Blank source cell → `""`, not `0`.** `SetLookup` wraps `INDEX` in an `IF` so a
  blank sub-process yields a key like `IFR::` instead of `IFR::0`.
- **Rebuild shows all rows first.** `ApplyLookupFormulas` calls
  `AutoFilter.ShowAllData` before writing formulas, so hidden GREEN/YELLOW rows
  don't keep stale formulas.
- **Cache baseline is stored as text.** `WriteSavedState` sets `NumberFormat = "@"`
  on `A:C` so sub-process labels like `25%` / `0%` aren't coerced into numbers
  (which previously broke the stage-key match and produced spurious `?` grades).
