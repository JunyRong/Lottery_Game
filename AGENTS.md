# AGENTS.md

## Project

Windows desktop lottery app: Python 3.10+ / Tkinter, Excel import/export via openpyxl. No server, build step, or packaging config. `python app.py` starts the GUI. All UI text is Chinese.

## Commands

Run from the repo root (PowerShell):

```powershell
python -m pip install -r requirements.txt
python app.py                                  # or double-click run.bat
python -m unittest discover -s tests -v        # all tests (stdlib unittest, NOT pytest)
python -m unittest tests.test_engine.TestLotteryEngine.test_tier_capacity_enforced -v
python scripts/smoke_ui.py                     # headless UI smoke test (batch draw, small-window layout, manual dialog)
python scripts/make_sample.py                  # regenerate sample/participants.xlsx and sample/prizes.xlsx
python -m PyInstaller --noconfirm --clean --onefile --windowed --name 幸运抽奖 app.py   # build dist/幸运抽奖.exe
```

Set `$env:PYTHONUTF8=1` if Chinese console output is garbled. No linter, formatter, or typechecker is configured.

## Structure

- `app.py` → `lottery.ui.run()`; the only entrypoint.
- `lottery/engine.py` — UI-free draw logic: unique winners across all tiers, per-tier capacity, `draw()` / `draw_all()`. Put rule changes here and add tests.
- `lottery/excel_io.py` — openpyxl reading, header/column sniffing (`inspect_sheet`), result export.
- `lottery/tier_io.py` — prize-tier Excel import (`read_tiers`), header scoring and column hints.
- `lottery/names.py` — parses manual text input; newline/comma/semicolon/顿号 separated, preserves inner spaces and duplicates.
- `lottery/ui.py` — every Tkinter widget/state lives here; `LotteryApp` owns `tiers`, `names`, `engine`.
- `lottery/config.py` — persists tiers to `config/tiers.json`; missing/corrupt file falls back to `default_tiers()`.
- `lottery/models.py` — `PrizeTier`, `Entrant`, `Winner` dataclasses.

## Invariants and quirks

- `Entrant.id` is the 0-based list index (Excel rows or manual entries); duplicate names are intentional distinct people. Never dedupe names.
- `Entrant.number` is the user-facing 1-based 编号 assigned in `LotteryApp._rebuild_engine()`; `Winner.number` mirrors it. Identity must key off `Entrant.id`, not the display number. Show the number alongside names in stage/results/export.
- Excel import and manual entry both go through `LotteryApp._apply_names()`; keep list replacement/reset logic there instead of duplicating it.
- `ManualInputDialog` is preloaded with the current names and returns the full edited list; `_on_manual_input()` skips resetting when the list is unchanged. It accepts both one-by-one additions and bulk paste (via `parse_manual_names`).
- `LotteryApp._rebuild_engine()` discards all winners. Tier edits and tier import are blocked while winners exist (UI requires reset); keep this if adding tier actions.
- Tier list order is a UI convention: high prize first, low prize last. Default selection is the last tier; `_advance_after_draw()` moves toward index 0. Keep `_apply_names`, reset, delete, and tier import consistent with this.
- Draw modes are `one` / `all` (`LotteryApp.draw_mode`). `draw_all()` fills remaining tier slots and stops early when the entrant pool is exhausted; the stage should report 名单不足 in that case.
- `_result_visible` keeps the last draw result on the stage; `_update_draw_state()` must not overwrite `stage_sub` while it is true. `_refresh_all()` clears it.
- Single and batch draws share the same stage format: each winner is shown as `编号 号 · 姓名`. Batch shows at most 3 winners plus `等 N 人`, one per line; the winners table remains the full record.
- Window minsize is 960x640. The right stage panel uses grid with the stage row weighted and the controls row unweighted, so shrinking the window compresses the stage, never the draw button. The winners card is packed with `side=BOTTOM` before the body for the same reason.
- `_fit_stage_font()` shrinks the stage font to fit the current stage size (including multi-line batch results) and hides `stage_sub` only when the stage is too short.
- `lottery/config.py` resolves `config/tiers.json` next to the executable when frozen (`sys.frozen`), otherwise under the repo root. Keep it writable; PyInstaller onefile would otherwise persist settings only in the temp extraction dir.
- `read_tiers()` replaces the whole tier list; it never merges. Header detection picks the row matching the most tier-header hint categories, so title rows above the real header are skipped.
- There are no automated GUI tests. After UI changes, run `python scripts/smoke_ui.py`; for visual/layout changes also launch `python app.py`.
- Excel support is `.xlsx`/`.xlsm` only; openpyxl cannot read legacy `.xls`.
- Header detection scans the first 10 rows for `NAME_HEADER_HINTS`; if none is found every row is treated as data. Column indices in `ColumnOption` are 0-based sheet columns.
- Exported workbooks intentionally contain two sheets: `中奖名单`（编号 first）and `奖品统计`.
- `config/tiers.json` is local, gitignored state; defaults are code in `lottery/config.py`, not JSON.
- Modal dialogs (`ImportDialog`, `PrizeImportDialog`, `ManualInputDialog`, `TierDialog`) rely on `grab_set()` + caller `wait_window()`; follow that pattern for new dialogs.
- UI palette/fonts/geometry constants are at the top of `lottery/ui.py`; reuse them instead of hardcoding colors.
- Stage palette: tech-blue `STAGE_BG`, cyan idle title (`STAGE_IDLE`), red draw results (`RESULT_RED`) and soft-red result subtitles (`RESULT_RED_SOFT`); `_set_stage_result()` switches between hint and result colors. Tree selection uses `#1d4ed8` background with white text so the current tier stays readable when the tree loses focus.
