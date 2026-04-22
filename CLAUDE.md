# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Windows desktop automation script for dental clinic insurance patient lookup. Queries MSSQL for eligible patients, then automates GUI interaction with **덴트웹** (DentWeb) dental practice management software to perform insurance eligibility checks for each patient.

## Running the Script

```bash
python test3_reconstructed.py
```

Requires Windows — uses `pywinauto` and `pygetwindow` for Win32 GUI automation.

## Dependencies

- `pygetwindow`, `pywinauto`, `pyautogui` — Windows GUI automation
- `pyodbc` — MSSQL connection (`127.0.0.1,1436`, database `DentWeb`)
- `pandas`, `openpyxl` — Excel export
- `tkinter` — completion popup

## Architecture

Single-file script with no classes. Flow:

1. **`get_patient_list()`** — Queries `TB_환자정보` for patients 65+ with valid 주민번호, excluding those already in `TB_보험임플환자`, sorted by name.
2. **`find_and_maximize_dentweb()`** — Finds the running 덴트웹 window (title contains `▶ 덴트웹`) and maximizes it via `pywinauto`.
3. **Per-patient loop** — For each patient: types `이름(차트번호)` into the last Edit field, presses Enter, clicks `보험자격조회`, handles pop-up dialogs.
4. **Dialog handlers** — `check_for_verification_window()` dismisses `본인확인 방법 선택` with `취소`; `check_insurance_result_window()` handles `진료 의뢰기관 기호 입력` and `조회 결과` dialogs.
5. **`save_to_excel()`** — Re-queries the DB and saves results to a timestamped `.xlsx` file.

## Key Constraints

- Script must run on the same machine as 덴트웹 (GUI automation via Win32).
- DB credentials are hardcoded (`sa` / `Q3xzJiwpv2zC`) — do not commit changes that expose these further.
- `find_EditText()` selects the **last** Edit control in the window hierarchy — fragile if the DentWeb UI layout changes.
- `time.sleep()` calls are load-bearing; removing them breaks UI interaction timing.
