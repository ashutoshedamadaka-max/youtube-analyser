# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Architecture: WAT Framework

This project uses the **WAT framework** (Workflows, Agents, Tools), which separates AI reasoning from deterministic execution:

- **`workflows/`** — Markdown SOPs that define objectives, required inputs, which tools to use, expected outputs, and edge case handling. These are the instructions you follow. Do not create or overwrite workflows without explicit user approval.
- **`tools/`** — Python scripts that do the actual work (API calls, data transformations, file operations). These are deterministic and testable. Always check here before writing new code — reuse existing scripts when possible.
- **`.tmp/`** — Temporary processing files (scraped data, intermediate exports). Fully disposable and regenerated as needed; never store deliverables here.
- **`.env`** — All API keys and credentials live here exclusively.
- **`credentials.json`, `token.json`** — Google OAuth files (gitignored).

Final outputs (deliverables) go to cloud services (Google Sheets, Slides, etc.), not local files.

## How to Operate

1. **Start with the relevant workflow** in `workflows/` before taking any action. The workflow defines the correct tool sequence.
2. **Run tools, don't do it yourself** — if a task requires hitting an API or transforming data, find the right script in `tools/` and execute it.
3. **When a tool fails**: read the full error, fix the script, retest (check with user before re-running if it consumes paid API credits), then update the workflow with what you learned (rate limits, quirks, better endpoints).
4. **Keep workflows current** — update them when you discover better methods or constraints.

## Self-Improvement Loop

Every failure should result in: fix the tool → verify the fix → update the workflow → move on with a stronger system.
