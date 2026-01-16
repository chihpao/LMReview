AGENTS.md

Purpose
- Keep this repository maintainable and safe for internal use.
- Prioritize clarity, stability, and repeatability over cleverness.

Project Overview
- Desktop GUI tool for document review workflow with CustomTkinter.
- Key file: notebooklm_single_folder_flow.py
- Build scripts: build_exe.ps1, build_exe.bat

Working Rules
- Preserve existing behavior unless explicitly asked to change it.
- Avoid adding dependencies unless required.
- Keep edits minimal and localized; prefer targeted changes.
- Use ASCII in files that are ASCII-only.
- Do not commit generated artifacts (logs, output docs, __pycache__).

UI Changes
- Maintain a clean, high-contrast light theme.
- Keep primary flow visible without scrolling when possible.
- Align controls and spacing; avoid overcrowding.

Testing
- At minimum: python -m py_compile notebooklm_single_folder_flow.py
- If GUI changes are significant, manually verify layout and flow.

Git Hygiene
- Keep .gitignore updated for logs, outputs, caches.
- Commit messages should be short and descriptive.
