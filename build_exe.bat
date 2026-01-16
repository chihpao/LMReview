@echo off
setlocal
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --noconsole --onefile --name "NotebookLM_Review_Assistant" --collect-all customtkinter --collect-all watchdog --collect-all docx notebooklm_single_folder_flow.py
echo Build complete: dist\NotebookLM_Review_Assistant.exe
endlocal
