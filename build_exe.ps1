$ErrorActionPreference = "Stop"

$python = "python"
& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
& $python -m pip install pyinstaller

& $python -m PyInstaller --noconfirm --clean --noconsole --onefile --name "LMReview" --collect-all customtkinter --collect-all watchdog --collect-all docx notebooklm_single_folder_flow.py

Write-Host "Build complete: dist\\LMReview.exe"
