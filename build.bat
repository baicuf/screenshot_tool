@echo off
pip install -r requirements.txt
pyinstaller --onefile --windowed screenshot_gui.py
pause
