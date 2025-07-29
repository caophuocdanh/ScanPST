rmdir /S /Q dist
rmdir /S /Q build
pip install ttkbootstrap; pywinauto; pyperclip
pyinstaller --windowed --onefile --icon=icon.ico --name ScanPST "scanpst.py"