@echo off
rem Drag files (PDF, docx/xlsx/pptx, txt/md) onto this icon and drop them.
rem Or: drop files into the "inbox" folder next to this file, then just
rem double-click this icon with nothing selected.
py "%~dp0metaclean.py" auto %*
echo.
pause
