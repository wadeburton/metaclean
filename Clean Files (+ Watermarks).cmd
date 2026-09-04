@echo off
rem Same as "Clean Files.cmd" but also attempts watermark removal:
rem PDF Watermark annotations, watermark-named layers, and the same overlay
rem drawn on every page; docx header watermark shapes (Word's built-in
rem Insert Watermark feature). Best-effort -- see README for what this can
rem and can't remove.
py "%~dp0metaclean.py" auto --watermarks %*
echo.
pause
