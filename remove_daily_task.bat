@echo off
schtasks /delete /tn "GL Campus - Daily Dashboard Email" /f
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: Daily email task removed.
) else (
    echo FAILED. Try running as Administrator.
)
pause
