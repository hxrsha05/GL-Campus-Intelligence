@echo off
schtasks /create /tn "GL Campus - Daily Dashboard Email" /tr "\"C:\Users\Harshavardhan J\AppData\Local\Programs\Python\Python310\python.exe\" \"E:\GLIM\send_report.py\"" /sc daily /st 08:00 /f /rl highest
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: Daily 8 AM email task created.
) else (
    echo FAILED. Run this batch file as Administrator.
)
pause
