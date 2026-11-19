@echo off
rem Zeigt, was in der Excel-Datei fuer heute steht (Name, Personalnummer, Projektzeilen). Aendert nichts.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Stundennachweis.ps1" -Aktion Test
pause
