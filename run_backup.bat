@echo off
cd /d "%~dp0"
echo Backing up the SharePoint repositories (Admin, LQAS, IM)...
python sp_versions.py backup
echo.
python sp_versions.py list
echo.
pause
