@echo off
rem BCS data story runner: dev activate (if needed) + creds + venv + pipeline.
cd /d %~dp0

rem -- Make python available (VDI: python only exists after `dev activate`) --
where python >nul 2>nul
if errorlevel 1 (
    echo [info] python not on PATH - running "dev activate" ...
    call dev activate
)
where python >nul 2>nul
if errorlevel 1 (
    echo [error] python still not available after "dev activate".
    echo         Run "dev activate" manually in this window, then rerun run.cmd.
    exit /b 1
)

rem -- Session credentials -------------------------------------------------
if exist creds.local.cmd (
    call creds.local.cmd
) else (
    echo [warn] creds.local.cmd not found - relying on AWS keys already set in this window.
)

rem -- Virtual environment -------------------------------------------------
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
) else (
    echo [warn] .venv not found - run setup.cmd once first. Using active python.
)

python run_all.py %*
