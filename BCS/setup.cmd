@echo off
rem One-time setup on the VDI: dev activate + uv venv + dependencies + creds file.
cd /d %~dp0

rem -- Make python/uv available (VDI: only exist after `dev activate`) -----
where python >nul 2>nul
if errorlevel 1 (
    echo [info] python not on PATH - running "dev activate" ...
    call dev activate
)
where python >nul 2>nul
if errorlevel 1 (
    echo [error] python still not available after "dev activate".
    echo         Run "dev activate" manually in this window, then rerun setup.cmd.
    exit /b 1
)

rem -- Create the virtual environment (uv preferred, stdlib venv fallback) --
if not exist .venv (
    where uv >nul 2>nul
    if errorlevel 1 (
        echo [info] uv not found - using python -m venv
        python -m venv .venv
    ) else (
        uv venv .venv
    )
)
call .venv\Scripts\activate.bat

rem -- Install dependencies ------------------------------------------------
where uv >nul 2>nul
if errorlevel 1 (
    pip install -r requirements.txt
) else (
    uv pip install -r requirements.txt
)

rem -- Credentials file ----------------------------------------------------
if not exist creds.local.cmd (
    copy creds.local.cmd.example creds.local.cmd >nul
    echo [action] creds.local.cmd created - paste your three "set AWS_..." lines into it.
) else (
    echo [ok] creds.local.cmd already exists.
)

echo.
echo Setup done. Next: edit creds.local.cmd with fresh keys, then run:  run.cmd
