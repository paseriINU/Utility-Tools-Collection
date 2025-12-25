@echo off
title JP1 Job Log Sample
setlocal enabledelayedexpansion

rem ============================================================================
rem JP1 Job Log Sample
rem
rem Description:
rem   Call JP1_REST_Job_Info_Tool.bat and save the result to a file.
rem
rem Usage:
rem   1. Edit UNIT_PATH below
rem   2. Edit OUTPUT_FILE below
rem   3. Double-click to run
rem ============================================================================

rem === Edit here ===
set "UNIT_PATH=/JobGroup/Jobnet"
set "OUTPUT_FILE=%~dp0joblog_output.txt"

rem Get script directory
set "SCRIPT_DIR=%~dp0"

echo.
echo ================================================================
echo   JP1 Job Log Sample
echo ================================================================
echo.
echo   Target: %UNIT_PATH%
echo   Output: %OUTPUT_FILE%
echo.

rem Call JP1_REST_Job_Info_Tool.bat
echo Getting log...

rem Save result to temp file
set "TEMP_FILE=%TEMP%\jp1_log_%RANDOM%.txt"

call "%SCRIPT_DIR%JP1_REST_ジョブ情報取得ツール.bat" "%UNIT_PATH%" > "%TEMP_FILE%" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"

rem Error check
if %EXIT_CODE% neq 0 (
    echo.
    echo [ERROR] Failed to get log (exit code: %EXIT_CODE%)
    echo.
    if exist "%TEMP_FILE%" (
        echo Error details:
        type "%TEMP_FILE%"
        del "%TEMP_FILE%" >nul 2>&1
    )
    echo.
    pause
    exit /b %EXIT_CODE%
)

rem Check if result is empty
for %%A in ("%TEMP_FILE%") do set "FILE_SIZE=%%~zA"
if "%FILE_SIZE%"=="0" (
    echo.
    echo [WARNING] Result is empty
    echo.
    del "%TEMP_FILE%" >nul 2>&1
    pause
    exit /b 1
)

rem Check for ERROR lines
findstr /B "ERROR:" "%TEMP_FILE%" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo.
    echo [ERROR] API call failed
    echo.
    type "%TEMP_FILE%"
    echo.
    del "%TEMP_FILE%" >nul 2>&1
    pause
    exit /b 1
)

rem Save to output file
copy /Y "%TEMP_FILE%" "%OUTPUT_FILE%" >nul

echo.
echo ================================================================
echo   Complete - Saved to file
echo ================================================================
echo.
echo Output file: %OUTPUT_FILE%
echo.
echo Content:
echo ----------------------------------------
type "%TEMP_FILE%"
echo ----------------------------------------
echo.

rem Delete temp file
del "%TEMP_FILE%" >nul 2>&1

pause
exit /b 0
