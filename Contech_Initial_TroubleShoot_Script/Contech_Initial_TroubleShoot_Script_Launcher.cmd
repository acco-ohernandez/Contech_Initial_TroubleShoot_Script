@ECHO OFF
@setlocal enableextensions
@cd /d "%~dp0"

REM  --> Check for permissions
"%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system">nul 2>NUL

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    if exist "%temp%\getadmin.vbs" (
        del "%temp%\getadmin.vbs"
        echo Failed to acquire elevated privilege.  Try saving this script and running it from your Desktop.
        echo;
        echo Press any key to exit.
        pause>NUL
        goto :EOF
    )
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "%*", "", "runas", 1 >> "%temp%\getadmin.vbs"

    cscript /nologo "%temp%\getadmin.vbs"
    goto :EOF

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
:--------------------------------------

cls


rem Powershell.exe -noprofile -executionpolicy bypass -file ".\Engineering-Lenovo_RunAll.ps1"
start Powershell.exe -noprofile -executionpolicy bypass -file ".\Contech_Initial_TroubleShoot_Script.ps1"