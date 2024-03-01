

@echo off

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin

@echo off
echo 偵測 Microsoft Office 2021 安裝目錄
set OfficePath="C:\Program Files\Microsoft Office\Office16\"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set OfficePath="C:\Program Files (x86)\Microsoft Office\Office16\"
For /F "tokens=2 delims=[]" %%G in ('ver') Do (set _version=%%G) 
For /F "tokens=2 delims=. " %%G in ('echo %_version%') Do (set _major=%%G) 
if "%_major%"=="5" (echo Restart KMS server.
cscript %OfficePath%ospp.vbs /osppsvcrestart)
echo Setting up the KMS server.
echo Please replace "kms-server.com" with your own.
cscript %OfficePath%ospp.vbs /sethst:kms-server.com
cscript %OfficePath%ospp.vbs /setprt:1688
echo Activatiating Office 2021
cscript %OfficePath%ospp.vbs /act
echo Startup successful.
echo If you see Product activation successful it is now done.

echo "Office Installed Successful!"

pause
