cls
@echo off
color 1f
title packages - registre...
if not "%1" == "max" start /max cmd /c %0 max & exit/b
start %comspec% /c "mode 40,10
>nul 2>&1 "%systemroot%\system32\cacls.exe" "%systemroot%\system32\config\system"
if '%errorlevel%' neq '0' (
    echo verification des privileges administrateur
    goto uacprompt
) else ( goto gotadmin )
:uacprompt
    echo set uac = createobject^("shell.application"^) > "%temp%\getadmin.vbs"
    set params = %*:"="
    echo uac.shellexecute "%~s0", "%params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /b
:gotadmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%cd%"
    cd /d "%~dp0"
color 0a
::----------------------------------------------- OCX---
echo GestionParc.dll
regsvr32 /s "GestionParc.dll"

timeout /t 0 /nobreak>nul
color 2f
echo msgbox"registre terminer avec succes..."+ vbnewline + vbnewline +"centra nord", vbokonly + vbinformation, "packages - installation">a.vbs&a.vbs&del a.vbs
exit
