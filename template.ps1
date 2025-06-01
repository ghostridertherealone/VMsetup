@echo off
setlocal EnableDelayedExpansion
set "chars=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
set "randomName="
for /L %%i in (1,1,7) do (
    set /a "index=!random! %% 36"
    for %%j in (!index!) do set "randomName=!randomName!!chars:~%%j,1!"
)
wmic computersystem where name="%COMPUTERNAME%" call rename name="DESKTOP-!randomName!" >nul 2>&1
