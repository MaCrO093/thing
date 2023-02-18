@setlocal DisableDelayedExpansion
@echo off



::========================================================================================================================================

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

if exist %SystemRoot%\Sysnative\cmd.exe (
set "_cmdf=%~f0"
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %*"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 (
set "_cmdf=%~f0"
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %*"
exit /b
)

::  Set Path variable, it helps if it is misconfigured in the system

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SysPath%\WindowsPowerShell\v1.0\"

::========================================================================================================================================

cls
color 07
title  Word + Powerpoint Recoverer

set _elev=
if /i "%~1"=="-el" set _elev=1

set winbuild=1
set "nul=>nul 2>&1"
set "_psc=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 2>nul | find /i "0x0" 1>nul && (set _NCS=0)


::========================================================================================================================================

if %winbuild% LSS 7600 (
%nceline%
echo Unsupported OS version detected.
echo This program is supported only for Windows 10
goto shitSend
)

if not exist %_psc% (
%nceline%
echo Powershell is not installed in the system.
echo Aborting...
goto shitSend
)

::========================================================================================================================================

::  Fix for the special characters limitation in path name

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set "_PSarg="""%~f0""" -el %_args%"

set "_ttemp=%temp%"

setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" 1>nul && (
%nceline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto shitSend
)

::========================================================================================================================================

setlocal DisableDelayedExpansion

::  Check desktop location

set _desktop_=
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_desktop_=%%b"
if not defined _desktop_ for /f "delims=" %%a in ('%_psc% "& {write-host $([Environment]::GetFolderPath('Desktop'))}"') do call set "_desktop_=%%a"

set "_pdesk=%_desktop_:'=''%"
setlocal EnableDelayedExpansion
set "mastemp=%SystemRoot%\Temp\__shit"

::========================================================================================================================================


:MainMenu
cls
color 07
title Word + Powerpoint Recoverer
mode 50, 30
if exist "%mastemp%\.*" rmdir /s /q "%mastemp%\" %nul%

echo:                                             
mode con: cols=100 lines=40
echo: Welcome to Word + Powerpoint Recoverer by marco because he has too much free time                                                                     

SET choice=
SET /p choice=Would you like to start? Y/N: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO :choice 
IF '%choice%'=='y' GOTO :choice 
IF '%choice%'=='N' GOTO :nope
IF '%choice%'=='n' GOTO :nope
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu
     

::========================================================================================================================================

:+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

::  Extract the text from batch script without character issue

:_Export

%nul% %_psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('%~2',$f[1].Trim(),[System.Text.Encoding]::%~3);"
exit /b

:oemexport

%nul% %_psc% "$f=[io.file]::ReadAllText('!_batp!') -split \":%~1\:.*`r`n\"; [io.file]::WriteAllText('!_pdesk!\$OEM$\$$\Setup\Scripts\%~2',$f[1].Trim(),[System.Text.Encoding]::ASCII);"
exit /b

::========================================================================================================================================

:nope
cls
echo Exiting...
timeout /t 1 >nul
exit /b



::========================================================================================================================================

:choice
SET choice=
SET /p choice=Word or Powerpoint? W/P: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='W' GOTO :word
IF '%choice%'=='w' GOTO :word
IF '%choice%'=='P' GOTO :powerpoint
IF '%choice%'=='p' GOTO :powerpoint
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu


::========================================================================================================================================

:word
cls
echo searching autosave locations...
echo x=msgbox("If you are unable to open any of the autosaved documents then at the end you will have the option to goto a website that can convert them and download them so they work" ,0, "Info") >> msgbox.vbs
attrib +h msgbox.vbs
start msgbox.vbs
pause
timeout 1 >nul
echo Autosaves found
timeout 1 >nul
echo Opening 1st location...
%SystemRoot%\explorer.exe "%UserProfile%\Roaming\Microsoft\Word\"
echo if it opens the "documents" folder then the location doesnt work/exist
pause 
echo Opening 2nd location...
%SystemRoot%\explorer.exe "%UserProfile%\AppData\Local\Microsoft\Word"
pause
echo Opening 3rd location...
%SystemRoot%\explorer.exe "%UserProfile%\AppData\Local\Temp"
pause
echo Opening 4th location...
%SystemRoot%\explorer.exe "%UserProfile%\AppData\Local\Microsoft\Office\UnsavedFiles"
pause
echo Opening 5th location...
%SystemRoot%\explorer.exe "%UserProfile%AppData\Roaming\Microsoft\Word"
pause
timeout 1 >nul
SET choice=
SET /p choice=Did you find the file(s) you needed? if so did they work when you ran it? Y/N: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO :found
IF '%choice%'=='y' GOTO :found
IF '%choice%'=='N' GOTO :nofound
IF '%choice%'=='n' GOTO :nofound
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu




::========================================================================================================================================

:powerpoint
cls
echo searching autosave locations...
echo x=msgbox("If you are unable to open any of the autosaved powerpoints then at the end you will have the option to goto a website that can convert them and download them so they work" ,0, "Info") >> msgbox.vbs
attrib +h msgbox.vbs
start msgbox.vbs
pause
timeout 1 >nul
echo Autosaves found
timeout 1 >nul
echo Opening 1st location...
%SystemRoot%\explorer.exe "%UserProfile%\Roaming\Microsoft\PowerPoint\"
echo if it opens the "documents" folder then the location doesnt work/exist
pause 
echo Opening 2nd location...
%SystemRoot%\explorer.exe "%UserProfile%\AppData\Local\Microsoft\PowerPoint"
pause
echo Opening 3rd location...
%SystemRoot%\explorer.exe "%UserProfile%AppData\Roaming\Microsoft\PowerPoint"
pause
echo Opening 4th location...
%SystemRoot%\explorer.exe "%UserProfile%\AppData\Local\Temp"
pause
echo Opening 5th location...
%SystemRoot%\explorer.exe "%UserProfile%\AppData\Local\Microsoft\Office\UnsavedFiles"
pause
timeout 1 >nul
SET choice=
SET /p choice=Did you find the file(s) you needed? if so did they work when you ran it? Y/N: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO :found
IF '%choice%'=='y' GOTO :found
IF '%choice%'=='N' GOTO :nofound
IF '%choice%'=='n' GOTO :nofound
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu

::========================================================================================================================================

:found
cls
echo Awesome, call me over so i can get the usb back
timeout 1 >nul
SET choice=
SET /p choice=Would you like to go back to the start? Y/N: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO :MainMenu
IF '%choice%'=='y' GOTO :MainMenu
IF '%choice%'=='N' GOTO :nope
IF '%choice%'=='n' GOTO :nope
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu


::========================================================================================================================================

:nofound

start https://cloudconvert.com
timeout 2 >nul
echo x=msgbox("Sorry if the website opened in the wrong browser Upload the file and select docx as the format  Once downloaded it you can open it however your text format and font might be stuffed up" ,0, "Info") >> msgbox.vbs
attrib +h msgbox.vbs
start msgbox.vbs           
pause

::========================================================================================================================================

:: add something here for if that doesnt work loL
:stillnofound
cls
SET choice=
SET /p choice=Did you have the file opened for more then 10 minutes before you lost it? Y/N: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO :lasttry
IF '%choice%'=='y' GOTO :lasttry
IF '%choice%'=='N' GOTO :ripfile
IF '%choice%'=='n' GOTO :ripfile
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu

::========================================================================================================================================


:lasttry


::========================================================================================================================================

:ripfile
cls
SET choice=
SET /p choice=Your probably out of luck and have to start again, but do you still want to try the file recoverer? Y/N: 
IF NOT '%choice%'=='' SET choice=%choice:~0,1%
IF '%choice%'=='Y' GOTO :lasttry
IF '%choice%'=='y' GOTO :lasttry
IF '%choice%'=='N' GOTO :byeiguess
IF '%choice%'=='n' GOTO :byeiguess
ECHO "%choice%" is not valid
ECHO.
GOTO MainMenu

::========================================================================================================================================

:byeiguess

echo Sorry this happened, next time try setting the autosave to 5 minutes instead of 10 or using onedrive more
timeout /t 10 