@echo off
Pushd "%~dp0"
:: For POWERSHELL V3+ only
::
:: STRUCTURE:
::
:: \SomeName.bat			- this file
:: \SomeName.ps1			- ps1 file with same name as this file
:: \logs					- log directory

SET scriptPathFull=%~dp0
SET scriptPath=%scriptPathFull:~0,-1%
SET BATfileName=%~n0
SET "logFile=%scriptPath%\logs\%COMPUTERNAME%_%date:~-4,4%%date:~-10,2%%date:~-7,2%-%time:~-11,2%%time:~-8,2%%time:~-5,2%-BAT.log"
SET "pslogFile=%scriptPath%\logs\%COMPUTERNAME%_%date:~-4,4%%date:~-10,2%%date:~-7,2%-%time:~-11,2%%time:~-8,2%%time:~-5,2%-PS.log"
SET "psFile=%BATfileName%.ps1"

call:dtecho Starting 
call:dtecho **************************************************************************** 
call:dtecho  "RUNNING AS (BAT): %USERNAME%" 
call:dtecho %scriptPathFull% 
call:dtecho scriptPath = %scriptPath% 
call:dtecho BATfileName = %BATfileName% 
call:dtecho logFile = %logFile% 
call:dtecho psFile = %psFile% 
call:dtecho psLogFile = %psLogFile% 
call:dtecho **************************************************************************** 
call:dtecho executing PS file %localtmpdir%\%psFile% 
call:dtecho **************************************************************************** 

powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -c "%scriptPath%\%psFile%" *> \"%psLogFile%\""

:: Example passing args
:: powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -c "%scriptPath%\%psFile% -batLogFile \"%logFile%\" -psLogFile \"%psLogFile%\" *> \"%psLogFile%\""

call:dtecho "End now my BAT has ended." 

goto:eof
::-----------------------------------------------------------------------
:: funcs
::-----------------------------------------------------------------------
:dtecho
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)
echo %mydate%-%TIME% %* >> %logFile% 2>>&1
::echo %mydate%-%TIME% %* 
goto:eof