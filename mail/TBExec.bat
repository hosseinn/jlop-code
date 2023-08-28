@echo off

rem TBExec "ad@fivedots.coe.psu.ac.th" "Test1" "Test 1" skinner.png
rem http://kb.mozillazine.org/Command_line_arguments_(Thunderbird)

setlocal

set PF=%ProgramFiles(x86)%
if [%PF%]==[] (
  set PF=%ProgramFiles%
)
rem echo %PF%

rem no args then give help
if [%1]==[] (
  echo Usage TBExec "to" "subject" "body" [ "attachment file in this dir" ]
  exit /B
)

echo Calling Thunderbird...

if [%4]==[] (
  "%PF%\Mozilla Thunderbird\thunderbird.exe" -compose "to='%1',subject='%2',body='%3'" 
) else (
  "%PF%\Mozilla Thunderbird\thunderbird.exe" -compose "to='%1',subject='%2',body='%3',attachment='%CD%\%4'" 
)

echo Finished.
