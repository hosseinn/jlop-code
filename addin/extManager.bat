
@echo off

setlocal

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

"%LO%\program\unopkg" gui

:: unopkg -h


