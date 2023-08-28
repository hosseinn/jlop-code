@echo off
setlocal

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
rem remove quotes around input

rem echo Using %LO%
echo Executing %* with LibreOffice SDK, JNA, Utils, and FreeTTS...

rem -Xlint:deprecation
rem "%LO%\URE\java\*;" only in LO 4

java  -Dmbrola.base="C:\freetts-1.2\mbrola" -cp "%LO%\program\classes\*;%LO%\URE\java\*;..\Utils;C:\jna\jna-4.1.0.jar;C:\jna\jna-platform-4.1.0.jar;C:\freetts-1.2\lib\freetts.jar;."  %*

echo Finished.
