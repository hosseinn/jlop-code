@echo off
setlocal

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
rem remove quotes around input

rem echo Using %LO%
echo Compiling %* with LibreOffice SDK, JNA, Utils, and FreeTTS...

rem -Xlint:deprecation
rem "%LO%\URE\java\*;" only in LO 4

javac  -cp "%LO%\program\classes\*;%LO%\URE\java\*;..\Utils;D:\jna\jna-4.1.0.jar;D:\jna\jna-platform-4.1.0.jar;D:\freetts-1.2\lib\freetts.jar;."  %*

echo Finished.
