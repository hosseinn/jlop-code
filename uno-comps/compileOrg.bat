
@echo off
:: compileOrg RandomSents RandomSents-fnm.java
:: - requires service name and file name

:: uses tr.exe from Gow (https://github.com/bmatzelle/gow/wiki)

setlocal

IF [%1] == [] (
  echo No service name or file supplied
  EXIT /B
)

IF [%2] == [] (
  echo No service name or file supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

echo %1| tr "[A-Z]" "[a-z]" > tempComp.txt
set /p PACK_NM=<tempComp.txt
del tempComp.txt
:: echo Type name is %PACK_NM%


rem echo Using %LO%
echo Compiling %2 with LibreOffice SDK, JNA, Utils, and %1 service...

rem -Xlint:deprecation
rem "%LO%\URE\java\*;" only in LO 4

javac  -cp "%LO%\program\classes\*;%LO%\URE\java\*;..\Utils;D:\jna\jna-4.1.0.jar;D:\jna\jna-platform-4.1.0.jar;org\openoffice\%PACK_NM%;."  %2

echo Finished.
