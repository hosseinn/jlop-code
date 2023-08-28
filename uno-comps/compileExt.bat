
@echo off
:: compileExt RandomSents PoemCreator.java
:: - requires service name and java file name

:: uses tr.exe from Gow (https://github.com/bmatzelle/gow/wiki)


setlocal

IF [%1] == [] (
  echo No service name or Java file supplied
  EXIT /B
)

IF [%2] == [] (
  echo No service name or Java file supplied
  EXIT /B
)


call lofind.bat
set /p LOQ=<lofindTemp.txt
:: echo %LOQ%
SET LO=%LOQ:"=%


echo %1| tr "[A-Z]" "[a-z]" > tempComp.txt
set /p PACK_NM=<tempComp.txt
del tempComp.txt
:: echo Package name is %PACK_NM%


call run FindExtJar org.openoffice.%PACK_NM%
set /p JARQ=<lofindTemp.txt

IF %JARQ%=="xx" (
  echo No JAR found
  EXIT /B
)

SET JAR=%JARQ:"=%
echo.
echo Using %1 JAR: %JAR%

:: echo Using %LO%
echo Compiling %2 with LibreOffice SDK, JNA, Utils, and %1...

javac  -cp "%LO%\program\classes\*;%LO%\URE\java\*;..\Utils;D:\jna\jna-4.1.0.jar;D:\jna\jna-platform-4.1.0.jar;%JAR%;."  %2

echo Finished.

