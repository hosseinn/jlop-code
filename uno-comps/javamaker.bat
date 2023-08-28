
@echo off
:: javamaker RandomSents
:: Generating Java org.openoffice package and
:: interface code using the RDB info and types.rdb
:: Also decompile the interface using cfg into X<nm>.java

:: uses tr.exe from Gow (https://github.com/bmatzelle/gow/wiki)
:: uses Java decompiler, cfr (http://www.benf.org/other/cfr/)

setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

:: add DLLs needed by javamaker.exe
SET PATH=%PATH%;%LO%\program
:: echo %PATH%

set CURRDIR="%~dp0"

set INTERFACE_NM=X%1
:: echo Interface name is %INTERFACE_NM%

echo %1| tr "[A-Z]" "[a-z]" > tempComp.txt
set /p PACK_NM=<tempComp.txt
del tempComp.txt
:: echo Package name is %PACK_NM%

:: "%LO%"\sdk\bin\javamaker
:: for help

echo Generating Java directories and classes for %1.rdb in org/
"%LO%"\sdk\bin\javamaker -Torg.openoffice.%PACK_NM%.%INTERFACE_NM% -nD "%LO%"\program\types.rdb %1.rdb
echo. 

:: javap --help
:: javap org.openoffice.%PACK_NM%

echo Decompiling %INTERFACE_NM%.class using cfr
java -jar cfr_0_118.jar org.openoffice.%PACK_NM%.%INTERFACE_NM% --outputdir "."

echo Finished.
