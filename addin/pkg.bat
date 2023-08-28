
@echo off
:: pkg RandomSents
:: Install the Oxt file using unopkg.exe

:: uses tr.exe from Gow (https://github.com/bmatzelle/gow/wiki)

setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

:: add unopkg.exe to path
SET PATH=%PATH%;%LO%\program
:: echo %PATH%


echo %1| tr "[A-Z]" "[a-z]" > tempAddIn.txt
set /p PACK_NM=<tempAddIn.txt
del tempAddIn.txt
:: echo Type name is %PACK_NM%

:: unopkg -h

echo Attempting to remove old version of %1.oxt with unopkg...
unopkg remove org.openoffice.%PACK_NM%

echo Installing %1.oxt with unopkg...
unopkg add -v %1.oxt
unopkg gui
::unopkg list

echo Finished.
