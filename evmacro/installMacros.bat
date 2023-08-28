
@echo off
:: usage: installMacros FormMacros
:: zip into an OXT, install it in Office

:: uses tr.exe and zip.exe from Gow (https://github.com/bmatzelle/gow/wiki)

setlocal

IF [%1] == [] (
  echo No macro folder supplied
  EXIT /B
)


call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

:: add unopkg.exe to path
SET PATH=%PATH%;%LO%\program
:: echo %PATH%


:: zipping --------------------------------------

if exist %1.oxt (
  echo Removing old OXT file 
  del %1.oxt
)

cd %1
echo Zipping %1\ into %1.oxt
zip -r ..\%1.oxt *.*
cd ..
echo.

:: installing OXT -------------------------------

echo %1| tr "[A-Z]" "[a-z]" > tempAddIn.txt
set /p LOWER_NM=<tempAddIn.txt
del tempAddIn.txt

set EXT_ID=org.openoffice.%LOWER_NM%
echo Extension ID is %EXT_ID%

:: unopkg -h

echo Attenpting to remove old version of %1.oxt using ID
unopkg remove %EXT_ID%
echo.

echo Installing %1.oxt with unopkg...
unopkg add -v %1.oxt
unopkg gui
::unopkg list

echo Finished.
