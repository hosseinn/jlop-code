
@echo off
:: installAddon Higher
:: Compile Java impl, jar, zip OXT, install OXT

:: uses tr.exe abd zip.exe from Gow (https://github.com/bmatzelle/gow/wiki)

setlocal

IF [%1] == [] (
  echo No Addon name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

:: add unopkg.exe to path
SET PATH=%PATH%;%LO%\program
:: echo %PATH%


:: compiling -----------------------------------

call compile %1AddonImpl.java
echo.


:: jaring --------------------------------------

IF NOT EXIST manifest.txt (
  echo No manifest.txt file supplied
  EXIT /B
)

echo Generating %1.jar
jar cvfm %1.jar manifest.txt %1AddonImpl.class

echo Moving JAR to %1\
move %1.jar %1\
echo.


:: mv Utils.jar %1\

cd %1

:: zipping --------------------------------------

echo Zipping %1\ into %1.oxt
zip -r ..\%1.oxt *.*
echo.

cd ..

:: installing OXT -------------------------------

echo %1| tr "[A-Z]" "[a-z]" > tempAddIn.txt
set /p LOWER_NM=<tempAddIn.txt
del tempAddIn.txt

set EXT_ID=org.openoffice.%LOWER_NM%Addon
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
