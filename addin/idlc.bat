
@echo off

:: idlc RandomSents
:: converts RandomSents.idl into RandomSents.urd

:: idlc Doubler
:: converts Doubler.idl into Doubler.urd

:: if you're not admin, try:
:: elevate.exe -k -w idlc.bat RandomSents

setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%


:: add DLLs needed by idlc.exe
SET PATH=%PATH%;%LO%\program
:: echo %PATH%

:: due to preprocessor, idlc.exe must be called from <OFFICE>\sdk\bin
set CURRDIR="%~dp0"
echo Current dir: %CURRDIR%

echo Copying %1.idl to "%LO%"\sdk\bin
copy %1.idl "%LO%"\sdk\bin
echo.

cd "%LO%"\sdk\bin
:: idlc -h

echo Compiling %1 with idlc...
idlc -w -verbose -C -I..\idl %1.idl 

:: copy stuff back and clean up
echo Copying %1.urd to %CURRDIR%
copy %1.urd  %CURRDIR%
del %1.idl %1.urd
cd %CURRDIR%

echo Finished.
