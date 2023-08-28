
:: infilter.bat 
:: usage: infilter fnm filter-name

:: usage:   infilter pay.xml "Pay"    // only if pay filters added to Office

@echo off

setlocal


IF [%1] == [] (
  echo No document name supplied
  echo Usage: infilter FILENAME FILTER-NAME
  EXIT /B
)

IF [%2] == [] (
  echo No filtername supplied
  echo Usage: infilter FILENAME FILTER-NAME
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
:: remove quotes around input

set OFFICE="%LO%\program\soffice.exe"
:: echo Calling %OFFICE%

:: %OFFICE% -h

%OFFICE% %1 --infilter=%2
