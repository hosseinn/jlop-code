
:: convert.bat 
:: usage: convert fnm extension[:filter-name]

:: usage:   convert payment.ods "xml:Pay"    
::      // only if pay filters added to Office, and there's a payment.ods file.


@echo off

setlocal


IF [%1] == [] (
  echo No document name supplied
  echo Usage: convert FILENAME EXT[:FILTER-NAME]
  EXIT /B
)

IF [%2] == [] (
  echo No filtername supplied
  echo Usage: convert FILENAME EXT[:FILTER-NAME]
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
:: remove quotes around input

set OFFICE="%LO%\program\soffice.exe"
:: echo Calling %OFFICE%

:: %OFFICE% -h

%OFFICE% --convert-to %2 %1
