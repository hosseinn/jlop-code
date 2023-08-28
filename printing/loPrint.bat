@echo off

rem https://help.libreoffice.org/Common/Starting_the_Software_With_Parameters

setlocal

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
rem remove quotes around input

echo Executing LibreOffice...

rem no args then give help
if [%1]==[] (
  "%LO%\program\soffice.exe" -h
  exit /B
)

rem if one arg use the default printer;
rem if two args, use the first as the printer name
if [%2]==[] (
  "%LO%\program\soffice.exe" -p %1
) else (
  "%LO%\program\soffice.exe" --pt %1 %2
)

echo Finished.
