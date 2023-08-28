
@echo off
:: regview RandomSents
:: displays contents of RandomSents.urd

:: regview Doubler
:: displays contents of Doubler.urd

setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

"%LO%"\program\regview %1.urd

echo Finished.
