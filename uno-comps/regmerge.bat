
@echo off
:: regmerge RandomSents
:: adds RandomSents.urd to RandomSents.rdb using UCR label

setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%

"%LO%"\program\regmerge -v %1.rdb UCR %1.urd

echo Finished.
