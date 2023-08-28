
:: zipDoc.bat

:: zips contents of <FNM>_<EXT>\ into new copy of file
:: usage: zipDoc FormDocMacros.odt

:: uses 7z.exe from 7-Zip (http://www.7-zip.org/)

@echo off

setlocal

IF [%1] == [] (
  echo No document name supplied
  EXIT /B
)

SET ZIP="C:\Program Files\7-Zip\7z.exe"

if not exist %ZIP% (
  echo Cannot find 7z.exe
  EXIT /B
) 

:: create a work dir name from the doc name
set fnm=%~n1
set ext=%~x1
set WORK_DIR=%fnm%_%ext:~1%

if exist %ZIP% (
  echo Deleting previously zipped doc
  del %1
) 

cd %WORK_DIR%

echo Updating %1 with contents of %WORK_DIR%
%ZIP%  u -tzip ..\%1 -r *

cd ..
:: list contents of zipped doc
:: %ZIP% l %1

