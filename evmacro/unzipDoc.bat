
:: unzipDoc.bat

:: unzips to <FNM>_<EXT>\
:: usage: unzipDoc FormDocMacros.odt

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

:: list contents of zipped doc
%ZIP% l %1

:: create a work dir name from the doc name
set fnm=%~n1
set ext=%~x1
set WORK_DIR=%fnm%_%ext:~1%

rmdir /s /q %WORK_DIR%

echo Extracting %1 to %WORK_DIR%
%ZIP% x %1 -aoa -o%WORK_DIR%

