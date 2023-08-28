
:: execMacro.bat 

:: opens the file in Office and executes the specified macro
:: usage execMacro (u|s|d) FILENAME MACRO-NAME
:: e.g execMacro s build.odt ShowEvent.ShowEvent.show
::     execMacro u FormMacrosTest.odt Utils.ShowEvent.show


@echo off

setlocal

IF [%1] == [] (
  echo No macro category supplied
  echo Usage: execMacro u|s|d FILENAME MACRO-NAME
  EXIT /B
)

IF [%2] == [] (
  echo No document name supplied
  echo Usage: execMacro u|s|d FILENAME MACRO-NAME
  EXIT /B
)

IF [%3] == [] (
  echo No macro name supplied
  echo Usage: execMacro u|s|d FILENAME MACRO-NAME
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
:: remove quotes around input

set OFFICE="%LO%\program\soffice.exe"
:: echo Calling %OFFICE%

:: %OFFICE% -h


IF /I [%1] == [u] (
  %OFFICE% --nologo %2 "vnd.sun.star.script:%3?language=Java&location=user"
  EXIT /B
) 

IF /I [%1] == [s] (
  %OFFICE% --nologo %2 "vnd.sun.star.script:%3?language=Java&location=share"
  EXIT /B
) 

IF /I [%1] == [d] (
  %OFFICE% --nologo %2 "vnd.sun.star.script:%3?language=Java&location=document"
  EXIT /B
) ELSE (
  echo Unknown macro category
)