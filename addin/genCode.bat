
@echo off

:: genCode Doubler
setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

echo ------ idlc ------
call idlc %1

:: call regview %1
echo.
echo ------ regmerge ------
call regmerge %1

echo.
echo ------ javamaker ------
call javamaker %1

echo.
echo ------ skelAddin ------
call skelAddin %1
