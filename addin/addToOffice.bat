
@echo off

:: addInOXT RandomSents
setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

echo ------ toJar ------
call toJar %1

echo.
echo ------ makeOxt ------
call makeOxt %1

echo.
echo ------ pkg ------
call pkg %1
