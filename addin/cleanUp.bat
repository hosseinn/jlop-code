@echo off
:: cleanUp Doubler

setlocal

IF [%1] == [] (
  echo No file supplied
  EXIT /B
)

echo - Removing URD and RDB files
del %1.urd
del %1.rdb

echo - Removing Java interface code
rmdir /s /q org

echo - Removing Java .class, JAR, OXT
del %1Impl.class
del %1Impl.jar
del %1.oxt

echo Finished.
