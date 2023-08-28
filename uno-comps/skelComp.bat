
@echo off
:: skelComp RandomSents
:: Generate a skeleton implementation of the Java interface as a component
:: The code is in <service-name>Impl.java, and must be completed by the user before moving on

:: uses tr.exe from Gow (https://github.com/bmatzelle/gow/wiki)

setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

call lofind.bat
set /p LOQ=<lofindTemp.txt
SET LO=%LOQ:"=%
:: echo Using %LO%

:: add DLLs needed by uno-skeletonmaker.exe
SET PATH=%PATH%;%LO%\program
:: echo %PATH%

SET CURRDIR="%~dp0"

set SERVICE_NM=%1
:: echo Using Service name: %SERVICE_NM%

echo %1| tr "[A-Z]" "[a-z]" > tempComp.txt
set /p PACK_NM=<tempComp.txt
del tempComp.txt
:: echo Type name is %PACK_NM%


set IMPL_FNM=%1Impl

:: "%LO%"\sdk\bin\uno-skeletonmaker.exe --help

:: uno-skeletonmaker only works if all the data is in <OFFICE>\sdk\bin
:: -- RDB file, org package

echo Copying %1.rdb to "%LO%"\sdk\bin
copy %1.rdb "%LO%"\sdk\bin

echo Copying Java classes in org/ to "%LO%"\sdk\bin
xcopy /e /i /q org "%LO%"\sdk\bin\org
echo.

cd "%LO%"\sdk\bin
:: uno-skeletonmaker.exe --help

echo Generating %IMPL_FNM%.java
uno-skeletonmaker component  -l "%LO%"\program\types.rdb -l %1.rdb  -n %IMPL_FNM% -t org.openoffice.%PACK_NM%.%SERVICE_NM%  
echo.

:: resulting <Service-name>Impl file must be moved to current dir 
echo Copying %IMPL_FNM%.java
copy %IMPL_FNM%.java  %CURRDIR%

echo Deleting copied RDB, org package, and original %IMPL_FNM%.java
rmdir /s /q org
del %1.rdb %IMPL_FNM%.java
cd %CURRDIR%
echo.

echo **TIME** for you to complete %IMPL_FNM%.java 
echo.

echo Finished.
