
@echo off
:: makeOXT Doubler
:: Zip up the JAR file and other files into an OXT file

:: uses zip.exe from Gow (https://github.com/bmatzelle/gow/wiki)


setlocal

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)

IF NOT EXIST %1\NUL (
  echo No service folder found; create %1\
  EXIT /B
)

IF NOT EXIST %1\META-INF\manifest.xml (
  echo No manifest.xml for OXT; create one
  EXIT /B
)

IF NOT EXIST %1\CalcAddIn.xcu (
  echo No CalcAddIn.xcu for OXT; create one
  EXIT /B
)

IF NOT EXIST %1\description.xml (
  echo No description.xml for OXT; it's a good idea to have one
)


echo Copying %1.rdb to %1\
copy %1.rdb %1\

echo Copying %1Impl.jar to %1\
copy %1Impl.jar %1\
echo.

echo Zipping %1\ as %1.oxt
cd %1\
:: zip -h 
zip -r ..\%1.oxt *.*
cd ..

echo Finished.
