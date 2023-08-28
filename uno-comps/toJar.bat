
@echo off
:: toJar RandomSents
:: Compile the Java code and package into a JAR along with
:: a special manifest

IF [%1] == [] (
  echo No service name supplied
  EXIT /B
)


echo Compiling %1Impl.java
call compileOrg %1 %1Impl.java
echo.

IF NOT EXIST Manifest%1.txt (
  echo No Manifest%1.txt file supplied
  EXIT /B
)

echo Generating %1Impl.jar
jar cvfm %1Impl.jar Manifest%1.txt %1Impl.class %1Impl.java org

echo Finished.
