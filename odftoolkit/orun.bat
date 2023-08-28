@echo off
setlocal

echo Executing %* with Apache Simple ODF Toolkit...

java -cp "D:\odftoolkit\*;D:\xerces-2_11_0\*;D:\apache-jena-2.12.1\lib\*;."  %*

echo Finished.
