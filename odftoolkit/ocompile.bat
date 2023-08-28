@echo off
setlocal

echo Compiling %* with Apache Simple ODF Toolkit...

javac -cp "D:\odftoolkit\*;."  %*

echo Finished.
