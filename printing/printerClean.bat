@echo off

rem http://www.techrepublic.com/blog/windows-and-office/discover-the-power-of-windows-7-hidden-vbscript-print-utilities/

if [%1]==[] (
  echo Usage: printerClean printer-name
  exit /B
)

cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnqctl.vbs  -x -p %1

