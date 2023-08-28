@echo off

rem http://www.techrepublic.com/blog/windows-and-office/discover-the-power-of-windows-7-hidden-vbscript-print-utilities/

if [%1]==[] (
  echo Usage: printerStatus printer-name
  exit /B
)

cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prncnfg.vbs  -g -p %1 | findstr "status"

