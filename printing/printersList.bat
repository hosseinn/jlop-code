@echo off

rem http://www.techrepublic.com/blog/windows-and-office/discover-the-power-of-windows-7-hidden-vbscript-print-utilities/


cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs  -l | findstr /C:"Printer name"

echo.

cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs  -g | findstr "default"
