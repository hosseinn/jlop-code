@echo off

echo Executing %* with Jaybird...

java -cp "D:\jaybird\jaybird-full-2.2.10.jar;." -Djava.library.path="D:\jaybird" %*

echo Finished.
