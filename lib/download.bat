for /f %%i in (links.txt) do curl -L -C - -O %%i
