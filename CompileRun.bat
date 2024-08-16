call tsc
echo Compiled TS.
copy %cd%\src\taskpane.html dist > NUL
copy %cd%\src\taskpane.css dist > NUL
echo Compiled files in ./dist
call ts-node Server.ts
pause