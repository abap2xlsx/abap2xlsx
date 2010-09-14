@echo on
shift
rem echo %2 %3 %4 %5 %6 %7 %8 %9
"%ProgramFiles%\TortoiseSVN\bin\TortoiseProc.exe" %1 %2 %3 %4 %5 %6 %7 %8 %9 > %0 2>&1
rem pause