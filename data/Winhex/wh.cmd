if /i "%PROCESSOR_IDENTIFIER:~0,3%" == "X86" goto 86
WinHex64.exe
exit
:86
WinHex.exe
exit
