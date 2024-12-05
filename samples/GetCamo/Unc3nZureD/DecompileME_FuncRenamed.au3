#RequireAdmin
#NoTrayIcon

#include <ProcessConstants.au3>
#include <Process.au3>

; No Explorer running ? ..or renamed file?  -> Die!!!
If @ScriptName <> "DecompileME.exe" Then Exit
If WinGetText("Program Manager") = "0" Then Exit

; On every second run, run the script as an interactive service process.
If RegRead("HKLM\SOFTWARE\DME", "S") = 1 Then
	RegDelete("HKLM\SOFTWARE\DME")
Else
	$COMMAND = 'cmd /c   sc create -- ' & _
			'binPath= "cmd /c   start  \"\"  \"' & @ScriptFullPath & '\" " ' & _
			'type= own ' & _
			'type= interact & ' & _
				'net start -- & ' & _
				'sc delete --' '
	_RunDos($COMMAND)

	RegWrite("HKLM\SOFTWARE\DME", "S", "REG_SZ", 1)
	Exit
EndIf

ToolTip("Program Started!", 0, 0)

While 1

	;Hide OllyDebug...
	ControlHide("", "", "[Class:ACPUDUMP]" ) ; CPU Window/Dump
	ControlHide("", "", "[Class:ACPUASM]"  ) ; CPU Window/Asm
	ControlHide("", "", "[Class:ICPUASM]"  ) ; ???
	ControlHide("", "", "[Class:ACPUSTACK]") ; CPU Window/Stack

	ControlHide("", "", "[Class:APROCESS]" ) ; Attach to process Window

	; Close ...
	WinKill("[CLASS:HexWorksClass]")		; Hex Workshop (http://www.hexworkshop.com/)
	WinKill("[CLASS:PROCMON_WINDOW_CLASS]")	; Process Monitor
	WinKill("[CLASS:PROCEXPL]")				; Process Explorer

	; More Generic approaches...
	WinKill("", "Breakpoint" )
	WinKill("", "Hex"        )
	WinKill("Hex", ""        )
	WinKill("", "Memory View")
	WinKill("", "Unpack"     )
	WinKill("", "Attach"     )
	WinKill("", "Entrypoint" )
	WinKill("", "OEP"        )
	WinKill("", "Rebuild PE" )
	WinKill("", "inject"     )
	WinKill("", "AHTeam"     )
	WinKill("", "disasm"     )
	WinKill("", "suspend"    )
	WinKill("", "freeze"     )

	Sleep( 50 )
	If 1 = 2 Then ExitLoop
WEnd


MsgBox(64, "Congratz!", "You successfully unpacked the file!")
; DeTokenise by myAut2Exe >The Open Source AutoIT/AutoHotKey script decompiler< 2.15 build(213)
