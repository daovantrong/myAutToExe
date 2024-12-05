Attribute VB_Name = "GlobalDefs"
Option Explicit

Public File As New FileStream
Public FileName As New ClsFilename

Public Const ERR_NO_AUT_EXE& = vbObjectError + &H10
Public Const ERR_NO_OBFUSCATE_AUT& = vbObjectError + &H20

Public Const DE_OBFUSC_TYPE_NOT_OBFUSC& = 0
Public Const DE_OBFUSC_TYPE_VANZANDE& = 1
Public Const DE_OBFUSC_TYPE_ENCODEIT& = 2

Public Const NO_AUT_DE_TOKEN_FILE& = &H100

Public ExtractedFiles As Collection

Public IsCommandlineMode As Boolean
Public IsOpt_QuitWhenFinish As Boolean
Public IsOpt_RunSilent As Boolean





Sub DoEventsSeldom()
   If Rnd < 0.01 Then DoEvents
End Sub

Sub DoEventsVerySeldom()
   If Rnd < 0.00001 Then
      DoEvents
   End If
End Sub

Sub ShowScript(ScriptData$)
   If isUTF16(ScriptData) Then
      FrmMain.Txt_Script = StrConv((Mid(ScriptData, 1 + Len(UTF16_BOM))), vbFromUnicode)
   ElseIf isUTF8(ScriptData) Then
      FrmMain.Txt_Script = Mid(ScriptData, 1 + Len(UTF8_BOM))
   Else
      FrmMain.Txt_Script = ScriptData
   End If

End Sub

Sub SaveScriptData(ScriptData$)

   With FrmMain
   ' Adding a underscope '_' for lines longer than 2047
   ' so Tidy will not complain
      FrmMain.Log "Try to breaks very long lines (about 2000 chars) by adding '_'+<NewLine> ..."
      ScriptData = AddLineBreakToLongLines(Split(ScriptData, vbCrLf))
   
'debug
'FrmMain.Chk_TmpFile.Value = vbChecked
   
    ' overwrite script
      If FrmMain.Chk_TmpFile.Value = vbChecked Then
         FileName.Name = FileName.Name & "_restore"
         .Log "Saving script to: " & FileName.FileName
      Else
'         FileDelete FileName.Name
         .Log "Save/overwrite script to: " & FileName.FileName
      End If

     
     File.Create FileName.FileName, True, False, False
     File.Position = 0
     File.FixedString(-1) = ScriptData
     File.setEOF
     File.CloseFile
     
     ShowScript ScriptData
     
     .Log ""
     .Log "Running 'Tidy.exe " & FileName.NameWithExt & "' to improve sourcecode readability."
     
     Dim cmdline$, parameters$, Logfile$
     cmdline = App.Path & "\Tidy\Tidy.exe"
     parameters = """" & FileName & """" ' /KeepNVersions=1
     .Log cmdline & " " & parameters
     
     Dim TidyExitCode&
     TidyExitCode = ShellEx(cmdline, parameters, vbNormalFocus)
     If TidyExitCode = 0 Then
         .Log "=> Okay (ExitCode: " & TidyExitCode & ")."
         Dim TidyBackupFileName As New ClsFilename
         TidyBackupFileName.mvarFileName = FileName.mvarFileName
         TidyBackupFileName.Name = TidyBackupFileName.Name & "_old1"
         
       ' Delete Tidy BackupFile
         If FrmMain.Chk_TmpFile.Value = vbUnchecked Then
            .Log "Deleting Tidy BackupFile..." ' & TidyBackupFileName.NameWithExt
            FileDelete TidyBackupFileName.FileName
         End If
        
        
        
        File.Create FileName.FileName
        ScriptData = File.FixedString(-1)
        File.CloseFile
      
        ShowScript ScriptData
        
     Else
        .Log "Tidy.exe ExitCode: " & TidyExitCode & " =>some failure!"
        .Log "Attention: Tidy.exe failed. Deobfucator will probably also fail because scriptfile is not in proper format."
     End If
  End With
End Sub

