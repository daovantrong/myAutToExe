Attribute VB_Name = "Unpacker"
Option Explicit


Public Function UnUPX() As Long

   On Error GoTo UnUPX_err
   Dim cmdline$, parameters$, Logfile$
   
   Dim FileNameIn As ClsFilename
   Set FileNameIn = FileName
   
   
   Const RESULT_OK = 0 'Default
   Const RESULT_ERR_NOT_UPX = 1
   Const RESULT_ERR_UNPACKFAIL = 2
   
   
   With FrmMain
         
         .Log_Stage "UPX Decompress", 0
         
       ' Check for already unpacked File
         Set FileName = FileNameIn 'in case we exit early...
         
         Dim FileNameOut As New ClsFilename
         FileNameOut = FileNameIn
         FileNameOut.Ext = "unupx"
         
         
         If FileNameIn.Ext = FileNameOut.Ext Then _
            Exit Function
         
         If FileExists(FileNameOut) Then
            .Log "Found previously UPX decompressed file using that..."
            Set FileName = FileNameOut
            Exit Function
         End If
         
         
        ' PreCheck is packed with UPX
         .log_verbose ""
         .log_verbose "Testing to see if exe is UPX compressed..."


         Dim filedata As New StringReader
         filedata.Data = FileLoad(FileNameIn.FileName, &H1000)

         If filedata.FixedString(2) <> "MZ" Then
            .log_verbose "NO MZ header found exiting..."
            Exit Function
         End If

         'Dim tmp As String
         'tmp = StrConv(FileData, vbUnicode, LocaleID)
         If filedata.FindString("UPX") = 0 Then
            .log_verbose "No UPX marker found exiting..."
            Exit Function
         End If
       ' ------------------------------
         
       
       
       ' Test for AutoIt2 Script
         File.create FileNameIn.FileName
         
         Dim isOldScript As Boolean
         isOldScript = TestForV2_0
         
         File.CloseFile
         
       ' unUPX AutoIt2 Scripts (without fixing the Pointer to script at the end) will break them
       ' so stop when AutoIt2 Script
         If isOldScript Then
            .log_verbose "unUPX stop because an AutoIt2 script was found!"
            Exit Function
         End If

         
       
       
       
       ' Invoke upx.exe
         .Log "Trying to decompress UPX binary..."

         cmdline = App.Path & "\data\upx.exe"
         If Not FileExists(cmdline) Then
            .Log "upx.exe binary not found!"
            .Log cmdline
            Exit Function
         End If
         
         parameters = Join(Array( _
               "-d", Quote(FileNameIn), _
               "-o", Quote(FileNameOut)), " ")
         .Log cmdline & " " & parameters
         
         Dim upxExitCode&
         
         'Dim ConsoleOut$
         'ConsoleOut =
         FrmMain.Console.ShellExConsole cmdline, parameters, upxExitCode
         UnUPX = upxExitCode
         If upxExitCode = 0 And FileExists(FileNameOut) Then
             .Log "=> UPX Decompress Okay!"
             Set FileName = FileNameOut
         Else
            .Log "=> Error (ExitCode: " & upxExitCode & ")"
            If upxExitCode = 1 Then
               .Log "Attention: upx decompress failed. Probably file was modified and requires manual unpacking/dumping."
               .Log "           On newer Autoit3 files (which stores the script in the .rsrc section of the exe) this error is critical since MATE can't work on packed data."
               
            ElseIf upxExitCode = 2 Then
               .Log "invalid file or not packed with UPX."
               
            End If

            
         End If
         
   End With

Exit Function

UnUPX_err:

myMsgBox Err.Description, vbCritical, "Error in UnUPX"

End Function


