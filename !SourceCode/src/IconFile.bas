Attribute VB_Name = "IconFile"
Private Au3Standard_IconFileCrcs As New Collection

Public Sub HandleIconFile(FileName As String)
                   
         If Frm_Options.chk_extractIcon.value <> vbUnchecked Then
                      
         ' ==> Create output fileName
           Dim IconFileName As New ClsFilename
           IconFileName = FileName      ' initialise with ScriptPath
           IconFileName.Ext = ".ico"
           
           Log "Extracting ExeIcon/s to: " & Quote(IconFileName.FileName)
           On Error Resume Next
           ShellEx App.Path & "\" & "data\ExtractIcon.exe", _
                   Quote(File.FileName) & " " & Quote(IconFileName.FileName), vbNormalFocus
           If Err Then
               FrmMain.Log "ERROR: " & Err.Description
               Exit Sub
            End If
           
         ' Test For AutoItStandard
           If Frm_Options.chk_extractIcon.value <> vbUnchecked Then
               
               
              'Get Data
               Dim IconFileData As New StringReader
               IconFileData.Data = FileLoad(IconFileName.FileName)
               
              'Calc CRC
               Dim IconFileDataCrc As String
               IconFileDataCrc = ADLER32(IconFileData)
               
              'Check CRC List
               On Error Resume Next
               Dim FileName_Au3Standard_IconFile As String
               
               FileName_Au3Standard_IconFile = _
                  IsStandard_IconFile(IconFileDataCrc)
              
              'Delete File if in CRC List
               If FileName_Au3Standard_IconFile <> "" Then
                  FileDelete IconFileName.FileName
                  FrmMain.Log "   ^- IconFile deleted because it's standard AU3-icon: (" & IconFileDataCrc & ")  '" & FileName_Au3Standard_IconFile & "'"
               End If
   
            End If
         End If

End Sub

Public Function IsStandard_IconFile(uniqueItemID$) As Boolean
 
 ' init
   On Error Resume Next

   Au3Standard_IconFileCrcs.add "AutoIt_Main_v10_48x48_RGB-A.ico", "E1E3EB6E"
   Au3Standard_IconFileCrcs.add "AutoIt_StandardEXE.ico", "43C5DB27"
   Au3Standard_IconFileCrcs.add "AutoIt_StandardEXE_33142.ico", "AF7BB98F"
   
   
   Au3Standard_IconFileCrcs.add "AHK_L___________48x48_RGB-A.ico", "B186AA0D"
   Au3Standard_IconFileCrcs.add "AHK_Classic_____32x32_RGB__.ico", "FCC71A4B"



   On Error Resume Next
   Au3Standard_IconFileCrcs.add "", uniqueItemID
   IsUnique = Err = 0
End Function
