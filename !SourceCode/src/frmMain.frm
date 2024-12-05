VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "myExe2Aut >The Open Source Autoit/AutoHotKey Script Decompiler<"
   ClientHeight    =   9045
   ClientLeft      =   2595
   ClientTop       =   4935
   ClientWidth     =   9390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   9390
   Begin VB.CheckBox chk_NoDeTokenise 
      Caption         =   "Disable Detokeniser"
      Height          =   195
      Left            =   3375
      TabIndex        =   10
      ToolTipText     =   "Enable that when you decompile AutoItScripts lower than ver 3.1.6"
      Top             =   8805
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CheckBox Chk_force_old_script_type 
      Caption         =   "Force Old Script Type"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   8760
      Value           =   2  'Zwischenzustand
      Width           =   3255
   End
   Begin VB.CommandButton cmd_MD5_pwd_Lookup 
      Caption         =   "Lookup Passwordhash"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      ToolTipText     =   "Copies hash to clipboard and does an online query."
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox Chk_NormalSigScan 
      Caption         =   "Use 'normal' Au3_Signature to find start of script"
      Height          =   195
      Left            =   4920
      TabIndex        =   5
      Top             =   8580
      Width           =   3975
   End
   Begin VB.CheckBox Chk_RestoreIncludes 
      Caption         =   "Restore Includes"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   8595
      Value           =   1  'Aktiviert
      Width           =   1560
   End
   Begin VB.Timer Timer_OleDrag 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   480
   End
   Begin VB.CheckBox Chk_TmpFile 
      Caption         =   "Don't delete temp file (compressed script)"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   8565
      Width           =   3855
   End
   Begin VB.CommandButton Cmd_About 
      Caption         =   "About"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Txt_Filename 
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   1
      Text            =   "Drag the compiled AutoItExe / AutoHotKeyExe or obfucated script in here, or enter/paste path+filename."
      ToolTipText     =   "Drag in or type in da file"
      Top             =   120
      Width           =   9135
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double click to see more !"
      Top             =   6480
      Width           =   9135
   End
   Begin VB.ListBox List_Source 
      Appearance      =   0  '2D
      Height          =   5685
      ItemData        =   "frmMain.frx":628A
      Left            =   120
      List            =   "frmMain.frx":628C
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox Txt_Script 
      Height          =   5775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   9135
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'for mt_MT_Init to do a multiplation without 'overflow error'
'Private Declare Function Mul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal dw2 As Long, ByVal dw3 As Long, ByVal dw4 As Long) As Long

'Mersenne Twister
Private Declare Function MT_Init Lib "MT.DLL" (ByVal initSeed As Long) As Long
Private Declare Function MT_GetI8 Lib "MT.DLL" () As Long

'Private Declare Function Uncompress Lib "LZSS.DLL" (ByVal CompressedData$, ByVal CompressedDataSize&, ByVal OutData$, ByVal OutDataSize&) As Long
'Private Declare Function GetUncompressedSize Lib "LZSS.DLL" (ByVal CompressedData$, ByRef nUncompressedSize&) As Long

'Dim PE As New PE_info
Dim DeObfuscate As New ClsDeobfuscator

Dim FilePath_for_Txt$


'Const MD5_CRACKER_URL$ = "http://gdataonline.com/qkhash.php?mode=txt&hash="
Const MD5_CRACKER_URL$ = "http://www.md5cracker.de/crack.php?form=Cracken&md5="
'   http://www.milw0rm.com/cracker/info.php?'


Sub FL_verbose(text)
   log_verbose H32(File.Position - 1) & " -> " & text
End Sub

Sub log_verbose(TextLine$)
   If Chk_TmpFile.Value = vbChecked Then log TextLine
End Sub



Sub FL(text)
   log H32(File.Position - 1) & " -> " & text
End Sub

Public Sub log2(TextLine$)
'   log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Public Sub log(TextLine$)
On Error Resume Next
   List1.AddItem TextLine
'   List1.AddItem H32(GetTickCount) & vbTab & TextLine
 
 ' Process windows messages (=Refresh display)
   If (List1.ListCount < 10000) Or (Rnd < 0.1) Then
       ' Scroll to last item ; when there are more than &h7fff items there will be an overflow error
      Dim ListCount&
      List1.ListIndex = List1.ListCount - 1
      DoEvents
   End If
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Public Sub log_clear()
On Error Resume Next
   List1.Clear
End Sub

'
'Private Sub DeleteBackup()
'     FileRename FileName.Name & ".vEx", FileName.Name & ".del"
'     FileDelete FileName.Name & ".del"
'End Sub

'Working but not need anymore
'Private Sub mt_MT_Init(Key)
'
'
'   Dim Table
'   ReDim Table(624) '0x270
'   Dim v1&, v2&
'   Table(1) = Key
'
'   For i = 1 To UBound(Table) - 1
'     v1 = Table(i)
'     Debug.Assert i <> 5
' ' Cutoff + rotate last 30 bits
' ' v2 = v1 \ &H40000000 '2^30
'   If (v1 >= 0) Then
'      If (v1) < &H40000000 Then '2^30
'         v2 = 0
'      Else
'         v2 = 1
'      End If
'   Else
'      If v1 < &HC0000000 Then '2^30
'         v2 = 2
'      Else
'         v2 = 3
'      End If
'   End If
'
'   v1 = v1 Xor v2
'
'
''    v1 = v1 * 1812433253 '6C078965
'     v1 = Mul(v1, 0, 1812433253, 0) '6C078965
'
''     MsgBox v1
''     v2 = Int(v1 / &H40000000 / 4)
''     ' 9B2 252ADAA2            '2482 623565474
''     ' 9B2 252ADAA2- 9B2 00000000
''     v1 = v1 - (v2 * &H40000000 * 4)
'
'     v1 = v1 + i
'
'     Table(i + 1) = v1
'   Next
'
'End Sub


Private Function GetEncryptStrNew(LenEncryptionSeed&, StrEncryptionSeed, hFile As FileStream) As String
      Dim StrLen&
      StrLen = hFile.longValue
      StrLen = StrLen Xor LenEncryptionSeed
      
     'Double size on new type because of Unicode
      Dim StrLenToRead
      StrLenToRead = StrLen + StrLen
      
      GetEncryptStrNew = DeCryptNew(hFile.FixedString(StrLenToRead), StrEncryptionSeed + StrLen)
End Function

Private Function DeCryptNew(ByVal Data$, Key&)
   
   MsgBox _
      "Sorry Decryptions for new au3 Files is not implemented yet." & vbCrLf & _
      "(...and so you can't extract files whose source you don't have.)" & vbCrLf & _
      "" & vbCrLf & _
      "But you can test the TokenDecompiler that is already finished!" & vbCrLf & _
      "" & vbCrLf & _
      "1. add this line at the beginning of the your au3-sourcecode:" & vbCrLf & _
      "  FileInstall('>>>AUTOIT SCRIPT<<<', @ScriptDir & '\ExtractedSource.au3')" & vbCrLf & _
      "2. Compile it with the AutoIt3Compiler." & vbCrLf & _
      "3. Run the exe -> 'ExtractedSource.au3' get's extracted." & vbCrLf & _
      "4. Now open 'ExtractedSource.au3' with this decompiler." & vbCrLf & _
      "" & vbCrLf, _
      vbInformation, "Decryptions for new au3 Files is not implemented yet"
      
   Err.Raise ERR_NO_AUT_EXE + 100, , "Sorry Decryptions for new Au3 files is not implemented yet :("
End Function



Private Function GetEncryptStr(LenEncryptionSeed&, StrEncryptionSeed, hFile As FileStream) As String
      Dim StrLen&
      StrLen = hFile.longValue
      StrLen = StrLen Xor LenEncryptionSeed
            
      GetEncryptStr = DeCrypt(hFile.FixedString(StrLen), StrEncryptionSeed + StrLen)
End Function

Private Function DeCrypt(ByVal Data$, Key&)
   'Mersenne Twister (MT) to generate 'random' values
   'http://eprint.iacr.org/2005/165.pdf page 4
   'http://www.ecrypt.eu.org/stream/svn/viewcvs.cgi/ecrypt/trunk/submissions/cryptmt/cryptmt.c?rev=1&view=markup
   'http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/MT2002/emt19937ar.html
   
 ' Key->StartSeed for MT
   MT_Init (Key)
   
   Dim inBuff As New StringReader
   Dim OutBuff As New StringReader
   inBuff.Data = Data
   OutBuff.Data = Data

 ' Decrypt/Encrypt by  Xor Data from MT with inData
   Do While inBuff.EOS = False
      OutBuff.int8 = inBuff.int8 Xor (MT_GetI8 And &HFF)
      'DeCrypt = DeCrypt & Chr(inBuff.int8 Xor (MT_GetI8 And &HFF))
   Loop
   
   DeCrypt = OutBuff.Data
End Function


Private Sub Cmd_About_Click()
   FrmAbout.Show vbModal
End Sub

Private Sub cmd_MD5_pwd_Lookup_Click()
   Clipboard.Clear
   Clipboard.SetText MD5PassphraseHashText

   Dim hProc&
   hProc = ShellExecute(0, "open", MD5_CRACKER_URL$ & MD5PassphraseHashText, "", "", 1)

End Sub

Private Sub Form_Load()

   Listbox_SetHorizontalExtent List1, 6000
   FrmMain.Caption = FrmMain.Caption & " " & App.Major & "." & App.Minor & " build(" & App.Revision & ")"
   

   Me.Show
   
   
   
'   Dim BenchtimeA&, BenchtimeB&
'   BenchtimeA = GetTickCount
'
'   Dim dofusc As New ClsDeobfuscator
'   dofusc.DeObfuscate
'
'   BenchtimeB = GetTickCount
'   Debug.Print BenchtimeB - BenchtimeA
'
'Exit Sub
'Stop

 ' Commandlinesupport   :)
   If FileExists(Command$) Then
      Txt_Filename = Replace(Command$, """", "")
   Else
      Txt_Filename_Change
   End If
End Sub
   


Private Function FileExists(FileName) As Boolean
   On Error GoTo FileExists_err
   FileExists = FileLen(FileName)

FileExists_err:
End Function

Public Function GetLogdata$()
   Dim LogData As New clsStrCat
   LogData.Clear
   Dim i
   For i = 0 To List1.ListCount
      LogData.Concat (List1.List(i) & vbCrLf)
   Next
   
   GetLogdata = LogData.Value
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub List1_DblClick()
   frmLogView.txtlog = GetLogdata()
   frmLogView.Show
End Sub


Private Sub Timer_OleDrag_Timer()
   Timer_OleDrag.Enabled = False
   Txt_Filename = FilePath_for_Txt
End Sub


Private Sub Txt_Filename_Change()
   
   
   On Error GoTo Txt_Filename_err
   If FileExists(Txt_Filename) Then
      
      List1.Clear
      Txt_Script = ""
      
      FileName = Txt_Filename
      
      log String(80, "=")
      log "           -=  " & Me.Caption & "  =-"
      log String(80, "=")
'      log ""
         
      Decompile
      
      log "Testing for Scripts that were obfuscate by 'Jos van der Zande AutoIt3 Source Obfuscator v1.0.15 [July 1, 2007]' or 'EncodeIt 2.0'"

      For Each FileName In ExtractedFiles
'         If FileName.Ext Like "*.au*" Then
            On Error Resume Next
            log String(79, "=")
         DeToken
            If Err Then log "ERR: " & Err.Description

            On Error Resume Next
            log String(79, "=")
         DeObfuscate.DeObfuscate
            If Err Then log "ERR: " & Err.Description
            
          Select Case Err
          Case 0, ERR_NO_OBFUSCATE_AUT
            
            If Chk_RestoreIncludes = vbChecked Then SeperateIncludes
          
          Case Else
            log Err.Description
          End Select
          
 '        End If
      Next


Err.Clear
GoTo Txt_Filename_err

' Decompile Err Handler

      
      
DeToken:
      log String(79, "=")
      DeToken

DeObfuscate:
      log String(79, "=")
      DeObfuscate.DeObfuscate
      
Txt_Filename_err:
  ' Note: Resume is necessary to reenable Errorhandler
  '       Else the VB-standard Handler will catch the error -> Exit Programm
    Select Case Err
    Case 0
    
    Case ERR_NO_AUT_EXE
       log Err.Description
       Resume DeToken
    
    Case NO_AUT_DE_TOKEN_FILE
       log Err.Description
       Resume DeObfuscate
    
    Case ERR_NO_OBFUSCATE_AUT
       log Err.Description
       Resume Txt_Filename_err
       
       
    Case Else
       log Err.Description
       Resume Txt_Filename_err
    End Select
   
    
    'Save Log Data
    On Error Resume Next
    
'    If UBound(ExtractedFiles) < 0 Then
    FileName = ExtractedFiles(1).FileName
    FileName.NameWithExt = "_myExeToAut.log"
    
    log ""
    log "Saving Logdata to : " & FileName.FileName
    File.Create FileName.FileName, True
    File.FixedString(-1) = GetLogdata
    File.CloseFile
   
   End If
   
End Sub


Private Function OpenFile(Target_FileName As ClsFilename) As Boolean
   
   On Error GoTo Scanfile_err
   log "------------------------------------------------"

   log Space(4) & Target_FileName.NameWithExt

   File.Create Target_FileName.mvarFileName, Readonly:=True
   
   Me.Show

Err.Clear
Scanfile_err:
Select Case Err
   Case 0

   Case Else
      log "-->ERR: " & Err.Description

End Select
   
End Function


Private Sub Txt_Filename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   FilePath_for_Txt = Data.Files(1)
   Timer_OleDrag.Enabled = True
End Sub


Private Sub Txt_Script_Change()
  If Len(Txt_Script) >= 65535 Then
      Txt_Script.ToolTipText = "Notice: Display limited to 65535 Bytes. File is bigger."
  Else
      Txt_Script.ToolTipText = ""
  End If
End Sub

Private Sub Txt_Script_KeyDown(KeyCode As Integer, Shift As Integer)
   Cancel = KeyCode <> vbKeySpace
End Sub
