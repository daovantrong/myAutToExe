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


'for mt_init_genrand to do a multiplation without 'overflow error'
'Private Declare Function Mul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal dw2 As Long, ByVal dw3 As Long, ByVal dw4 As Long) As Long

'Mersenne Twister
Private Declare Function init_genrand Lib "MT.DLL" (ByVal initSeed As Long) As Long
Private Declare Function genrand_int31 Lib "MT.DLL" () As Long

'Private Declare Function Uncompress Lib "LZSS.DLL" (ByVal CompressedData$, ByVal CompressedDataSize&, ByVal OutData$, ByVal OutDataSize&) As Long
'Private Declare Function GetUncompressedSize Lib "LZSS.DLL" (ByVal CompressedData$, ByRef nUncompressedSize&) As Long


Dim AU3Sig As New StringReader, AU3SigSize&
'Dim PE As New PE_info
Dim DeObfuscate As New ClsDeobfuscator

Dim FilePath_for_Txt$


Public MD5PassphraseHashText$
Const MD5_HASH_EMPTY_STRING$ = "D41D8CD98F00B204E9800998ECF8427E"

'Const MD5_CRACKER_URL$ = "http://gdataonline.com/qkhash.php?mode=txt&hash="
Const MD5_CRACKER_URL$ = "http://www.md5cracker.de/crack.php?form=Cracken&md5="
'   http://www.milw0rm.com/cracker/info.php?'

Const Script_KEY& = &HAAAAAAAA
Dim bIsProbablyOldScript As Boolean

Dim bIsNewScriptType As Boolean


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
'Private Sub mt_init_genrand(Key)
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
   init_genrand (Key)
   
   Dim inBuff As New StringReader
   Dim OutBuff As New StringReader
   inBuff.Data = Data
   OutBuff.Data = Data

 ' Decrypt/Encrypt by  Xor Data from MT with inData
   Do While inBuff.EOS = False
      OutBuff.int8 = inBuff.int8 Xor (genrand_int31 And &HFF)
      'DeCrypt = DeCrypt & Chr(inBuff.int8 Xor (genrand_int31 And &HFF))
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
   FrmMain.Caption = FrmMain.Caption & " " & App.Major & "." & App.Minor & " [alpha]" '  build(" & App.Revision & ")"
   
   AU3Sig = HexvaluesToString("A3 48 4B BE 98 6C 4A A9 99 4C 53 0A 86 D6 48 7D") ' & "AU3!"  "£HK¾˜lJ©™LS.†ÖH}"


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
      Txt_Filename = Command$
   Else
      Txt_Filename_Change
   End If
End Sub
   

Private Function TestForV3_26() As Boolean
   With File
      .Position = .Length - 4 - 4
      TestForV3_26 = .FixedString(8) = "AU3!EA06"
   End With
End Function


Private Function TestForV3_2() As Boolean
   With File
      .Position = .Length - 4 - 4
      TestForV3_2 = .FixedString(8) = "AU3!EA05"
   End With
End Function

Private Function TestForV3_1() As Boolean
   With File
      .Position = .Length - 4 - 4 - 4

      Dim Script_End&
      Script_End = .longValue Xor Script_KEY
'      FL "Script_End: " & H32(Script_End) & "  (XORed with 0x" & H32(Script_KEY)

      Dim Script_Start&
      Script_Start = .longValue Xor Script_KEY
'      FL "Script_Start: " & H32(Script_Start) & "  (XORed with 0x" & H32(Script_KEY)

      Dim Script_CRC&
      Script_CRC = .longValue Xor Script_KEY
'      FL "Script_CRC: " & H32(Script_CRC) & "  (XORed with 0x" & H32(Script_KEY)
      
      If (Script_Start < Script_End) Then
         If RangeCheck(Script_Start, .Length, &H1001) And RangeCheck(Script_End, .Length, &H1001) Then
            bIsProbablyOldScript = True
            Dim Script_Data As New StringReader
            .Position = 0
            Script_Data = .FixedString(Script_End)
      
            Dim Script_CRC_Calculated&
            Script_CRC_Calculated = "&h" & ADLER32(Script_Data)
'            log "Script_CRC_Calculated: " & H32(Script_CRC_Calculated)
            
            TestForV3_1 = Script_CRC_Calculated = Script_CRC
            If TestForV3_1 Then
                  .Position = Script_Start
                  Dim Script_lengh&
                  Script_lengh = .longValue Xor 44460 '&HADAC
               FL "skipping " & H16(Script_lengh) & " byte of random fill data"
                  Dim FillData As New StringReader
                  FillData = .FixedString(Script_lengh)
               log ValuesToHexString(FillData)
'               log FillData.mvardata
            
            End If   'CRC_test
         End If   'RangeCheck
      End If   'Script_Start < Script_End
   
   End With
End Function

Private Function TestForV3_0() As Boolean
   
   With File
      .Position = .Length - 4 - 4
      
      ' ==== Handler for old Scripts ====
      Dim Script_Start&
      Script_Start = .longValue
'     FL "Script_Start: " & H32(Script_Start)
      
      Dim Script_CRC&
      Script_CRC = .longValue Xor Script_KEY&
'     FL "Script_CRC: " & H32(Script_CRC) & "  (XORed with 0x" & H32(Script_KEY)
      
      Dim Script_End&
      Script_End = .Length - 4
'     log " ===> Script_End: " & H32(Script_End)
      
      If RangeCheck(Script_Start, .Length, &H1001) Then
         

         bIsProbablyOldScript = True
       
       ' Check CRC32 to be sure that it's in the right format
         CRCInit (79764919) '&H4C11DB7)
      
         Dim Script_CRC_Calculated&
         .Position = 0
         Script_CRC_Calculated = CRC32(StrConv(.FixedString(Script_End), vbFromUnicode))
 '       log "Script_CRC_Calculated: " & H32(Script_CRC_Calculated)
      
         TestForV3_0 = Script_CRC_Calculated = Script_CRC
         If TestForV3_0 Then
            .Position = Script_Start
            
            Dim modified_AU3_Signature As New StringReader
            modified_AU3_Signature = .FixedString(Len(AU3Sig))
            log IIf(modified_AU3_Signature <> AU3Sig, "Modified ", "") & "AU3_Signature: " & ValuesToHexString(modified_AU3_Signature) & "  " & modified_AU3_Signature
         
            Exit Function
         End If
      End If
   End With
   
End Function



Private Sub FindStartOfScriptAlternative()
   With File
      
      bIsProbablyOldScript = Chk_force_old_script_type.Value = vbChecked
      If FrmMain.Chk_force_old_script_type.Value = CheckBoxConstants.vbGrayed Then
      
         If TestForV3_26 Then
            log "Script Type 3.2.5+ found."
            bIsNewScriptType = True
         
         ElseIf TestForV3_2 Then
            log "Script Type 3.2 found."
         
         ElseIf TestForV3_1 Then
            log "Script Type 3.1 found."
            'log "Script_Start is invalid trying something else..."
            Exit Sub
         
         ElseIf TestForV3_0 Then
            log "Script Type 3.0 found."
            Exit Sub
            
         End If 'of New ScriptType
      
      End If 'of Force ScriptType
      
      .Position = 0
          
' === Alternativ Scan ===
          
          
    '  Signature not found - try alternative search...
      'Err.Raise vbObjectError + 41022, , "Error: The executable file is not recognised as a compiled AutoIt script."
log "AlternativeSigScan for 'FILE'-signature in au3-body..."
     
     'FF 6D B0 CE                    ÿm°Î
      .FindString HexvaluesToString("FF 6D B0 CE")
      If .Position = 0 Then
         
         'Not Found - Search for signature of new Aut3Script
         .FindString HexvaluesToString("6B 43 CA 52")
         If .Position = 0 Then
            '...Error Exit Sub
            Err.Raise ERR_NO_AUT_EXE, , "'FILE'-signature not found."
            Exit Sub
         Else
            '...Finally found :)
            If bIsNewScriptType = False Then
               log "Modified Script Type 3.2.5+ found."
               bIsNewScriptType = True
            End If
         End If

      End If
   
   '  FILE-Signature found ! Now move back to script start...
      
      Dim FilePos_BodyStart&
      FilePos_BodyStart = .Position
      
   '  Determinate if it's an AUTOHOTKEY or AUTOIT SCRIPT ...
      .Move 4 ' skip over string 'FILE'
   
      Dim Au3SrcFile_FileInst As Boolean
      Dim SrcFile_FileInst$
      If bIsNewScriptType Then
         SrcFile_FileInst = GetEncryptStrNew(44476, 45887, File)  '0xADBC B33F
      Else
         SrcFile_FileInst = GetEncryptStr(10684, 41566, File)  '0x29BC A25E
      End If
      

FL "SrcFile_FileInst: " & SrcFile_FileInst
      If SrcFile_FileInst = ">AUTOIT UNICODE SCRIPT<" Then
         Au3SrcFile_FileInst = True
         
      ElseIf SrcFile_FileInst = ">AUTOIT SCRIPT<" Then
       ' use AHK_Mode for old scripts
         Au3SrcFile_FileInst = Not (bIsProbablyOldScript)
         
      ElseIf SrcFile_FileInst = ">AUTOHOTKEY SCRIPT<" Then
         Au3SrcFile_FileInst = False
         
      ElseIf SrcFile_FileInst = ">AHK WITH ICON<" Then
         Au3SrcFile_FileInst = False
      Else
log "WARNING: unknown SrcFile_FileInst!"
         Au3SrcFile_FileInst = vbYes = MsgBox("Press YES to process this as an AUTOIT SCRIPT." & vbCrLf & "Press NO to process this as an AUTOHOTKEY SCRIPT.", vbQuestion + vbYesNo, "Unknown SrcFile_FileInst : " & SrcFile_FileInst)
      End If

    ' Now seek back to script start position....
log "Seeking back to script start position..."
      .Position = FilePos_BodyStart
      If Au3SrcFile_FileInst Then
       ' MD5PasswordHash
         .Move -&H10
       
       ' "EA05"
         .Move -4
       
       ' SubType ["AU3!"]
         .Move -4
       
       ' AU3Signature ["£HK¾..."]
         .Move -Len(AU3Sig)
         
      Else

       '  working but to cryptic
'                  .Move -4
'                  Do Until (4 + .Position + (.longValue Xor 64193)) = FilePos_BodyStart
'                     .Move (-1 - 4)
'                  Loop
'                  .Move -4
        
         
       ' Determinating length of au3-password
       ' Expected format:
       '   <DWORD>Len  <String>Password   <Offset>FilePos_BodyStart...
         Const AU3_MAX_PASSWORDLEN = 263
         Do
          ' Get Length
            Dim PasswordLen As Double
            .Move -4
            PasswordLen = .longValue Xor 64193

          ' If Current File Position + Length is FilePos_BodyStart...
            If .Position + PasswordLen = FilePos_BodyStart Then
              '...it's the correct length; so seek back a Dword[4byte]
               .Move -4
              'Exit Loop
               Exit Do
            ElseIf (FilePos_BodyStart - .Position) >= AU3_MAX_PASSWORDLEN Then
               Err.Raise vbObjectError, , "Determinating length of au3-password failed (length >= " & AU3_MAX_PASSWORDLEN & ")"
            End If

           ' Seek to next position to try
            .Move -1
         Loop While True

       ' SubType
         .Move -1
       
       ' AU3Signature
         .Move -Len(AU3Sig)
      End If

      
      Dim modified_AU3_Signature As New StringReader
      modified_AU3_Signature = .FixedString(Len(AU3Sig))
      
      log IIf(modified_AU3_Signature <> AU3Sig, "Modified ", "") & "AU3_Signature: " & ValuesToHexString(modified_AU3_Signature) & "  " & modified_AU3_Signature

'            log "Not found trying heuristic search..."
'            PE.Create
'            Dim LastSection As PE_Section
'            With LastSection
'               LastSection = PE_Header.Sections(PE_Header.NumberofSections - 1)
'               log "LastSection in PE_Header is: " & szNullCut(.SectionName) & " at: " & H32(.PointertoRawData) & " Size: " & H32(.RawDataSize)
'
'               Dim ScriptStart
'
'            End With
   End With

End Sub

'// FindStartOfScript - Locate startoffset of scriptdata
'// Note:  Tries to find AutoIt3 start signature
'//        if it fails search for 'FILE' (encrypted)
Private Sub FindStartOfScript()
   
   With File
       
      ' ===> Find Script Signature in FileData  (and place FileReadPointer behind it)
log "Scanning for AutoIt Signature:" & ValuesToHexString(AU3Sig) & "   " & AU3Sig
        .FindString AU3Sig.Data '"AU3!"
        If .Position <> 0 Then
           .Move Len(AU3Sig)
        Else
        
         '  Signature not found - try alternative search...
           'Err.Raise vbObjectError + 41022, , "Error: The executable file is not recognised as a compiled AutoIt script."
log "...not found."
          
          'FF 6D B0 CE                    ÿm°Î
           FindStartOfScriptAlternative
        End If
         
   End With

End Sub

Private Function FormatFileTime(TimeStamp As FILETIME) As String
   Dim SysTime As SYSTEMTIME
   With SysTime
      FileTimeToSystemTime TimeStamp, SysTime
      FormatFileTime = Format(.wDay & "." & .wMonth & "." & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond, "dd.mm.yyyy hh:mm:ss") & " [" & .wMilliseconds & "]"
   End With
End Function

Private Sub UserPassWordCheck(MD5PassphraseHashText$, bIsClearTextPwd As Boolean)
   #If DoUserPassWordCheck Then
'////////////////////////////////////////////////////////////////////
'//
'//  A t t e n t i o n , W a r n i n g , A t t e n t i o n , W a r n i n g
'//                P r o t e c t e d  C o d e
'// It is strictly FORBIDDEN to REMOVE or modify the following code:
         
         Dim md5 As ClsMD5
         Set md5 = New ClsMD5
         Dim userPassword$, userPassword_Hash$, scriptPassword_Hash$
         scriptPassword_Hash = LCase(MD5PassphraseHashText)
         Do
            userPassword = InputBox("Please Password:", "Script File is Password Protected", "Sorry but for legal reason you must enter a valid password to continue.")
            If userPassword = "" Then Err.Raise vbObjectError, , "Stopped because user didn't entered a valid password!"
            
              'According to type test clearTextPWD or Hash
               If bIsClearTextPwd Then
                  userPassword_Hash = userPassword
               Else
                  userPassword_Hash = md5.md5(userPassword)
               End If
         
         Loop Until userPassword_Hash = scriptPassword_Hash

'//                   E N D  O F  'untouchable code'               //
'////////////////////////////////////////////////////////////////////
#End If


End Sub



'////////////////////////////////////////////////////
'/// Decompile - Decompiles .exe[->File] to .au3 or .ahk
'//
'//  Notes:
'//   Not indented lines are for log purpose only (and not so important)
Private Sub Decompile()

'log "---------------------------------------------------------"
   
  'Clear ExtractedFiles
   Set ExtractedFiles = New Collection
   
   With File
    
      log "Unpacking: " & FileName.FileName
      .Create FileName.FileName, False, False, True
      .Position = 0
      
     'Find Start of Script and Quit this function with runtime error if search fails
      If Chk_NormalSigScan = vbChecked Then
         FindStartOfScript
      Else
         FindStartOfScriptAlternative
      End If
      
      
      
    
    ' ===> Check if it's Old or New AutoIt Script
      Dim SubType As New StringReader:   SubType.DisableAutoMove = True
      SubType = .FixedString(4)
      FL "SubType: 0x" & H8(SubType.int8) & "  " & SubType.mvardata
      
      Dim bIsOldScript As Boolean
      If SubType.Data = "AU3!" Then
         bIsOldScript = False
    
    ' the offical AutoHotkey Script Decompiler checks this to be '3'
      ElseIf SubType.int8 = 3 Then
         bIsOldScript = True
      
      ElseIf SubType.int8 = 4 Then
         bIsOldScript = True
      
      Else
         'err.Raise vbObjectError,,"Unexpected Script subtype"
         FL "Unexpected Script subtype: " & "0x" & H32(SubType.int32) & " " & SubType.Data
      End If

      log "~ Note:  The following offset values are were the data ends (and not were it starts) ~"

    
    ' ===> Get Script Password
      Dim MD5PassphraseHash As New StringReader
      If bIsOldScript Then
       ' Old AutoIT Script if branch...
       ' Move three bytes back since SubType is only 1 Byte but befroe we read 4 byte
         .Move -3
         MD5PassphraseHash = GetEncryptStr(64193, 50130, File) '&HFAC1, &HC3D2
         MD5PassphraseHashText = MD5PassphraseHash
      
      Else
       ' New AutoIT script if branch...
         
log "Note:  The offset values are were the data ENDs (and not as usual where it BEGINs)"
         
         Dim Type2$
         Type2 = .FixedString(4)

         If Type2 <> "EA05" Then
FL "Type2 = " & Type2 & "  Normally you would get 'Error: Unsupported Version of AutoIt script.' here"
         End If
         
         FL "AutoIt Script Found.  - Type2 = " & Type2
         
         'Err.Raise vbObjectError + 41022, , "Error: Unsupported Version of AutoIt script."

      
         ' GetPassword Hash from with later the key to decrypt the script is calculated
           MD5PassphraseHash = .FixedString(&H10)
           MD5PassphraseHashText = ValuesToHexString(MD5PassphraseHash, "")
           
           Dim IsHashForEmptyPassword As Boolean
           IsHashForEmptyPassword = MD5PassphraseHashText = MD5_HASH_EMPTY_STRING$
           If IsHashForEmptyPassword Then MD5PassphraseHashText = ""
            
      End If
      
      
     '==> Ask User For Password
      If (MD5PassphraseHashText = "") Then
         log "Script has no password (MD5PassphraseHash for password """" )"

      Else
         log "Script is password protected!"

         #If DoUserPassWordCheck Then
         '////////////////////////////////////////////////////////////////////
         '//
         '//  A t t e n t i o n , W a r n i n g , A t t e n t i o n , W a r n i n g
         '//                P r o t e c t e d  C o d e
         '// It is strictly FORBIDDEN to REMOVE or modify the following code:
                  
          UserPassWordCheck MD5PassphraseHashText$, bIsOldScript
          
         
         '//                   E N D  O F  'untouchable code'               //
         '////////////////////////////////////////////////////////////////////
         #End If
      
      End If

      FL "Password/MD5PassphraseHash: " & ValuesToHexString(MD5PassphraseHash, "")
      log Space(8 + 4) & MD5PassphraseHash.Data
      
      cmd_MD5_pwd_Lookup.Visible = (IsHashForEmptyPassword = False) And (bIsOldScript = False)


    
    ' ==> Prepare decryption of script...
    ' A 32 bit checksumvalue over all bytes from the MD5PassphraseHash is the decryptionkey
      Dim MD5PassphraseHash_ByteSum&
      MD5PassphraseHash_ByteSum = 0
      
      MD5PassphraseHash.EOS = False
      Do Until MD5PassphraseHash.EOS

         If bIsOldScript Then
          ' For AHK scripts use signed int8 to also compute äöü correct
            MD5PassphraseHash_ByteSum = MD5PassphraseHash_ByteSum + MD5PassphraseHash.int8Sig
         Else
          ' For new MD5 scripts use unsigned int8 to compute
            MD5PassphraseHash_ByteSum = MD5PassphraseHash_ByteSum + MD5PassphraseHash.int8
         End If
         
'         Debug.Print MD5PassphraseHash.Position, H32(MD5PassphraseHash_ByteSum)
      Loop
      log "MD5PassphraseHash_ByteSum: " & H32(MD5PassphraseHash_ByteSum) & "  '+ 22AF' => decryption key!"

   
   
   
      log "------------ Processing Body -------------"
      Dim FileCount&
      For FileCount = 1 To &H7FFFFFF

      '===> read various Header data
         Dim ResType$
         ResType = DeCrypt(.FixedString(4), 5882) '000016FA
      If ResType <> "FILE" Then
         log "Processing Finished!"
       ' No valid FILE Marker so seek back
         .Move -4
         Exit For
      End If
   
         log "=== > Processing FILE: #" & FileCount
         FL "ResType: " & ResType
      
      
         Dim SrcFile_FileInst$
         SrcFile_FileInst = GetEncryptStr(10684, 41566, File) '0x29BC A25E
         FL "SrcFile_FileInst: " & SrcFile_FileInst
      
         Dim CompiledPathName As New ClsFilename
         CompiledPathName = GetEncryptStr(10668, 62046, File) '29AC  F25E
         FL "CompiledPathName: " & CompiledPathName
         
         Dim bIsAHK_Script As Boolean
         If SrcFile_FileInst = ">AUTOIT UNICODE SCRIPT<" Then
            bIsAHK_Script = False
            
         ElseIf SrcFile_FileInst = ">AUTOIT SCRIPT<" Then
            bIsAHK_Script = False
            
         ElseIf SrcFile_FileInst = ">AUTOHOTKEY SCRIPT<" Then
            bIsAHK_Script = True
            
         ElseIf SrcFile_FileInst = ">AHK WITH ICON<" Then
            bIsAHK_Script = True
            
         Else
            log Space(8 + 4) & "WARNING: unknown SrcFile_FileInst!"
         End If
            
            
      ' ==> Is script compressed
         Dim IsCompressed&
         IsCompressed = .ByteValue
         FL "IsCompressed: " & CBool(IsCompressed)
        
      ' ==> Get size of compressed script data
        Dim ScriptSize&
        ScriptSize = .longValue
        ScriptSize = ScriptSize Xor 17834 '45AA
        FL "ScriptSize Compressed: " & H32(ScriptSize) & "  Decimal:" & ScriptSize
   
        Dim SizeUncompressed&
        SizeUncompressed = .longValue Xor 17834 '45AA
        FL "ScriptSize UnCompressed(not used by aut2exe so far): " & H32(SizeUncompressed) & "  Decimal:" & SizeUncompressed
        
         If bIsOldScript = False Then
         ' ==> CRC32 value of uncompressed script data
            Dim ScriptData_CRC&
            ScriptData_CRC = .longValue Xor 50130 '0C3D2

'            If &H1C00000 = (ScriptData_CRC And &HFFF00000) Then
'               log "Rewinded due to suspiciously CRC that is probably a date"
'               .Move -4
''                 bIsOldScript = True
'            End If

            FL "ADLER32 CRC of unencrypted script data: " & H32(ScriptData_CRC)
         End If
         
         Dim pCreationTime As FILETIME, pLastWrite As FILETIME
         pCreationTime.dwHighDateTime = .longValue
         pCreationTime.dwLowDateTime = .longValue
         pLastWrite.dwHighDateTime = .longValue
         pLastWrite.dwLowDateTime = .longValue
         FL "FileTime (number of 100-nanosecond intervals since January 1, 1601) "
         log Space(4) & "pCreationTime:  " & H32(pCreationTime.dwHighDateTime) & H32(pCreationTime.dwLowDateTime) & "  " & FormatFileTime(pCreationTime)
         log Space(4) & "pLastWrite   :  " & H32(pLastWrite.dwHighDateTime) & H32(pLastWrite.dwLowDateTime) & "  " & FormatFileTime(pLastWrite)
       
        '==> Read encrypted script data
         FL "Begin of script data"
         
         Dim ScriptData As New StringReader
         ScriptData = .FixedString(ScriptSize)
   
         
   ' ~~~ Process decrypted scriptdata ~~~
         log "Decrypting script data..."
       
         'MD5PassphraseHash_ByteSum = MD5PassphraseHash_ByteSum + 8879 '&H22AF
         init_genrand (MD5PassphraseHash_ByteSum + 8879) '&H22AF
      
         With ScriptData
            
           ' ==> Decrypt scriptdata

'            Dim Benchmark&
'            Benchmark = GetTickCount
            Dim StrCharPos&, tmpBuff$
            tmpBuff = StrConv(.mvardata, vbFromUnicode)
            
            For StrCharPos = 1 To Len(.mvardata)
               
               MidB$(tmpBuff, StrCharPos, 1) = ChrB$(AscB(MidB$(tmpBuff, StrCharPos, 1)) _
                     Xor _
                     (genrand_int31& And &HFF))
               
               If 0 = (StrCharPos Mod &H8000) Then DoEvents
               
            Next
            
            .mvardata = StrConv(tmpBuff, vbUnicode)

'            Debug.Print GetTickCount - a 'Benchmark:4453 (6171 mid version)


'Note: This Version is 4x slower
'            Dim Benchmark&
'            Benchmark = GetTickCount


'            .EOS = False
'            .DisableAutoMove = True
'            Do Until .EOS
'               .int8 = .int8 Xor (genrand_int31 And &HFF)
'               .Move 1
'            Loop

'            Debug.Print GetTickCount - Benchmark 'Benchmark:24063
            
          ' ==> Create output fileName
            Dim OutFileName As New ClsFilename
          ' initialise with ScriptPath
            OutFileName = File.FileName
            
            
            If (CompiledPathName.Name Like "*>*") Or (CompiledPathName.Ext Like "*tmp*") Then
               
               OutFileName.Ext = IIf(bIsAHK_Script, ".ahk", ".au3")
               If ExtractedFiles.Count > 0 Then
                  OutFileName.Name = OutFileName.Name & "_" & ExtractedFiles.Count
               End If
           
            Else
               
               'if its an absolute path like "C:\Documents and Settings\EnCodeItInfo\Restart_EnCoded1.au3"
               'Just use the filename and don't create subdirs
               If InStr(SrcFile_FileInst, ":") Then
                  OutFileName.NameWithExt = CompiledPathName.Dir & CompiledPathName.NameWithExt
               Else
               ' Set Dir
                 OutFileName.NameWithExt = SrcFile_FileInst
               End If
               
               ' create Dir if it doesn't exists
               OutFileName.MakePath
               
            End If
      
      
      
          ' Do ADLER32 CRCTest for AutoIT Scripts
            If bIsOldScript = False Then
      
               log "Calculating ADLER32 checksum from decrypted scriptdata"
               
               Dim ScriptData_CRC_Calculated&
               ScriptData_CRC_Calculated = "&h" & ADLER32(ScriptData)
               If (H32(ScriptData_CRC)) = (ScriptData_CRC_Calculated) Then
                  log "   OK."
               Else
                  log "   FAILED!"

               End If
               log "   Calculate ADLER32: " & ScriptData_CRC_Calculated
               log "   CRC from script  : " & H32(ScriptData_CRC)
            End If
            
      
            If IsCompressed Then
            
              ' ==> Decompress Script
               .EOS = False
               .DisableAutoMove = False
               
               Dim LZSS_Signature$
               LZSS_Signature = .FixedString(4)
               log "JB LZSS Signature:" & LZSS_Signature
      
               If LZSS_Signature = "EA04" Then
                  Dim LZSS_Signature_new$
                  LZSS_Signature_new = "EA05"
                  log "Forcing/overwrite signature to '" & LZSS_Signature_new
                  .Move -4
                  .FixedString(4) = LZSS_Signature_new
               Else
      
                  ' Check signature of compressed data
                  Dim ExpectedSignature$
                  ExpectedSignature = IIf(bIsOldScript, "JB01", "EA05")
                  If LZSS_Signature <> ExpectedSignature Then
                  log "WARNING: Normally signature is '" & ExpectedSignature & "' - possible reasons: 'modified' AutToExe, decryption failure, new version..."
                     'If signature looks weird probably decryption fail and this is of no use

                     Do
                        LZSS_Signature_new = InputBox("Current value is '" & LZSS_Signature & "'" & vbCrLf & "Valid values are 'EA05' and 'JB01'." & vbCrLf & "Note: If current value looks weird probably decryption fail and so data might be garbage." & vbCrLf & vbCrLf & "Since this is an Auto" & IIf(bIsOldScript, "HotKey", "IT") & " Script the recommanded value is '" & ExpectedSignature & "'" & vbCrLf & vbCrLf & "Press >OK< to change this value or" & vbCrLf & ">Cancel< to keep this it unchanged.", "Compression signature is invalid !", ExpectedSignature)
                     Loop Until (Len(LZSS_Signature_new) = 4) Or (Len(LZSS_Signature_new) = 0)
                     If (Len(LZSS_Signature_new) = 4) Then
   '                  If vbYes = MsgBox("Do you want to force it to : " & ExpectedSignature & " so this stream can be decompressed?" & vbCrLf & vbCrLf & "Note: If signature looks weird probably decryption fail and this is of no use", vbYesNo + vbDefaultButton1 + vbExclamation, "LZSS_Signature of decrypted data is '" & LZSS_Signature & "'") Then
                        log "Forcing/overwrite signature to '" & LZSS_Signature_new
                        .Move -4
                        .FixedString(4) = LZSS_Signature_new 'InputBox("You may change the signature here", "", ExpectedSignature)
                     End If
                  End If
                  
               End If
         
               
      '         Dim SizeUncompressed& ', w1&, w2&
      '         SizeUncompressed = .int8
      '         SizeUncompressed = .int8 Or (SizeUncompressed * &H100)
      '         SizeUncompressed = .int8 Or (SizeUncompressed * &H100)
      '         SizeUncompressed = .int8 Or (SizeUncompressed * &H100)
      
      '         RetVal = GetUncompressedSize(.data, SizeUncompressed)
      '         If RetVal <> 0 Then Err.Raise 0, , "GetUncompressedSize() failed"
      'log "Uncompressed script size:" & H32(SizeUncompressed)
      
      '
             ' save compressed script data to *.tmp in current Dir
             '    if 'Create DebugFile' was not checked it will be delete on close
               Dim tmpFile As New FileStream
               With tmpFile
                  .Create OutFileName.Path & OutFileName.Name & ".tmp", True, Chk_TmpFile.Value = vbUnchecked, False
                  .Data = ScriptData.Data
                   log "Compressed scriptdata written to " & .FileName
         
                  
                  Dim RetVal&
                ' About LZSS see: http://de.wikipedia.org/wiki/Lempel-Ziv-Storer-Szymanski-Algorithmus
                  log "Expanding script data..."
         
         '         Dim tmpstr$
         '         tmpstr = Space(SizeUncompressed)
         '         RetVal = Uncompress(.data, .Length, tmpstr, SizeUncompressed)
                 ' write decompressed Data back to stream
         '         .data = tmpstr
                 
              log "Saving script to """ & OutFileName.NameWithExt & """ at " & OutFileName.Path
                  
                ' Run "LZSS.exe -d *.debug *.au3" to extract the script (...and wait for it execution to finish)
                  ShellEx App.Path & "\" & "lzss.exe", _
                        "-d """ & .FileName & """ """ & OutFileName & """"
                  
                  .CloseFile
               End With
               
      
             ' Read data from new script file
               Dim outFile As New FileStream
               outFile.Create OutFileName.FileName, False, False, False
               outFile.Position = 0
               .Data = outFile.FixedString(-1)
     
            
            Else
            '... data was not compress, so just save the script data
               log "Saving script to """ & OutFileName.NameWithExt & """ at " & OutFileName.Path
            
               outFile.Create OutFileName.FileName, True, False, False
               outFile.Data = .Data
            
            End If
            
          ' Add extracted FileName to global ExtractedFiles List
          ' Clear global list with extracted files
            Dim newFileName As ClsFilename
            Set newFileName = New ClsFilename

            newFileName.FileName = OutFileName.FileName
            ExtractedFiles.Add newFileName

            log "Setting Creation and LastWrite time"
            Err.Clear
            RetVal = SetFileTime(outFile.hFile, pCreationTime, 0, pLastWrite)
            If RetVal = 0 Then
               RetVal = Err.LastDllError
               log "LastDllError: " & RetVal
            End If
            
          
            outFile.CloseFile
            
            
          ' Show scriptdata
            If SrcFile_FileInst = ">AUTOIT UNICODE SCRIPT<" Then
               log "Convert from FromUnicode to Accii and write data in textbox"
               Txt_Script = StrConv(.Data, vbFromUnicode)
            Else
               log "Write data in textbox"
               Txt_Script = .Data
            End If
   
         End With 'ScriptData
         
'        'Run Tidy on script
'         Dim tmpob As ClsFilename
'         Set tmpob = FileName
'         Set FileName = OutFileName
'            SaveScriptData Txt_Script
'         Set FileName = tmpob
         
         
         log String(79, "-")
      
   Next
   
      
   ' if there are more than 8 bytes overlay save them to *.overlay file
   ' For clearity reason I pasted overlay logging to a seperated function
   Decompile_Log_ProcessOverlay .Length - .Position, .FixedString(-1), bIsOldScript
   ' ==> Exe Processing finished
   .CloseFile
   
   log String(79, "=")

End With

End Sub
Private Sub Decompile_Log_ProcessOverlay(overlaySize&, overlaybytes$, bIsOldScript As Boolean)
   
   With File
      
FL "End of script data " & "  FileLen: " & H32(.Length) & "  => Overlay: " & H32(overlaySize)

Dim tmp As New StringReader
tmp = Left(overlaybytes, &H20)
log "overlaybytes: " & ValuesToHexString(tmp) & "  " & overlaybytes
      If overlaySize > (IIf(bIsOldScript, 3, 2) * 4) Then
         Dim ovlFile As New FileStream
         With ovlFile
            .Create File.FileName & ".overlay", True, False, False
            log ">>>ATTENTION: There are more overlay data than usual <<<"
            log "saving overlaydata to: " & .FileName
            
            .Data = overlaybytes
            
            .CloseFile
         End With
      
      End If
   
   End With

End Sub

'Private Sub TestCRC()
'
'End Sub
'

'Private Sub UncompressLZSS(InData, DeComp As StringReader)
'
'
'     'BitStreamRead.data=InData
'
''     Dim DeComp As New StringReader
'     Dim BitsLeft
'     Do While BitsLeft 'BitStreamRead.BitsLeft
'        If GetBits(1) = 0 Then
'        ' literal
'           DeComp.int8 = GetBits(8)
'        Else
'        '  Tupel
'           Dim RewindBytes&, size&
'           RewindBytes = GetBits(15)
'
'         ' Handle Size
'           Dim SizePlus
'           size = GetBits(2): SizePlus = &H0
'           If size = 3 Then
'
'              size = GetBits(3): SizePlus = &H3
'              If size = 7 Then
'
'                 size = GetBits(5): SizePlus = &HA
'                 If size = &H1F Then
'
'                    size = GetBits(8): SizePlus = &H29
'                    If size = &HFF Then
'
'                       size = GetBits(8): SizePlus = &H128
'                       Do While size = &HFF
'                          size = GetBits(8): SizePlus = SizePlus + &HFF
'                       Loop
'
'                    End If
'                 End If
'              End If
'           End If
'
'         ' Duplicate/Copy String
'           DeComp.FixedString = Mid(DeComp.data, DeComp.Length - RewindBytes, size + SizePlus + 3)
'
'        End If
'
'      Loop
'
'End Sub
     
     

      
      
'Private Function GetBits(NumOfBit) As Long
'TODO       : GetBits implementation
'TODO Status: incomplete
''         Dim bits%
''         For i = 0 To bits
''            Dim CompData&
''            CompData = .int16
''            CompData = CompData * 2 'shl 1
''            Bitcount = 16
''            Bitcount = Bitcount - 1
''         Next
''         CompData = CompData \ &H10000 'shr 0x10
'
'End Function
   
Private Function ADLER32$(Data As StringReader)
   With Data
'            Dim a
            
            Dim l&, H&
            H = 0: l = 1
'            a = GetTickCount
' taken out for performance reason
'               .EOS = False
'               .DisableAutoMove = False
'               Do Until .EOS
'                 'The largest prime less than 2^16
'                  l = (.int8 + l) Mod 65521 '&HFFF1
'                  H = (H + l) Mod 65521 '&HFFF1
'                  If (l And 8) Then DoEvents
'               Loop
'
'            Debug.Print "a: ", GetTickCount - a 'Benchmark: 20203

 '           a = GetTickCount
               
               Dim StrCharPos&, tmpBuff$
               tmpBuff = StrConv(.mvardata, vbFromUnicode)
'               tmpBuff = .mvardata
               For StrCharPos = 1 To Len(.mvardata)
                  'The largest prime less than 2^16
                  l = (AscB(MidB$(tmpBuff, StrCharPos, 1)) + l) Mod 65521 '&HFFF1
                  H = (H + l) Mod 65521 '&HFFF1
                  
                  If 0 = (StrCharPos Mod &H8000) Then DoEvents

               Next
'            Debug.Print "b: ", GetTickCount - a 'Benchmark: 5969

      ADLER32 = H16(H) & H16(l)
   End With
End Function


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
            
            log String(79, "=")
            DeObfuscate.DeObfuscate
            
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
