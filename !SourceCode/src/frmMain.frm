VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "myAut2Exe >The Open Source AutoIT/AutoHotKey script decompiler<"
   ClientHeight    =   9465
   ClientLeft      =   2670
   ClientTop       =   1005
   ClientWidth     =   9300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9465
   ScaleWidth      =   9300
   Begin VB.ListBox List_Positions 
      Height          =   2010
      Left            =   6840
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer_TriggerLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7800
      Top             =   120
   End
   Begin VB.TextBox txt_FILE_DecryptionKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   14
      Tag             =   "18EE"
      Text            =   "18EE"
      ToolTipText     =   "That Box is mean for the FILE-decryptionKey - normally there should be no reason to touch this."
      Top             =   9000
      Width           =   495
   End
   Begin VB.Frame Fr_Options 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   8520
      Width           =   9135
      Begin VB.CommandButton cmd_scan 
         Caption         =   "<<"
         Height          =   255
         Left            =   7605
         TabIndex        =   16
         Top             =   495
         Width           =   375
      End
      Begin VB.TextBox Txt_Scriptstart 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   13
         ToolTipText     =   $"frmMain.frx":628A
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Chk_TmpFile 
         Caption         =   "Don't delete temp files (for ex. compressed scriptdata)"
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox Chk_NormalSigScan 
         Caption         =   "Use 'normal' Au3_Signature to find start of script"
         Height          =   195
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.CheckBox Chk_NoDeTokenise 
         Caption         =   "Disable Detokeniser"
         Height          =   195
         Left            =   7440
         TabIndex        =   10
         ToolTipText     =   "Enable that when you decompile AutoItScripts lower than ver 3.1.6"
         Top             =   600
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.CheckBox Chk_verbose 
         Caption         =   "Verbose LogOutput"
         Height          =   195
         Left            =   4920
         MaskColor       =   &H8000000F&
         TabIndex        =   9
         Top             =   510
         Width           =   1800
      End
      Begin VB.CheckBox Chk_ForceOldScriptType 
         Caption         =   "Force Old Script Type"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   480
         Value           =   2  'Grayed
         Width           =   3255
      End
      Begin VB.CheckBox Chk_RestoreIncludes 
         Caption         =   "Restore Includes"
         Height          =   195
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmd_MD5_pwd_Lookup 
      Caption         =   "Lookup Passwordhash"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      ToolTipText     =   "Copies hash to clipboard and does an online query."
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer_TriggerLoad_OLEDrag 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   480
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
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "Drag the compiled AutoItExe / AutoHotKeyExe or obfucated script in here, or enter/paste path+filename."
      ToolTipText     =   "Drag in or type in da file"
      Top             =   120
      Width           =   9135
   End
   Begin VB.ListBox ListLog 
      Height          =   2010
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "Double click to see more !"
      Top             =   6480
      Width           =   9135
   End
   Begin VB.ListBox List_Source 
      Appearance      =   0  'Flat
      Height          =   5685
      ItemData        =   "frmMain.frx":6321
      Left            =   120
      List            =   "frmMain.frx":6323
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox Txt_Script 
      Height          =   5775
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   600
      Width           =   9135
   End
   Begin VB.Menu mu_Tools 
      Caption         =   "&Tools"
      Begin VB.Menu RegExp_Renamer 
         Caption         =   "&RegExp_Renamer"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mi_FunctionRenamer 
         Caption         =   "&FunctionRenamer"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mi_SeperateIncludes 
         Caption         =   "&Seperate includes of *.au3"
      End
   End
   Begin VB.Menu mu_Info 
      Caption         =   "&Info"
      Begin VB.Menu mi_About 
         Caption         =   "About"
         Visible         =   0   'False
      End
      Begin VB.Menu mi_Update 
         Caption         =   "&Update"
      End
      Begin VB.Menu mi_Forum 
         Caption         =   "&Forum"
      End
   End
   Begin VB.Menu mi_MD5_pwd_Lookup 
      Caption         =   "Lookup Passwordhash"
      Visible         =   0   'False
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

'Const MD5_CRACKER_URL$ = "http://www.md5cracker.de/crack.php?form=Cracken&md5="
'Const MD5_CRACKER_URL$ = "http://web18.server10.nl.kolido.net/md5cracker/crack.php?form=Cracken&md5="

Const MD5_CRACKER_URL$ = "http://hashkiller.com/api/api.php?md5="

'   http://www.milw0rm.com/cracker/info.php?'


Sub FL_verbose(Text)
   log_verbose H32(File.Position) & " -> " & Text
End Sub

Sub log_verbose(TextLine$)
   If Chk_verbose.value = vbChecked Then Log TextLine
End Sub



Sub FL(Text)
   Log H32(File.Position) & " -> " & Text
End Sub

Public Sub LogSub(TextLine$)
   Log "  " & TextLine
End Sub


Public Sub log2(TextLine$)
'   log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Public Sub Log(TextLine$)
On Error Resume Next
   ListLog.AddItem TextLine
'   ListLog.AddItem H32(GetTickCount) & vbTab & TextLine
 
 ' Process windows messages (=Refresh display)
   If RangeCheck(ListLog.ListCount, 10000) Then
       ' Scroll to last item ; when there are more than &h7fff items there will be an overflow error
      Dim ListCount&
      ListLog.ListIndex = ListLog.ListCount - 1
      DoEvents
      
   ElseIf (Rnd < 0.1) Then
      DoEvents
      
   End If
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Public Sub Log_Clear()
On Error Resume Next
   ListLog.Clear
End Sub




Private Sub Chk_ForceOldScriptType_Click()
   Static value
   Checkbox_TriStateToggle Chk_ForceOldScriptType, value
End Sub
Private Sub Checkbox_TriStateToggle(CheckBox As CheckBox, value)
   Static Block_Click As Boolean
   If Block_Click = False Then
      Block_Click = True
      
      With CheckBox

         If value = vbGrayed Then
            value = vbUnchecked
         Else
            value = value + 1
         End If
         .value = value
         
      End With
      
      Block_Click = False
   End If
End Sub

Private Sub Chk_verbose_Click()
   Static value
   Checkbox_TriStateToggle Chk_verbose, value

End Sub

Private Sub Cmd_About_Click()
   FrmAbout.Show vbModal
End Sub

Private Sub ListLogClear()
   ListLog.Clear
End Sub

Private Sub ListLog_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete, vbKeyBack
         ListLogClear
   End Select

End Sub

Private Sub ListLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = MouseButtonConstants.vbRightButton Then
      ListLogClear
   End If
End Sub

'Copies hash to clipboard and does an online query.
Private Sub mi_MD5_pwd_Lookup_Click()
   Clipboard.Clear
   Clipboard.SetText MD5PassphraseHashText

   Dim hProc&
   hProc = ShellExecute(0, "open", MD5_CRACKER_URL$ & LCase$(MD5PassphraseHashText), "", "", 1)

End Sub


'///////////////////////////////////////////
'// General Load/Save Configuration Setting
Private Function ConfigValue_Load(Key$, Optional DefaultValue)
   ConfigValue_Load = GetSetting(App.Title, Me.Name, Key, DefaultValue)
End Function
Property Let ConfigValue_Save(Key$, value As Variant)
      SaveSetting App.Title, Me.Name, Key, value
End Property



'///////////////////////////////////////////
'// Load/Save a CheckBox State
Sub CheckBox_Load(ByVal ChkBox As CheckBox)
   ChkBox.value = ConfigValue_Load(ChkBox.Name, ChkBox.value)
End Sub
Sub CheckBox_Save(ByVal ChkBox As CheckBox)
   ConfigValue_Save(ChkBox.Name) = ChkBox.value
End Sub


Sub TextBox_Load(ByVal Txt As Textbox)
   With Txt
      'signal [txt]_change that were and load the settings
      'so it might react on this i.e. like not the execute the event handler code
      .Enabled = False
         .Text = ConfigValue_Load(Txt.Name, Txt.Text)
      .Enabled = True
   End With
 End Sub
Sub TextBox_Save(ByVal Txt As Textbox)
  'don't save Multiline Textbox
   If Txt.MultiLine = False Then
      ConfigValue_Save(Txt.Name) = Txt.Text
   End If
End Sub



'///////////////////////////////////////////
'// Load/Save a Form Setting
  'Iterate through all Item on the OptionsFrame
  'incase it's no Checkbox a 'type mismatch error' will occur
  'and due to "On Error Resume Next" it skip the call
Sub FormSettings_Load()
   On Error Resume Next
   
   Dim controlItem
   For Each controlItem In Fr_Options.Container
      
      Select Case TypeName(controlItem)
      Case "TextBox"
         If (controlItem Is Txt_Filename) = False Then
            TextBox_Load controlItem
         End If

      Case "CheckBox"
         CheckBox_Load controlItem
      
      End Select
   
   Next
 
End Sub
Sub FormSettings_Save()
   On Error Resume Next
   
   Dim controlItem
   For Each controlItem In Fr_Options.Container
      CheckBox_Save controlItem
      TextBox_Save controlItem
   Next
End Sub


Private Sub cmd_scan_Click()
   LongValScan
   List_Positions.Visible = True
End Sub

Private Sub Form_Load()

'   Dim str$, i&
'   Dim leni%
'   Do
'      BenchStart
'      For i = 0 To 5000000
'         Dim a
'         ArrayEnsureBounds a
'
'      Next
'      BenchEnd
'   Loop While True

   
   FrmMain.Caption = FrmMain.Caption & " " & App.Major & "." & App.Minor & " build(" & App.Revision & ")"
   
   FormSettings_Load
  
  'Just for the case of the first run
   txt_FILE_DecryptionKey_Change
   txt_FILE_DecryptionKey_Validate True
   
   'Extent Listbox width
   Listbox_SetHorizontalExtent ListLog, 6000
   
 
 ' Commandlinesupport   :)
   ProcessCommandline

  'Show Form if SilentMode is not Enable
   If IsOpt_RunSilent = False Then Me.Show

  
  'Open the File that was set by the commandline
   If IsCommandlineMode Then
      Txt_Filename = FileName
   Else
    ' try Load file in the 'File textbox'
      Timer_TriggerLoad.Enabled = True
   End If

End Sub
   
   
Private Sub ProcessCommandline()

   Dim CommadLine As New CommandLine
   With CommadLine
   
      If .NumberOfCommandLineArgs Then
      
         Log "Cmdline Args: " & .CommandLine
         
         Dim arg
         For Each arg In .getArgs
            
           'Check for options
            If arg Like "[/-]*" Then

               If arg Like "?[qQ]" Then
                  IsOpt_QuitWhenFinish = True
                  LogSub "Option 'QuitWhenFinish' enabled."
                  
               ElseIf arg Like "?[sS]" Then
                  IsOpt_RunSilent = True
                  LogSub "Option 'RunSilent' enabled."
                  
               Else
                  LogSub "ERR_Unknow option: '" & arg & "'"
                  
               End If
               
          ' Check if CommandArg is a FileName
            Else
           
               If IsCommandlineMode Then
                  LogSub "ERR_Invalid Argument ('" & arg & "') filename already set."
                  
               Else
                  If FileExists(arg) Then
                     IsCommandlineMode = True
                     FileName = arg
                     LogSub "FileName : " & arg
                  Else
                     LogSub "ERR_Invalid Argument. Can't open file '" & arg & "'"
                  End If
               End If
               
            End If
         Next
      End If
   End With

   'Verify
   If IsOpt_RunSilent And Not (IsOpt_QuitWhenFinish) Then
      LogSub "ERR 'RunSilent' only makes sence together with 'QuitWhenFinish'. As long as you don't also enable 'QuitWhenFinish' 'RunSilent' is ignored "
      IsOpt_RunSilent = False
   End If

End Sub


Public Function GetLogdata$()
   Dim LogData As New clsStrCat
   LogData.Clear
   Dim i
   If (ListLog.ListCount >= 0) Then
      For i = 0 To ListLog.ListCount
         LogData.Concat (ListLog.List(i) & vbCrLf)
      Next
   Else
      For i = 0 To &H7FFE
         LogData.Concat (ListLog.List(i) & vbCrLf)
      Next
      LogData.Concat "<Data cut due to VB-listbox.ListCount bug :( >"
      
'   Do While ListLog.ListCount < 0
'      LogData.Concat (ListLog.List(&H7FFF) & vbCrLf)
'      ListLog.RemoveItem &H7FFF
'   Loop
   
   End If
   
   GetLogdata = LogData.value
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   FormSettings_Save
  
 'Close might be clicked 'inside' some DoEvents so
 'in case it was do a hard END
   End
End Sub



Private Sub List_Positions_DblClick()
   Txt_Scriptstart = List_Positions.Text
   List_Positions.Visible = False
End Sub

Private Sub ListLog_DblClick()
   frmLogView.txtlog = GetLogdata()
   frmLogView.Show
End Sub

Private Sub mi_Update_Click()
   Dim hProc&
   hProc = ShellExecute(0, "open", "http://myauttoexe2.tk/", "", "", 1)

End Sub
Private Sub mi_Forum_Click()
   Dim hProc&
   hProc = ShellExecute(0, "open", "http://defcon5.biz/phpBB3/viewtopic.php?f=5&t=234", "", "", 1)

End Sub

Private Sub mi_FunctionRenamer_Click()
   Load FrmFuncRename
'   If FileExists(Txt_Filename) Then
'      FrmFuncRename.Txt_Fn_Org_FileName = Txt_Filename
'   Else
'
'
'   End If
   
   FrmFuncRename.Show vbModal
   Unload FrmFuncRename
   
End Sub

Private Sub mi_SeperateIncludes_Click()
   Dim File$
   File = InputBox("Normally seperating includes is done automatically after you decompiled some au3.exe(of old none tokend format)." & vbCrLf & _
          "However that tool is useful in the case you have some decompiled *.au3 with these '; <AUT2EXE INCLUDE-START: C:\ ...' comments you like to process." & vbCrLf & vbCrLf & _
          "Please enter(/paste) full path of the file: (Or drag it into the myAutToExe filebox and then run me again)", "Manually run 'seperate au3 includes' on file", Txt_Filename)
   If File <> "" Then
      FileName.FileName = File
      SeperateIncludes
   End If
End Sub





Private Sub RegExp_Renamer_Click()
   FrmRegExp_Renamer.Show vbModal
   Unload FrmRegExp_Renamer
End Sub

Private Sub Timer_TriggerLoad_OLEDrag_Timer()
   Timer_TriggerLoad_OLEDrag.Enabled = False
   Txt_Filename = FilePath_for_Txt
End Sub


Private Sub Timer_TriggerLoad_Timer()
   Timer_TriggerLoad.Enabled = False
   
   Txt_FileName_Change

End Sub

Private Sub txt_FILE_DecryptionKey_Change()
   With txt_FILE_DecryptionKey
      On Error Resume Next
      .ForeColor = IIf(txt_FILE_DecryptionKey_IsValid, vbBlack, vbRed)
   End With
End Sub

Function txt_FILE_DecryptionKey_IsValid() As Boolean
   With txt_FILE_DecryptionKey
      On Error Resume Next
      H16 "&h" & .Text
      txt_FILE_DecryptionKey_IsValid = (Err = 0) And _
                                       (Len(.Text) <= 4)
   End With
End Function
Private Sub txt_FILE_DecryptionKey_Validate(Cancel As Boolean)
   With txt_FILE_DecryptionKey
      
      If txt_FILE_DecryptionKey_IsValid Then
         .Text = H16("&h" & .Text)
      Else
         .Text = .Tag
      End If
      
      FILE_DecryptionKey_New = "&h" & .Text

   End With

End Sub

Private Sub Txt_FileName_Change()
  'Avoid to be triggered during load settings
   If Txt_Filename.Enabled = False Then Exit Sub
  
   On Error GoTo Txt_Filename_err
   
   cmd_scan.Visible = FileExists(Txt_Filename)
   If cmd_scan.Visible Then
   'If FileExists(Txt_Filename) Then
      
     'Clear Log (expect when run via commandline)
      If IsCommandlineMode = False Then ListLog.Clear
      Txt_Script = ""
      
      FileName = Txt_Filename
      
      Log String(80, "=")
'      log "           -=  " & Me.Caption & "  =-"
      Log Me.Caption
      Log String(80, "=")
         
      Decompile
         Log "Testing for Scripts that were obfuscate by 'Jos van der Zande AutoIt3 Source Obfuscator v1.0.15 [July 1, 2007]' or 'EncodeIt 2.0'"
         Log String(79, "=")
   
      
      FileName = ExtractedFiles("MainScript")
         On Error Resume Next
      DeToken
         If Err Then Log "ERR: " & Err.Description

         On Error Resume Next
         Log String(79, "=")
      
      
      DeObfuscate.DeObfuscate
         If Err Then Log "ERR: " & Err.Description
         Select Case Err
         Case 0, ERR_NO_OBFUSCATE_AUT
            If Chk_RestoreIncludes.value = vbChecked Then _
               SeperateIncludes
               
         Case Else
            Log Err.Description
            
         End Select


      CheckScriptFor_COMPILED_Macro
 
' ErrorHandle for For-Each-Loop
Err.Clear
GoTo Txt_Filename_err

' Decompile Err Handler

      
      
DeToken:
      Log String(79, "=")
      DeToken

DeObfuscate:
      Log String(79, "=")
      DeObfuscate.DeObfuscate
      
Txt_Filename_err:
  ' Add some fileName if it weren't done during decompile()
    If IsAlreadyInCollection(ExtractedFiles, "MainScript") = False Then
       ExtractedFiles.Add File.FileName, "MainScript"
    End If

  
  ' Note: Resume is necessary to reenable Errorhandler
  '       Else the VB-standard Handler will catch the error -> Exit Programm
    Select Case Err
    Case 0
    
    Case ERR_NO_AUT_EXE
       Log Err.Description
       Resume DeToken
    
    Case NO_AUT_DE_TOKEN_FILE
       Log Err.Description
       Resume DeObfuscate
    
    Case ERR_NO_OBFUSCATE_AUT
       Log Err.Description
       Resume Txt_Filename_err
       
       
    Case Else
       Log Err.Description
       Resume Txt_Filename_err
    End Select
'-----------------------------------------------
   
    
    'Save Log Data
    On Error Resume Next
    
    FileName = ExtractedFiles("MainScript").FileName
    FileName.NameWithExt = FileName.Name & "_myExeToAut.log"
    
    Log ""
    Log "Saving Logdata to : " & FileName.FileName
    File.Create FileName.FileName, True
    File.FixedString(-1) = GetLogdata
    File.CloseFile
    
    
    IsCommandlineMode = False
    
    If IsOpt_QuitWhenFinish Then Unload Me
   
   End If
   
End Sub


Private Function OpenFile(Target_FileName As ClsFilename) As Boolean
   
   On Error GoTo Scanfile_err
   Log "------------------------------------------------"

   Log Space(4) & Target_FileName.NameWithExt

   File.Create Target_FileName.mvarFileName, Readonly:=True
   
   Me.Show

Err.Clear
Scanfile_err:
Select Case Err
   Case 0

   Case Else
      Log "-->ERR: " & Err.Description

End Select
   
End Function


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub List_Source_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub ListLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub Txt_Script_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub txt_FileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub File_DragDrop(Data As DataObject)
   
   On Error GoTo Txt_FileName_OLEDragDrop_err
   
   FilePath_for_Txt = Data.Files(1)
   Timer_TriggerLoad_OLEDrag.Enabled = True
   

Txt_FileName_OLEDragDrop_err:
Select Case Err
Case 0

Case Else
   Log "-->Drop'n'Drag ERR: " & Err.Description

End Select

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

Private Sub Txt_Scriptstart_Change()
   On Error Resume Next
   Dim scriptstart&
   scriptstart = "&h" & Txt_Scriptstart
   
   Chk_NormalSigScan.Enabled = (Err.Number <> 0)
   
End Sub
