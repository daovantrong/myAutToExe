VERSION 5.00
Begin VB.Form FrmFuncRename 
   Caption         =   "Function Renamer"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List_Fn_String_Org 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   0
      TabIndex        =   26
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmd_AutoAdd 
      Caption         =   "AutoAdd"
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   6000
      TabIndex        =   21
      Top             =   4200
      Width           =   6255
      Begin VB.TextBox Txt_SearchSync 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   23
         Text            =   "Select below some text to search for"
         Top             =   0
         Width           =   2775
      End
      Begin VB.CommandButton Cmd_FindNext_Inc 
         Caption         =   "Find next"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_SearchSyncStatus_Org 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3960
         TabIndex        =   24
         Top             =   0
         Width           =   2235
      End
   End
   Begin VB.Frame Frame_Search_Org 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   17
      Top             =   4080
      Width           =   6015
      Begin VB.TextBox Txt_SearchSync_Org 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Text            =   "Select below some text to search for"
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton Cmd_FindNext_Org 
         Caption         =   "Find next"
         Height          =   255
         Left            =   2805
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_SearchSyncStatus_Inc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3840
         TabIndex        =   20
         Top             =   120
         Width           =   2235
      End
   End
   Begin VB.TextBox Txt_Include 
      Height          =   285
      Left            =   10320
      TabIndex        =   16
      Text            =   "Txt_Include"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_Remove_assign 
      Caption         =   "v  Remove func assignment  v"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   1800
      Width           =   2895
   End
   Begin VB.FileListBox File_Includes 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   10320
      Pattern         =   "*.au3"
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_DoSearchAndReplace 
      Caption         =   "Apply search and replace"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7080
      TabIndex        =   6
      Top             =   645
      Width           =   2295
   End
   Begin VB.CommandButton Cmd_AddSearchAndReplace 
      Caption         =   "^Add func search 'n' replace^"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      ToolTipText     =   "Short cut: double click or 'Enter'"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox Txt_Fn_Inc 
      Appearance      =   0  'Flat
      Height          =   4935
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Text            =   "FrmFuncRename.frx":0000
      Top             =   4560
      Width           =   5535
   End
   Begin VB.ListBox List_Fn_Inc 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   5040
      TabIndex        =   13
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox Txt_Fn_Inc_FileName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "<Drag some au3-include file in here> For example: C:\AutoIt3\Include\Array.au3"
      Top             =   0
      Width           =   5775
   End
   Begin VB.ListBox List_Fn_Assigned 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   6975
   End
   Begin VB.TextBox Txt_Fn_Org_FileName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "<Drag deObfuscated au3-file in here>"
      Top             =   0
      Width           =   5775
   End
   Begin VB.TextBox Txt_Fn_Org 
      Appearance      =   0  'Flat
      Height          =   4935
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Text            =   "FrmFuncRename.frx":000D
      Top             =   4560
      Width           =   3255
   End
   Begin VB.ListBox List_Fn_Org 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CommandButton cmd_org_reload 
      Appearance      =   0  'Flat
      Caption         =   "Reload"
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   285
      Width           =   2175
   End
   Begin VB.CommandButton cmd_inc_reload 
      Appearance      =   0  'Flat
      Caption         =   "Reload"
      Height          =   330
      Left            =   6120
      TabIndex        =   3
      Top             =   285
      Width           =   1815
   End
   Begin VB.CommandButton cmd_Save 
      Caption         =   "Save"
      Height          =   615
      Left            =   7080
      TabIndex        =   7
      Top             =   1125
      Width           =   855
   End
   Begin VB.CommandButton cmd_Load 
      Caption         =   "Load"
      Height          =   615
      Left            =   7920
      TabIndex        =   8
      Top             =   1125
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   615
      Left            =   8640
      TabIndex        =   9
      Top             =   1125
      Width           =   735
   End
   Begin VB.Line Line 
      BorderWidth     =   3
      Index           =   2
      X1              =   2520
      X2              =   8280
      Y1              =   285
      Y2              =   885
   End
   Begin VB.Line Line 
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   4080
      X2              =   4080
      Y1              =   4680
      Y2              =   4320
   End
   Begin VB.Line Line 
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   4080
      X2              =   6600
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "FrmFuncRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private File_Org As New FileStream
'Private File_Inc As New FileStream

Private File_Org_FileName As New ClsFilename

Private Script_Org As New StringReader
Private Script_Inc As New StringReader

Dim Functions_Org
Dim Functions_Inc

Const FN_ASSIGNED_FUNC_REPL_SEP$ = " => "
Dim List_Fn_Assigned_FuncIdxs As New Collection

Dim NumOccurrenceFound& 'From SearchSync

Dim List_Fn_String_Org_EventBlocker As Boolean

'///////////////////////////////////////////
'// General Load/Save Configuration Setting
Private Function ConfigValue_Load(Key$, Optional DefaultValue)
   ConfigValue_Load = GetSetting(App.Title, Me.Name, Key, DefaultValue)
End Function
Property Let ConfigValue_Save(Key$, value As Variant)
      SaveSetting App.Title, Me.Name, Key, value
End Property

Private Sub Log(Text$)
   FrmMain.Log "FuncRepl: " & Text
End Sub

Private Sub cmd_AutoAdd_Click()
 ' event Blocker
   Static on_cmd_AutoAdd As Boolean
   If on_cmd_AutoAdd = True Then Exit Sub
   on_cmd_AutoAdd = True
      
      Dim i
      With List_Fn_Inc
         
 '        Do
 '           List_Fn_String_Org_EventBlocker = True
            
            Dim ListItems&
            ListItems = .ListCount - 1
            For i = 0 To ListItems
               
               If SearchAndReplace_AddItem(True) = False Then
                  GoTo cmd_AutoAdd_ClickQuit
               End If
               
            Next
            
 '           List_Fn_String_Org_EventBlocker = False
 '           List_Fn_String_Org_Click
 '
 '        Loop While List_Fn_Org.ListCount And ListItems > 0
         
      End With
cmd_AutoAdd_ClickQuit:
'   List_Fn_String_Org_EventBlocker = False
   on_cmd_AutoAdd = False
End Sub

Private Sub Cmd_DoSearchAndReplace_Click()

On Error GoTo Cmd_DoSearchAndReplace_Click_Err

   If List_Fn_Assigned.ListCount = 0 Then
      Exit Sub
   End If
 
 
 ' Open file
   Dim File_Org As New FileStream
   Dim SearchAndReplBuff$
   With File_Org
      .Create File_Org_FileName.FileName, False, False
      SearchAndReplBuff = .FixedString(-1)
      .CloseFile
   End With
   
 'Search and Replace
   Dim SearchAndReplaceJob_Line$
   Dim SearchAndReplace_LookFor$
   Dim SearchAndReplace_ReplaceWith$
   Dim SearchAndReplace_Include$
   
   Dim MousePointer_Backup%
   MousePointer_Backup = MousePointer
   MousePointer = vbHourglass
   
   With List_Fn_Assigned
      Dim ListItemIdx%
      For ListItemIdx = 0 To .ListCount - 1
         SearchAndReplaceJob_Line = .List(ListItemIdx)
         
         Dim tmp
         tmp = Split(SearchAndReplaceJob_Line, FN_ASSIGNED_FUNC_REPL_SEP)
         SearchAndReplace_LookFor = tmp(0)
         SearchAndReplace_ReplaceWith = tmp(1)
         SearchAndReplace_Include = tmp(2)
      
       ' Delete Function (& insert include)
         Dim FunctionBody_ReplaceText$
         FunctionBody_ReplaceText = ";  Func " & SearchAndReplace_ReplaceWith
         
         If SearchAndReplace_Include <> "" Then
            '";" & String(79, "=") &
            FunctionBody_ReplaceText = vbCrLf & vbCrLf & "#include <" & SearchAndReplace_Include & ">" & vbCrLf & FunctionBody_ReplaceText
            Log "Deleting Func '" & SearchAndReplace_LookFor & "' & adding include '" & SearchAndReplace_Include & "'"
         Else
            Log "Deleting Func " & SearchAndReplace_LookFor
         End If
         
         
         
         strCropAndDelete SearchAndReplBuff, _
                           vbCrLf & vbCrLf & "Func " & SearchAndReplace_LookFor, _
                           "EndFunc" & vbCrLf, , , _
                           FunctionBody_ReplaceText
       
       
       ' Replace all
         Dim ReplacementsDone&
         ReplacementsDone = &H7FFFFFFF
         ReplaceDo SearchAndReplBuff, SearchAndReplace_LookFor, SearchAndReplace_ReplaceWith, 1, ReplacementsDone
'         .List(ListItemIdx) = ReplacementsDone & vbTab & .List(ListItemIdx)
         Log ReplacementsDone & " occurence of " & SearchAndReplaceJob_Line & " found & replaced."
         
       ' Mark unused functions with ;;
         If ReplacementsDone = 0 Then
            ReplaceDo SearchAndReplBuff, FunctionBody_ReplaceText, Replace(FunctionBody_ReplaceText, "; ", ";;"), 1
         End If
      Next
   End With
   
   MousePointer = MousePointer_Backup
'   Cmd_DoSearchAndReplace.Enabled = False
   
 
 
 
 ' save file
   Dim File_Org_Save As New FileStream
   With File_Org_Save
      .Create File_Org_FileName.Path & File_Org_FileName.Name & "_FuncRenamed.au3", True, False, False
      .FixedString(-1) = SearchAndReplBuff
      Log "Search&Replace complete.  Output File: " & .FileName
      .CloseFile
   End With
   
  


Err.Clear
Cmd_DoSearchAndReplace_Click_Err:
Select Case Err
   Case 0

   Case Else
      MsgBox Err.Description, vbCritical, "Error " & Hex(Err.Number) & "  in Formular FrmFuncRename.Cmd_DoSearchAndReplace_Click()"

End Select

End Sub


Private Sub OpenAndFill(FileName$, ScriptData As StringReader, FuncList, List_Func As Listbox)
' Open RightFile
   Dim InputFile As New FileStream
   With InputFile
      .Create FileName, False, False, True
      ScriptData.Data = .FixedString(-1)
      .CloseFile
   End With
   
'Seperate functions
'   Txt_Fn_Org = ScriptData.Data

   FuncList = Split(ScriptData.Data, vbCrLf & "Func ", , vbTextCompare)
   'ReDim Preserve FuncList(1 To UBound(FuncList))
   
   With List_Func
      .Clear
      Dim itemidx
      For itemidx = 1 To UBound(FuncList)
        'Add FunctionName
         .AddItem (Split(FuncList(itemidx), vbCrLf)(0))
        
        'Store index of FuncList to find it later
         .ItemData(.ListCount - 1) = itemidx
      Next
      
    If .ListCount Then .ListIndex = 0
   End With
End Sub


Private Sub Cmd_AddSearchAndReplace_Click()
   SearchAndReplace_AddItem
End Sub



Private Sub Cmd_Remove_assign_Click()
   SearchAndReplace_RemoveItems
End Sub


Private Sub SearchAndReplace_RemoveItems()
'   On Error GoTo Cmd_Remove_assign_Click_err
   With List_Fn_Assigned
      If .ListCount = 0 Then Exit Sub
      
      Dim Indexes()
      Indexes = List_Fn_Assigned_FuncIdxs(.ItemData(.ListIndex))
      
     'Add FunctionName & Index to Original
      With List_Fn_Org
         .AddItem Split(Functions_Org(Indexes(0)), vbCrLf)(0), 0
         .ItemData(0) = Indexes(0)
      End With
      
     'Add FunctionName & Index to includes
      With List_Fn_Inc
         .AddItem Split(Functions_Inc(Indexes(1)), vbCrLf)(0), 0
         .ItemData(0) = Indexes(1)
      End With
      
     'remove assign entry
      Listbox_removeCurrentItemAndSelectNext List_Fn_Assigned
      
   End With
Cmd_Remove_assign_Click_err:
End Sub





Private Sub cmd_Load_Click()
   LoadSearchReplaceData
End Sub
'---------------------------------------------------------------------------------------
' Procedure : LoadSearchReplaceData
' Author    : Administrator
' Date      : 17.03.2008
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub LoadSearchReplaceData()
   
'Load Data
   Dim Textlines
   Dim myfile As New FileStream
   
   

On Error GoTo LoadSearchReplaceData_Err

   With myfile
   .Create "SearchReplaceData.txt", False, False, True
      Textlines = Split(.FixedString(-1), vbCrLf)
   .CloseFile
   End With
   
   With List_Fn_Assigned

'seperate Org & Inc Functions
      Dim item
      For Each item In Textlines
         Dim Textline_items
         Textline_items = Split(item, " => ")
         
         If UBound(Textline_items) > 0 Then
            
            Dim funcNameOrg$
            funcNameOrg = Textline_items(0)
            
            Dim funcNameInc$
            funcNameInc = Textline_items(1)
            
            Dim Include$
            Include = Textline_items(2)
            
            ListBox_FindAndSelectedItem File_Includes, Include

            
   ' Find Item in Org list
            With List_Fn_Org
               
               Dim i&, Found As Boolean
               Found = False
               
               For i = 0 To .ListCount
                  If funcNameOrg = Left(.List(i), Len(funcNameOrg)) Then
                     .ListIndex = i
                     Found = True
                     Exit For
                  End If
               Next
            End With
            
            If Found Then
   ' Find Item in Inc List
               With List_Fn_Inc
      
                  Found = False
                  For i = 0 To .ListCount
                  If funcNameInc = Left(.List(i), Len(funcNameInc)) Then
                     .ListIndex = i
                     Found = True
                     Exit For
                     End If
                  Next
               End With
               If Found Then
   'Add items to search'n'replace
                  SearchAndReplace_AddItem
               Else
                  Log "Load_Error: Item not found in IncludeList. CurLine: '" & item & "'"
                  
'                 'Add item to search'n'replace
'                  With List_Fn_Assigned
'                     .AddItem funcNameOrg & FN_ASSIGNED_FUNC_REPL_SEP & funcNameInc
'                     .ListIndex = .ListCount - 1
'                    ' Store Functionidx finding&display functionText on click
''                      List_Fn_Assigned_FuncIdxs.Add Array(List_Fn_Org.ListIndex, 1)
''                     .ItemData(.ListIndex) = List_Fn_Assigned_FuncIdxs.Count
'
'                  End With
'                 ' Delete from list
'                  With List_Fn_Org
'                     .RemoveItem .ListIndex
'                  End With
                  
               End If ' Found in org
            Else
               Log "Load_Error: Item not found in OrginalList. CurLine: '" & item & "'"
            End If ' Found in Inc
'         Else
'            Log "Load_Error: Missing ' => ' seperator in line: '" & item & "'"
         End If 'Split at " => "
      Next
   End With

Err.Clear
LoadSearchReplaceData_Err:
Select Case Err
   Case 0

   Case Else
      MsgBox Err.Description, vbCritical, "Error " & Hex(Err.Number) & "  in Formular FrmFuncRename.LoadSearchReplaceData()"

End Select
   
End Sub


Private Sub Cmd_Save_Click()
   SaveSearchReplaceData
End Sub
Private Sub SaveSearchReplaceData()
   Dim myfile As New FileStream
   

On Error GoTo SaveSearchReplaceData_Err

   With myfile
   .Create "SearchReplaceData.txt", True
      .FixedString(-1) = GetListBoxData(List_Fn_Assigned)
   .CloseFile
   End With

Err.Clear
SaveSearchReplaceData_Err:
Select Case Err
   Case 0

   Case Else
      MsgBox Err.Description, vbCritical, "Error " & Hex(Err.Number) & "  in Formular FrmFuncRename.SaveSearchReplaceData()"

End Select

End Sub

Private Sub cmdHelp_Click()
   MsgBox ("First of all you only need that if the decompiled file was obfuscated and so variable and function got lost." & vbCrLf & _
   "" & vbCrLf & _
   "1. Drag the decompiled file on the upper left textbox." & vbCrLf & _
   "2. Drag the some of au3 include file on the upper right textbox." & vbCrLf & _
   "" & vbCrLf & _
   "3.Choose some function of the includes and mark some string or other unique in the function detail view to start the search" & vbCrLf & _
   "" & vbCrLf & _
   "4. Now if both seem to match, doubleclick on it or click on 'Add func'" & vbCrLf & _
   "" & vbCrLf & _
   "5. when you're done click on 'apply search'n'replace'." & vbCrLf & _
   "" & vbCrLf & _
   "'Save' will save current Search'n'Replace Data to 'SearchReplaceData.txt'." & vbCrLf & _
   "'Load' will load 'SearchReplaceData.txt' and automatically select both functions + do an 'Add func' click." & vbCrLf & _
   "")
   
End Sub

Private Function GetListBoxData$(Listbox As Listbox)
   
   With Listbox
      Dim LogData As New clsStrCat
      LogData.Clear
      Dim i
      For i = 0 To .ListCount
         LogData.Concat (.List(i) & vbCrLf)
      Next
   End With
   
   GetListBoxData = LogData.value
   
End Function

Private Sub File_Includes_Click()
   Dim tmpFileName$
   tmpFileName = File_Includes.Path & "\" & File_Includes.FileName
   If Txt_Fn_Inc_FileName <> tmpFileName Then
     'Triggers _ChangeText
      Txt_Fn_Inc_FileName = tmpFileName
   End If
End Sub

'///////////////////////////////////////////
'// Form_Load
Private Sub Form_Load()
   
 ' Load Configuration Setting
   With Txt_Fn_Org_FileName
      .Text = ConfigValue_Load(.Name, .Text)
   End With
   
   With Txt_Fn_Inc_FileName
      .Text = ConfigValue_Load(.Name, .Text)
   End With
   
End Sub

Private Sub Form_Resize()
 'Width
   'rightside
      Txt_Fn_Org_FileName.Width = (Me.Width \ 2)
      
      List_Fn_Org.Width = (Me.Width \ 2)
      
      Txt_Fn_Org.Width = (Me.Width \ 2)
      
      
      
      
   'leftside
      Txt_Fn_Inc_FileName.Left = Me.Width \ 2
      Txt_Fn_Inc_FileName.Width = (Me.Width \ 2) - 150
      
      cmd_inc_reload.Left = Me.Width \ 2
      
      File_Includes.Left = Me.Width - File_Includes.Width - 150
      
      Txt_Include.Left = File_Includes.Left
      
      Txt_Fn_Inc.Left = Me.Width \ 2
      Txt_Fn_Inc.Width = (Me.Width \ 2) - 150
      
      List_Fn_Inc.Left = Me.Width \ 2
      List_Fn_Inc.Width = (Me.Width \ 2) - 150
  
  'center
  
      Cmd_AddSearchAndReplace.Left = (Me.Width \ 2) - (Cmd_AddSearchAndReplace.Width)
      Cmd_Remove_assign.Left = (Me.Width \ 2)
 
 
 'Height
   On Error Resume Next
   
   With List_Fn_String_Org
      .Height = (Me.Height - .Top) - 550
   End With
   
   
   With Txt_Fn_Inc
      .Height = (Me.Height - .Top) - 550
   End With
   
   With Txt_Fn_Org
      .Height = (Me.Height - .Top) - 550
   End With
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   With Txt_Fn_Org_FileName
      If FileExists(.Text) Then ConfigValue_Save(.Name) = .Text
   End With
   
   With Txt_Fn_Inc_FileName
      If FileExists(.Text) Then ConfigValue_Save(.Name) = .Text
   End With

End Sub

Private Sub List_Fn_Assigned_Click()
   On Error GoTo List_Fn_Assigned_Click_err
   With List_Fn_Assigned
      Dim Indexes()
      Indexes = List_Fn_Assigned_FuncIdxs(.ItemData(.ListIndex))
      
      Txt_Fn_Inc = Functions_Inc(Indexes(1))
      Txt_Fn_Org = Functions_Org(Indexes(0))
   End With
List_Fn_Assigned_Click_err:
End Sub

Private Sub List_Fn_Assigned_DblClick()
   SearchAndReplace_RemoveItems
End Sub

Private Sub List_Fn_Assigned_KeyPress(KeyAscii As Integer)
   
   Select Case KeyAscii
      Case vbKeyDelete
         SearchAndReplace_RemoveItems
      Case Else
      
   End Select
   

End Sub

Private Sub List_Fn_Org_Click()
   With List_Fn_Org
      Txt_Fn_Org = Functions_Org(.ItemData(.ListIndex))
   End With
   List_Fn_String_Org_Fill
End Sub

Private Sub List_Fn_String_Org_Fill()
   List_Fn_String_Org.Clear
 
   Dim Matches As MatchCollection
   GetStrings Txt_Fn_Org, Matches
   
   
   Dim ItemStringMaxLength&: ItemStringMaxLength = 0
   Dim List_Fn_String_Org_LongestStringIndex&: List_Fn_String_Org_LongestStringIndex = 0
   
   
   Dim item As Match
   For Each item In Matches
      
      If ItemStringMaxLength < Len(item) Then
            ItemStringMaxLength = Len(item)
            List_Fn_String_Org_LongestStringIndex = List_Fn_String_Org.ListCount
      End If
      
      
      List_Fn_String_Org.AddItem item
   Next
   
  'Only select a string if it's at least 5 bytes long
   If ItemStringMaxLength >= 5 Then
      
      List_Fn_String_Org_EventBlocker = True
      List_Fn_String_Org.ListIndex = List_Fn_String_Org_LongestStringIndex
      List_Fn_String_Org_EventBlocker = False
   End If
   
End Sub
Private Sub List_Fn_Inc_Click()
 ' Show FunctionText
   With List_Fn_Inc
      Txt_Fn_Inc = Functions_Inc(.ItemData(.ListIndex))
   End With
End Sub

Private Sub ListBox_FindAndSelectedItem(Listbox As Object, itemText$)
   'On Error Resume Next
   With Listbox
      Dim item
      For item = 0 To .ListCount - 1
         If .List(item) Like itemText Then
            .ListIndex = item
            Exit For
         End If
      Next
      
   End With

End Sub


Private Sub ListBox_ScrollToFirstSelected(Listbox As Listbox)
   'On Error Resume Next
   With Listbox
      Dim ListIndex_Backup%
      ListIndex_Backup = .ListIndex
      
      .ListIndex = .ListCount - 1
      
      .ListIndex = ListIndex_Backup
      
   End With

End Sub

'Private Sub SearchAndReplace_AddItems()
'
'   'Go through all selected in includes
'   With List_Fn_Inc
'      Dim item
'      For item = 0 To .ListCount - 2
'         If .Selected(item) Then
'          ' Set to Listitem in Function Includes
'            .ListIndex = item
'
'           'Find first selected in Orginal
'            With List_Fn_Org
'            Dim item2
'            For item2 = 0 To .ListCount - 1
'               If .Selected(item2) Then
'                ' Set to Listitem in Function Original
'                  .ListIndex = item2
'
'                 'Do Add to Search'n'Replace list
'                  SearchAndReplace_AddItem
'
'               Exit For
'               End If
'            Next
'
'
'         End With
'
'         End If
'      Next
'
'   End With
'
'End Sub

Private Sub Listbox_removeCurrentItemAndSelectNext(Listbox As Listbox)
   With Listbox
      If (.ListIndex < (.ListCount - 1)) Then
         .ListIndex = .ListIndex + 1
         .RemoveItem .ListIndex - 1
      Else
         If .ListIndex = 0 Then
            .RemoveItem .ListIndex
         Else
            .ListIndex = .ListIndex - 1
            .RemoveItem .ListIndex + 1
         End If
      End If
   End With
End Sub

'Private Sub RE_MatchesToArray(Matches As MatchCollection, MatchArray As Variant)
'
'End Sub

Private Sub GetStrings(Textbox As Textbox, Matches As MatchCollection) 'Strings As Variant)
   
   RemoveComments Textbox
     
   Dim myRegExp As New RegExp
   With myRegExp 'New RegExp
      .Global = True
      .MultiLine = True

      .Pattern = "(([""']).*?\2)+"
      
'      Dim Matches As MatchCollection
      Set Matches = .Execute(Textbox.Text)
      

   End With
End Sub
Private Sub RemoveComments(Textbox As Textbox) 'Strings As Variant)
   
   Const StringBody_SingleQuoted As String = "[^']*"
   Const String_SingleQuoted = "(?:'" & StringBody_SingleQuoted & "')+"

   Const StringBody_DoubleQuoted As String = "[^""]*"
   Const String_DoubleQuoted As String = "(?:""" & StringBody_DoubleQuoted & """)+"
   
   Const StringPattern As String = String_DoubleQuoted & "|" & String_SingleQuoted
  ' /r => carriage return @CR, chr(13)   -   /n => linefeed        @LF, chr(10)
  ' 2 - in Windows it's @CR@LF        -   1 - in Linux/Unix it's just @LF
  ' 0 - at the end of the file there is none of these
   Const LineBreak As String = "\r?\n?"
   Const WhiteSpaces As String = "\s*"

' LineComment_EntiredLine should include the LineBreak in the match -> so whole line
  ' can be deleted - while at 'NotEntiredLineComments' the line break is keept as it is.
  ' BlockCommentEnd is in there for the case there's a line like this: " #ce ;some comment"
   Const LineComment As String = LineBreak & WhiteSpaces & ";[^\r\n]*"
   

   Const BlockCommentStart As String = LineBreak & WhiteSpaces & "\#c(?:s|omments-start)"
   Const BlockCommentEnd As String = "\#c(?:e|omments-end)"
   Const BlockComment As String = BlockCommentStart & "(?:" & _
                        StringPattern & "|" & "." & _
                        ")*?" & BlockCommentEnd

     
   Dim myRegExp As New RegExp
   With myRegExp 'New RegExp
      .Global = True
      .MultiLine = True
            
      .Pattern = "(?:" & LineComment & ")" & _
                 "|" & BlockComment & "|" & _
                 "(" & StringPattern & ")"
      
     'Remove Comments
      Textbox.Text = .Replace(Textbox.Text, "$1")
      

   End With
End Sub

Private Function RE_MatchesCompare(Matches1 As MatchCollection, Matches2 As MatchCollection, Optional UseSubmatch As Integer = -1) As Boolean
   
  'Matches must have the same size to be equal
   If Matches1.Count = Matches2.Count Then
      'Compare Content
      Dim i
      For i = 0 To Matches1.Count - 1
         If UseSubmatch > -1 Then
            If UndoAutoItString(Matches1(i).SubMatches(UseSubmatch)) <> _
               UndoAutoItString(Matches2(i).SubMatches(UseSubmatch)) Then Exit For
         Else
            If UndoAutoItString(Matches1(i)) <> _
               UndoAutoItString(Matches2(i)) Then Exit For
         End If
      Next
      RE_MatchesCompare = (i = Matches1.Count)
      
   End If
End Function



Private Function SearchAndReplace_AddItem(Optional bQuietMode As Boolean = False) As Boolean
     
   If (List_Fn_Inc.ListCount = 0) Or (List_Fn_Org.ListCount = 0) Then Exit Function
   
 ' Validate Arguments
   Dim numArgOrg&, numArgInc&
      numArgOrg = GetNumOccurrenceIn(List_Fn_Org.Text, ",") + 1
      numArgInc = GetNumOccurrenceIn(List_Fn_Inc.Text, ",") + 1
   If numArgOrg <> numArgInc Then
      If bQuietMode Then Exit Function
      If vbYes <> MsgBox("Number of argument of both function differs( " & numArgOrg & " <> " & numArgInc & " )." & vbCrLf & _
                         "Add them anyway?", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Attention") Then
         Exit Function
      End If
   End If
   
 ' Validate Strings
   Dim fn_inc As MatchCollection
   GetStrings Txt_Fn_Inc, fn_inc
   
   Dim fn_org As MatchCollection
   GetStrings Txt_Fn_Org, fn_org
   
   If RE_MatchesCompare(fn_org, fn_inc) = False Then
      If bQuietMode Then Exit Function
'      FrmFuncRename_StringMismatch.Create fn_org, fn_inc
'      FrmFuncRename_StringMismatch.Show vbModal
      If vbNo = MsgBox("Local strings don't match continue anyway?", vbYesNo + vbDefaultButton2) Then
         Exit Function
      End If
   End If
   
  'Add S&R item
   Dim FuncOldName$, FuncOldNameIdx&
   With List_Fn_Org
   
     'Cut at '(' of for ex Func MyNewFunc(Arg1,arg2...
      FuncOldName = Split(.Text, "(")(0)
      FuncOldNameIdx = .ItemData(.ListIndex)
      Listbox_removeCurrentItemAndSelectNext List_Fn_Org
   End With
   
   
   Dim FuncNewName$, FuncNewNameIdx&
   With List_Fn_Inc
   
      FuncNewName = Split(.Text, "(")(0)
      FuncNewNameIdx = .ItemData(.ListIndex)
      Listbox_removeCurrentItemAndSelectNext List_Fn_Inc
      
   End With
   
   
   
   With List_Fn_Assigned
   
      .AddItem FuncOldName & FN_ASSIGNED_FUNC_REPL_SEP & FuncNewName & FN_ASSIGNED_FUNC_REPL_SEP & Txt_Include
      .ListIndex = .ListCount - 1
      
            
      Txt_Include = ""
      

      
    ' Store Functionidx finding&display functionText on click
      List_Fn_Assigned_FuncIdxs.Add Array(FuncOldNameIdx, FuncNewNameIdx)
      .ItemData(.ListIndex) = List_Fn_Assigned_FuncIdxs.Count
   End With
   
   ListBox_ScrollToFirstSelected List_Fn_Org
   ListBox_ScrollToFirstSelected List_Fn_Inc
   
   Cmd_DoSearchAndReplace.Enabled = True
   
  'Success
   SearchAndReplace_AddItem = True

End Function

'On Enter Do Add Item
Private Sub List_Fn_Inc_KeyPress(KeyAscii As Integer)
   
   Select Case KeyAscii
      Case vbKeyReturn
         SearchAndReplace_AddItem
   End Select

End Sub
Private Sub List_Fn_Org_KeyPress(KeyAscii As Integer)
   List_Fn_Inc_KeyPress KeyAscii
End Sub

'On DoubleClick Do Add Item
Private Sub List_Fn_Org_DblClick()
   List_Fn_Inc_DblClick
End Sub
Private Sub List_Fn_Inc_DblClick()
   SearchAndReplace_AddItem
End Sub


'Send Selected Text to Org-SearchSync Textbox
Private Sub Send_TxtFn_Org_TO_Txt_SearchSync()
   With Txt_Fn_Org
      If .SelText <> "" Then
         Txt_SearchSync_Org = .SelText
      End If
   End With
End Sub



'Send Selected Text to Inc SearchSync Textbox
Private Sub Send_TxtFn_Inc_TO_Txt_SearchSync()
   With Txt_Fn_Inc
      If .SelText <> "" Then
         Txt_SearchSync = .SelText
      End If
   End With
End Sub
Private Sub cmd_inc_reload_Click()
   Txt_Fn_Inc_FileName_Change
End Sub


Private Sub List_Fn_String_Org_Click()
   If List_Fn_String_Org_EventBlocker Then Exit Sub
    Txt_SearchSync_Org = List_Fn_String_Org.Text
End Sub

Private Sub Txt_Fn_Inc_FileName_Change()

On Error GoTo Txt_Fn_Inc_FileName_Change_Err

   If FileExists(Txt_Fn_Inc_FileName) Then
      
      OpenAndFill Txt_Fn_Inc_FileName, Script_Inc, Functions_Inc, List_Fn_Inc

'   ElseIf DirExists(Txt_Fn_Inc_FileName) Then
   
   Else
'      Txt_Fn_Inc_FileName.SetFocus
   End If
   
   Dim FileName As New ClsFilename
   With FileName
      FileName = Txt_Fn_Inc_FileName
      File_Includes.Path = .Path
      ListBox_FindAndSelectedItem File_Includes, .NameWithExt
      
      Txt_Include = .NameWithExt
   End With

Err.Clear
Txt_Fn_Inc_FileName_Change_Err:
Select Case Err
   Case 0

   Case Else
      MsgBox Err.Description, vbCritical, "Error " & Hex(Err.Number) & "  in Formular FrmFuncRename.Txt_Fn_Inc_FileName_Change()"

End Select
End Sub

Private Sub Txt_Fn_Inc_FileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Txt_Fn_Inc_FileName_OLEDragDrop_err
   
   Txt_Fn_Inc_FileName = Data.Files(1)
'   Timer_OleDrag.Enabled = True
   

Txt_Fn_Inc_FileName_OLEDragDrop_err:
Select Case Err
Case 0

Case Else
'   log "-->Drop'n'Drag ERR: " & Err.Description

End Select


End Sub

Private Sub Txt_Fn_Inc_KeyUp(KeyCode As Integer, Shift As Integer)
   Send_TxtFn_Inc_TO_Txt_SearchSync
End Sub
Private Sub Txt_Fn_Inc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Send_TxtFn_Inc_TO_Txt_SearchSync
End Sub

Private Function GetNumOccurrenceIn&(SearchText$, FindStr$)
   Do
      Dim FoundPos&
      FoundPos = InStr(FoundPos + 1, SearchText, FindStr, vbTextCompare)
      If FoundPos = 0 Then Exit Do
      Inc GetNumOccurrenceIn
   Loop While FoundPos
   
End Function



Private Sub cmd_org_reload_Click()
   Txt_Fn_Org_FileName_Change
End Sub

Private Sub Txt_Fn_Org_FileName_Change()
  
   If FileExists(Txt_Fn_Org_FileName) Then
      File_Org_FileName = Txt_Fn_Org_FileName
   
      OpenAndFill File_Org_FileName.FileName, Script_Org, Functions_Org, List_Fn_Org
      
      List_Fn_Assigned.Clear
      Set List_Fn_Assigned_FuncIdxs = New Collection
      
'   Else
'      Txt_Fn_Org_FileName.SetFocus
   End If

End Sub

Private Sub Txt_Fn_Org_FileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Txt_Fn_Org_FileName_OLEDragDrop_err
   
   Txt_Fn_Org_FileName = Data.Files(1)
'   Timer_OleDrag.Enabled = True
   

Txt_Fn_Org_FileName_OLEDragDrop_err:
Select Case Err
Case 0

Case Else
'   log "-->Drop'n'Drag ERR: " & Err.Description

End Select

End Sub


Private Sub Txt_Fn_Org_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Txt_Fn_Org.SelStart Then
      Txt_Fn_Org.SetFocus
   End If
End Sub

Private Sub SeekToSearchString(SearchText$, FnList As Listbox, FnData As Textbox, Optional StartAtFunction = 0)
   With FnList
   
      Dim ListIndex_Backup%
      ListIndex_Backup = .ListIndex
      
      Dim Func_Current
      For Func_Current = StartAtFunction To .ListCount - 1
         .ListIndex = Func_Current
         Dim Found_At_Pos
         With FnData
            Found_At_Pos = InStr(1, .Text, SearchText, vbTextCompare)
            If Found_At_Pos Then
               
               'Scroll to the end of Textbox
               .SelStart = Len(.Text)
               
               .SelStart = Found_At_Pos - 1
               .SelLength = Len(SearchText)
               
'               .SetFocus
               Exit Sub
            End If
         End With
            
      Next
      
      .ListIndex = ListIndex_Backup
      
   End With

End Sub


Private Sub Txt_Fn_Org_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Send_TxtFn_Org_TO_Txt_SearchSync
End Sub

Private Sub Txt_Fn_Org_KeyUp(KeyCode As Integer, Shift As Integer)
   Send_TxtFn_Org_TO_Txt_SearchSync
End Sub

Private Sub Txt_SearchSync_Org_Change()
   Lbl_SearchSyncStatus_Org = ""
   
 ' Go through all include files
   With File_Includes
      Dim old_ListIndex%
      old_ListIndex = .ListIndex
      
      Dim i&
      For i = 0 To .ListCount - 1
         SearchSync Txt_SearchSync_Org.Text, _
                    Functions_Inc, List_Fn_Inc, _
                    Txt_Fn_Inc, _
                    Cmd_FindNext_Org, Lbl_SearchSyncStatus_Inc
       ' Exit loop if something found
         If NumOccurrenceFound > 0 Then Exit For
         
       ' Seek to next include file
         .ListIndex = i

      Next

         
      If NumOccurrenceFound > 0 Then
      ' Okay found something - Try AutoAdd
         cmd_AutoAdd_Click
      Else
        ' Nothing found something - Restore old ListPosition
         .ListIndex = old_ListIndex
      End If

   End With

End Sub

Private Sub Txt_SearchSync_Change()
   Lbl_SearchSyncStatus_Inc = ""
   
   SearchSync Txt_SearchSync.Text, _
              Functions_Org, List_Fn_Org, _
              Txt_Fn_Org, _
              Cmd_FindNext_Inc, Lbl_SearchSyncStatus_Org
End Sub
Private Sub SearchSync(SearchText$, FuncList, Fn_List As Listbox, Fn_Data As Textbox, FindNext As CommandButton, Status As Label)

   If SearchText = "" Then Exit Sub
   
   Dim SearchBuffer As clsStrCat
   Set SearchBuffer = New clsStrCat
   
   Dim item
   For item = 0 To Fn_List.ListCount - 1
    ' That is not optimal and might slow down speed
      SearchBuffer.ConcatVariant FuncList(Fn_List.ItemData(item))
   Next
   
   NumOccurrenceFound = GetNumOccurrenceIn(SearchBuffer.value, SearchText)
   
   FindNext.Visible = (NumOccurrenceFound > 1)
   
   If NumOccurrenceFound = 0 Then
      Status = "not found."
      
   ElseIf NumOccurrenceFound = 1 Then
      Status = "found."
      SeekToSearchString SearchText, Fn_List, Fn_Data
   Else
      Status = "found at " & NumOccurrenceFound & " locations."
      SeekToSearchString SearchText, Fn_List, Fn_Data
   End If
   
End Sub
Private Sub Cmd_FindNext_Inc_Click()
   SeekToSearchString Txt_SearchSync.Text, List_Fn_Org, Txt_Fn_Org, _
                      List_Fn_Org.ListIndex + 1
End Sub
Private Sub Cmd_FindNext_Org_Click()
   SeekToSearchString Txt_SearchSync_Org.Text, List_Fn_Inc, Txt_Fn_Inc, _
                      List_Fn_Inc.ListIndex + 1

End Sub

