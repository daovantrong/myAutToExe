VERSION 5.00
Begin VB.Form FrmFuncRename_StringMismatch 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "org"
   ClientHeight    =   4836
   ClientLeft      =   2448
   ClientTop       =   2436
   ClientWidth     =   12084
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4836
   ScaleWidth      =   12084
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_inc 
      Height          =   1452
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FrmFuncRename_StringMismatch.frx":0000
      Top             =   480
      Width           =   5772
   End
   Begin VB.TextBox txt_org 
      Height          =   1452
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "FrmFuncRename_StringMismatch.frx":0006
      Top             =   480
      Width           =   5772
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Press 'esc' to reject this pair of functions"
      Top             =   0
      Width           =   6015
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Press 'enter' to accept this"
      Top             =   0
      Width           =   5895
   End
   Begin VB.ListBox List_Inc 
      Appearance      =   0  'Flat
      Height          =   6744
      ItemData        =   "FrmFuncRename_StringMismatch.frx":000C
      Left            =   6120
      List            =   "FrmFuncRename_StringMismatch.frx":000E
      TabIndex        =   1
      Top             =   2040
      Width           =   5895
   End
   Begin VB.ListBox List_Org 
      Appearance      =   0  'Flat
      Height          =   6744
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   5775
   End
End
Attribute VB_Name = "FrmFuncRename_StringMismatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_fn_org As MatchCollection
Private m_fn_inc As MatchCollection

Public fn_org_idx As Collection
Public fn_inc_idx  As Collection


Private Const List_ID_Text_Sep = ": "

Public Enum AcceptResult_enum
   Result_True
   Result_False
   Result_Undefined
End Enum

Private mAcceptResult As AcceptResult_enum
Public Property Get AcceptResult() As AcceptResult_enum
   AcceptResult = mAcceptResult
End Property



Public Sub create( _
   fn_org As MatchCollection, fn_inc As MatchCollection, _
   Optional str_org As String, Optional str_inc As String, _
   Optional var_org As Collection, Optional var_inc As Collection)
   
   
   Set m_fn_org = fn_org
   Set m_fn_inc = fn_inc
   
   txt_inc.text = str_inc
   txt_org.text = str_org
   
   Set fn_org_idx = var_org
   Set fn_inc_idx = var_inc
   
   
   
   FillList List_Org, fn_org
   FillList List_Inc, fn_inc
   
   mAcceptResult = Result_Undefined
End Sub

Private Sub FillList(List As Listbox, match As MatchCollection)
   List.Clear
   
   Dim counter&
   counter = 0
   
   Dim Line As New clsStrCat
   Dim i As match
   For Each i In match
      Line.Concat Format(counter, "00") & List_ID_Text_Sep
      'line.Concat i.SubMatches(0) & vbTab
      Line.Concat i.value
      'line.Concat vbTab & vbTab i.FirstIndex
      
      List.AddItem Line.value
      
      Line.Clear
      Inc counter
   Next
End Sub

Private Sub cmd_cancel_Click()
   Unload Me
   mAcceptResult = Result_False
End Sub

Private Sub cmd_ok_Click()

'   Set fn_inc = GetResults(List_Inc, m_fn_inc)
'   Set fn_org = GetResults(List_Org, m_fn_org)


   Unload Me
   mAcceptResult = Result_True
End Sub

Private Sub List_Inc_DblClick()
On Error GoTo List_Inc_DblClick_err
   With List_Inc
      Dim idx&
      idx = .ListIndex
      
    ' remove on index list as well
      fn_inc_idx.Remove idx
      
      .RemoveItem idx
      
   End With
   
   cmd_ok.enabled = (List_Inc.ListCount = List_Org.ListCount)
   
List_Inc_DblClick_err:
End Sub

Private Function GetID(List As Listbox) As Integer
   With List
      GetID = Split(.text, List_ID_Text_Sep)(0)
   End With
End Function


Private Sub List_Org_Click()
   ShowInSource _
      m_fn_org(GetID(List_Org)), _
      txt_org

End Sub

Private Sub List_Inc_Click()
   ShowInSource _
      m_fn_inc(GetID(List_Inc)), _
      txt_inc
End Sub
Private Sub ShowInSource(item As match, text As TextBox)
   On Error Resume Next
   With text
      .SelStart = item.FirstIndex
      .SelLength = item.Length
      .SetFocus
   End With
End Sub

'Private Function GetResults( _
'   List As Listbox, _
'   match As MatchCollection) As Collection
'
'   Dim Collection As New Collection
'
'   With List
'      Dim i&, matchID&
'      For i = 0 To .ListCount - 1
'         .ListIndex = i
'         matchID = GetID(List)
'         Collection.add match(matchID).value
'      Next
'   End With
'
'   Set GetResults = Collection
'End Function

