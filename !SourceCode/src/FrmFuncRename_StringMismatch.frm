VERSION 5.00
Begin VB.Form FrmFuncRename_StringMismatch 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Function Renamer String Mismatch"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.ListBox List_Inc 
      Appearance      =   0  'Flat
      Height          =   8610
      ItemData        =   "FrmFuncRename_StringMismatch.frx":0000
      Left            =   4800
      List            =   "FrmFuncRename_StringMismatch.frx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
   Begin VB.ListBox List_Org 
      Appearance      =   0  'Flat
      Height          =   8610
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "FrmFuncRename_StringMismatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Create(fn_org As MatchCollection, fn_inc As MatchCollection)
   FillList List_Org, fn_org
   FillList List_Inc, fn_inc
End Sub

Private Sub FillList(List As Listbox, Match As MatchCollection)
   List.Clear
   
   Dim i As Match
   For Each i In Match
      List.AddItem i '.SubMatches(1)
   Next
End Sub

Private Sub cmd_cancel_Click()
   Unload Me
End Sub

Private Sub cmd_ok_Click()
   Unload Me
End Sub
