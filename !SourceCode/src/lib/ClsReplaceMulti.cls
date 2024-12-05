VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReplaceMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mData$
Private mTree As Collection
Private mReplacements As Collection

Public Sub Init(Data)
'   mData = Data
   
   Set mTree = New Collection
   
   Set mReplacements = New Collection
   
End Sub


Public Sub add(search, replace)
   mReplacements.add replace
   
   treeAdd search, mTree
   
End Sub

Private Sub treeAdd(search, ByRef TreePart)
      Dim chunk
      For Each chunk In TreePart
            treeAdd search, chunk
            'Dim level
            'level = mReplacements.Count ' mReplacements(level)

         
      Next
      If IsObject(chunk) Then
      Else
         Set TreePart = Split(TreePart, search)
      End If

End Sub

Public Function execute()
   execute = Join(mTree, mReplacements(mReplacements.Count))
   execute = "ClsReplaceMulti Not implemented yet!"
End Function
