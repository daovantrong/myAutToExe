VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tc_Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Type Class for Au3 Functions
'
Private Const FUNC_DELIMITER = "EndFunc"
Public FuncData As String


Public Function create(FuncData) As tc_Func
   Me.FuncData = FuncData
   Set create = Me
End Function

Public Property Get FuncName() As String
   FuncName = GetFuncName(FuncData)
End Property

'Public Property Let FuncName(ByVal vNewValue As string)
'
'End Property

Public Static Function splitFunc(Data As String) As Collection

   Set splitFunc = New Collection
   
   Dim item
   For Each item In split(Data, FUNC_DELIMITER, , vbTextCompare)
   
      Dim myfunc As tc_Func
      Set myfunc = New tc_Func
      
    ' Create & Add myfunc object to functions Collection
      splitFunc.Add _
         myfunc.create(item), myfunc.FuncName
      
   Next

End Function


Public Static Function FuncJoin(Func As Collection) As String
   
   ' --- CollectionToArray with Func ---
   Dim tmp
   ReDim tmp(Func.Count - 1)
   
   Dim i
   i = LBound(tmp)
   
   Dim item As tc_Func
   For Each item In Func
      tmp(i) = item.FuncData
      Inc i
   Next
   
   '----
   
   FuncJoin = Join(tmp, FUNC_DELIMITER)

End Function

