VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBinaryReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DisableAutoMove As Boolean
Public LoadedFile As String
Public Position As Long
Private m_Length As Long

Private m_Data() As Byte
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal src As Long, ByVal Length&)

Property Get Length() As Long
    Length = m_Length
End Property

Property Get EOS() As Boolean
    EOS = (Position >= m_Length)
End Property

Function Load(fpath) As Boolean
    
    On Error GoTo Load_err
    
    LoadedFile = fpath
    
    Dim hFile As Long
    hFile = FreeFile
    
    Open fpath For Binary As hFile
    
    m_Length = LOF(hFile)
    
    ReDim m_Data(m_Length - 1)
    
    Get hFile, , m_Data()
    
    Close hFile
    
    Position = 0
    
    
    Load = True

Load_err:
End Function
Sub Move(ByVal Size As Long)
    Position = Position + Size
End Sub

Public Property Let ToMove(ByVal vData As Long)
   If DisableAutoMove Then Exit Property
   Move vData
End Property


'should not be a property because ide will query properties
'at runtime if you mouse over them and then it would move pointer!!
Function int32() As Long
   Const Size& = 4
   If Position + Size <= m_Length Then
      MoveMemory int32, VarPtr(m_Data(Position)), Size
      ToMove = Size
   End If
End Function


Function int16() As Integer
   Const Size& = 4
   If Position + Size <= m_Length Then
      MoveMemory int16, VarPtr(m_Data(Position)), Size
      ToMove = Size
   End If
End Function

Function int8() As Byte
   Const Size& = 4
   If Position + Size <= m_Length Then
      MoveMemory int8, VarPtr(m_Data(Position)), Size
      ToMove = Size
   End If
End Function

Private Function FileExists(Path As String) As Boolean
  On Error GoTo FileExists_err
    
   If Len(Path) <> 0 Then _
      If Right(Path, 1) <> "\" Then _
         If Dir(Path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then _
            FileExists = True
  
  Exit Function
FileExists_err: FileExists = False
End Function
