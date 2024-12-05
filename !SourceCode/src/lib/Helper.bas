Attribute VB_Name = "Helper"
Option Explicit
Option Compare Text

Public Cancel As Boolean


'Konstantendeklationen für Registry.cls

'Registrierungsdatentypen
Public Const REG_SZ As Long = 1                         ' String
Public Const REG_BINARY As Long = 3                     ' Binär Zeichenfolge
Public Const REG_DWORD As Long = 4                      ' 32-Bit-Zahl

'Vordefinierte RegistrySchlüssel (hRootKey)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0


Public Const ERR_FILESTREAM = &H1000000
Public Const ERR_OPENFILE = vbObjectError Or ERR_FILESTREAM + 1
Private i, j As Integer

Public Declare Sub MemCopyAnyToAny Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Any, src As Any, ByVal Length&)
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal src As Any, ByVal Length&)
Public Declare Sub MemCopyAnyToStr Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal Length&)
Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As String, src As Long, ByVal Length&)

Public Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (Dest As Long, ByVal src As String, ByVal Length&)
'Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As String, src As Long, ByVal Length&)
Public Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (Dest As Long, ByVal src As Integer, ByVal Length&)
    
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SM_DBCSENABLED = 42
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer


Private BenchtimeA&, BenchtimeB&

'for mt_MT_Init to do a multiplation without 'overflow error'
Private Declare Function iMul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal dw2 As Long, ByVal dw3 As Long, ByVal dw4 As Long) As Long


'Ensure that 'myObjRegExp.MultiLine = True' else it will use the beginning of the string!
Public Const RE_Anchor_LineBegin$ = "^"
Public Const RE_Anchor_LineEnd$ = "$"

Public Const RE_Anchor_WordBoarder$ = "\b"
Public Const RE_Anchor_NoWordBoarder$ = "\B"
Public Const RE_AnyChar$ = "."



Function MulInt32&(a&, b&)
  MulInt32 = iMul(a, 0, b, 0)
End Function

Function AddInt32&(a As Double, b As Double)
  AddInt32 = "&h" & H32(a + b)
End Function


'Returns whether the user has DBCS enabled
Private Function isDBCSEnabled() As Boolean
   isDBCSEnabled = GetSystemMetrics(SM_DBCSENABLED)
End Function


Function LeftButton() As Boolean
    LeftButton = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Function RightButton() As Boolean
    RightButton = (GetAsyncKeyState(vbKeyRButton) And &H8000)
End Function

Function MiddleButton() As Boolean
    MiddleButton = (GetAsyncKeyState(vbKeyMButton) And &H8000)
End Function

Function MouseButton() As Integer
    If GetAsyncKeyState(vbKeyLButton) < 0 Then
        MouseButton = 1
    End If
    If GetAsyncKeyState(vbKeyRButton) < 0 Then
        MouseButton = MouseButton Or 2
    End If
    If GetAsyncKeyState(vbKeyMButton) < 0 Then
        MouseButton = MouseButton Or 4
    End If
End Function

Function KeyPressed(Key) As Boolean
   KeyPressed = GetAsyncKeyState(Key)
End Function

Public Function HexStringToString$(ByVal HexString$)
   HexStringToString = Space(Len(HexString) \ 2)
   For i = 1 To Len(HexString) Step 2
      Mid$(HexStringToString, (i \ 2) + 1) = Chr("&h" & Mid$(HexString, i, 2))
   Next
End Function

Public Function HexvaluesToString$(Hexvalues$)
   Dim tmpchar
   For Each tmpchar In Split(Hexvalues)
      'HexvaluesToString = HexvaluesToString & ChrB("&h" & tmpchar) & ChrB(0)
      'Note ChrB("&h98") & ChrB(0) is not correct translated
      HexvaluesToString = HexvaluesToString & Chr("&h" & tmpchar)
   Next
End Function

Public Function ValuesToHexString$(Data As StringReader, Optional seperator = " ")
'ValuesToHexString = ""
   With Data
      .EOS = False
      Do Until .EOS
         ValuesToHexString = ValuesToHexString & H8(.int8) & seperator
      Loop
   End With
  
End Function


Function Max(ParamArray values())
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function Min(ParamArray values())
   Dim item
   Min = &H7FFFFFFF
   For Each item In values
      Min = IIf(Min > item, item, Min)
   Next
End Function

Function limit(value&, Optional ByVal upperLimit = &H7FFFFFFF, Optional lowerLimit = 0) As Long
   'limit = IIf(Value > upperLimit, upperLimit, IIf(Value < lowerLimit, lowerLimit, Value))

   If (value > upperLimit) Then _
      limit = upperLimit _
   Else _
      If (value < lowerLimit) Then _
         limit = lowerLimit _
      Else _
         limit = value
   
End Function

Function isEven(Number As Long) As Boolean
   isEven = ((Number And 1) = 0)
End Function

Function RangeCheck(ByVal value&, Max&, Optional Min& = 0, Optional ErrText, Optional ErrSource$) As Boolean
   RangeCheck = (Min <= value) And (value <= Max)
   If (RangeCheck = False) And (IsMissing(ErrText) = False) Then _
       Err.Raise vbObjectError, ErrSource, _
           ErrText & " Value must between '" & Min & "'  and '" & Max & "' !"
End Function

Public Function H8(ByVal value As Long)
   H8 = Right(String(1, "0") & Hex(value), 2)
End Function

Public Function H16(ByVal value As Long)
   H16 = Right(String(3, "0") & Hex(value), 4)
End Function

Public Function H32(ByVal value As Double)
   If value <= &H7FFFFFFF Then
      H32 = Hex(value)
   Else
    ' split Number in High a Low part...
      Dim High&, Low&
      High = Int(value / &H10000)
      Low = value - (CDbl(High) * &H10000)
      
      H32 = H16(High) & H16(Low)
   End If
   
   H32 = Right(String(7, "0") & H32, 8)
End Function

Public Function Swap(ByRef a, ByRef b)
   Swap = b
   b = a
   a = Swap
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_r  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_l("Summe",4) -> "umme"
Public Function BlockAlign_r(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Right(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_r = RawString & Space(Blocksize - Len(RawString))
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_l  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_l("Summe",4) -> "umme"
Public Function BlockAlign_l(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Left(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_l = Space(Blocksize - Len(RawString)) & RawString
End Function

Public Function qw()
   Cancel = True
   Do
      DoEvents
   Loop While Cancel = True
End Function
Public Function szNullCut$(zeroString$)
   Dim nullCharPos&
   nullCharPos = InStr(1, zeroString, Chr(0))
   If nullCharPos Then
      szNullCut = Left(zeroString, nullCharPos - 1)
   Else
      szNullCut = zeroString
   End If
   
End Function


Public Function Inc(ByRef value, Optional Increment& = 1)
   value = value + Increment
   Inc = value
End Function

Public Function Dec(ByRef value, Optional DeIncrement& = 1)
   value = value - DeIncrement
   Dec = value
End Function



Public Function CollectionToArray(Collection As Collection) As Variant
   
   Dim tmp
   ReDim tmp(Collection.Count - 1)
   
   Dim i
   i = LBound(tmp)
   
   Dim item
   For Each item In Collection
      tmp(i) = item
      Inc i
   Next
   
   CollectionToArray = tmp
   
End Function
Public Function isString(StringToCheck) As Boolean
   'isString = False
   Dim i&
   For i = 1 To Len(StringToCheck)
      If RangeCheck(Asc(Mid$(StringToCheck, i, 1)), &H7F, &H20) Then
      
      Else
         Exit Function
      End If
   Next
   
   isString = True
   
End Function



'Searches for some string and then starts there to crop
Function strCropWithSeek$(Text$, LeftString$, RightString$, Optional errorvalue, Optional SeektoStrBeforeSearch$)
   strCropWithSeek = strCrop1(Text$, LeftString$, RightString$, errorvalue, _
            InStr(1, Text, SeektoStrBeforeSearch))
End Function


Function strCrop1$(ByVal Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1)
   
   Dim cutend&, cutstart&
      cutstart = InStr(StartSearchAt, Text, LeftString)
   If cutstart Then
      cutstart = cutstart + Len(LeftString)
      cutend = InStr(cutstart, Text, RightString)
      If cutend > cutstart Then
         strCrop1 = Mid$(Text, cutstart, cutend - cutstart)
      Else
        'is Rightstring empty?
         If RightString = "" Then
            strCrop1 = Mid$(Text, cutstart)
         Else
            strCrop1 = errorvalue
         End If
      End If
   Else
      strCrop1 = errorvalue
   End If

End Function

Function strCropAndDelete(Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1, Optional ReplaceString$ = "")
   strCropAndDelete = strCrop1(Text$, LeftString$, RightString$, errorvalue, StartSearchAt)
   Text = Replace(Text, LeftString & strCropAndDelete & RightString, ReplaceString, , , vbTextCompare)
End Function



Function strCrop$(Text$, LeftString$, RightString$, Optional errorvalue, Optional StartSearchAt = 1)
   
   Dim cutend&, cutstart&
      cutend = InStr(StartSearchAt, Text, RightString)
   If cutend Then
      cutstart = InStrRev(Text, LeftString, cutend, vbBinaryCompare) + Len(LeftString)
      strCrop = Mid$(Text, cutstart, cutend - cutstart)
   Else
      strCrop = errorvalue
   End If

End Function

Function MidMbcs(ByVal Str As String, Start, Length)
    MidMbcs = StrConv(MidB$(StrConv(Str, vbFromUnicode), Start, Length), vbUnicode)
End Function


Function strCutOut$(Str$, pos&, Length&, Optional TextToInsert = "")
   strCutOut = Mid(Str, pos, Length)
   Str$ = Mid(Str, 1, pos - 1) & TextToInsert & Mid(Str, pos + Length)
End Function


Public Function Int16ToUInt32&(value%)
      Const N_0x8000& = 32767
      If value >= 0 Then
         Int16ToUInt32 = value
      Else
         Int16ToUInt32 = CLng(value And N_0x8000) + N_0x8000
      End If
      
End Function




Public Function BenchStart()

   BenchtimeA = GetTickCount

End Function
Public Function BenchEnd()

   BenchtimeB = GetTickCount
   Debug.Print Time & " - " & BenchtimeB - BenchtimeA

End Function


Public Function FileExists(FileName) As Boolean
   On Error GoTo FileExists_err
   FileExists = FileLen(FileName)

FileExists_err:
End Function

Public Function Quote(ByRef Text As String) As String
   Quote = """" & Text & """"
End Function

Public Function Brackets(ByRef Text As String) As String
   Brackets = "(" & Text & ")"
End Function

Public Function RE_WSpace(ParamArray Elements()) As String
   Dim WS$ ' WhiteSpace
   WS = "\s*"
   
   RE_WSpace = Join(Elements, WS)
End Function



Public Function RE_LookHead_positive(ExpressionThatShouldBeFound$) As String
   RE_LookHead_positive = "(?=" & ExpressionThatShouldBeFound & ")"
End Function

Public Function RE_LookHead_negative(ExpressionThatShouldNOTBeFound$) As String
   RE_LookHead_negative = "(?!" & ExpressionThatShouldNOTBeFound & ")"
End Function

Public Function RE_Repeat(Optional MinRepeat& = 0, Optional MaxRepeat = "") As String
   If (MinRepeat = MaxRepeat) Then
      RE_Repeat = "{" & MinRepeat & "}"
   Else
      RE_Repeat = "{" & MinRepeat & "," & MaxRepeat & "}"
   End If
   
End Function


Public Function RE_AnyCharRepeat(Optional MinRepeat& = 0, Optional MaxRepeat = "") As String
   RE_AnyCharRepeat = "." & RE_Repeat(MinRepeat, MaxRepeat)
End Function

Public Function RE_Group(RegExpForTheGroup$) As String
   RE_Group = "(" & RegExpForTheGroup & ")"
End Function

Public Function RE_Group_NonCaptured(RegExpForTheNonCapturedGroup$) As String
   RE_Group_NonCaptured = "(?:" & RegExpForTheNonCapturedGroup & ")"
End Function

Public Function RE_Literal(TextWithLiterals) As String
   'Mask metachars
   RE_Literal = RE_Mask(TextWithLiterals, "][{}()*+?.\\^$|")
                                           
End Function

Private Function RE_Mask(ByVal Text, CharsToMask$) As String
'   Dim t As New RegExp
   With New RegExp
      .Global = True
      
     ' Unmask it first to avoid double mask
      .Pattern = "\\" & _
                  "([" & CharsToMask & "])"
       Text = .Replace(Text, "$1")
   
     ' Mask MetaChars like with a preciding '\'
      .Pattern = "[" & CharsToMask & "]"
      RE_Mask = .Replace(Text, "\$&")
   
   
   End With

End Function

Public Function RE_CharCls(Chars$) As String
   ' mask ']' and '-'
   RE_CharCls = "[" & RE_Mask(Chars, "]\\-") & "]"
End Function

Public Function RE_CharCls_Excluded(Chars$) As String
   ' mask ']' and '-'
   RE_CharCls_Excluded = "[^" & RE_Mask(Chars, "]\\-") & "]"

End Function

Public Function IsAlreadyInCollection(CollectionToTest As Collection, Key$) As Boolean
   Dim Description$, Number&, Source$
   Description = Err.Description
   Number = Err.Number
   Source = Err.Source
   
      On Error Resume Next
      CollectionToTest.item Key
      IsAlreadyInCollection = (Err = 0)
      
   Err.Description = Description
   Err.Number = Number
   Err.Source = Source


End Function

