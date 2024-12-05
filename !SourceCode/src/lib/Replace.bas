Attribute VB_Name = "mod_Replace"
Option Explicit

Public Function replace(ByRef Text, _
    ByRef sOld As String, ByRef sNew As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Count As Long = 2147483647, _
    Optional ByVal compare As VbCompareMethod = vbBinaryCompare _
  ) As String

  If LenB(sOld) = 0 Then

    'Suchstring ist leer:
    replace = Text

  Else
     
    'Do necessary init
     If Len(sOld) = Len(sOld) Then
        replace = Text
     End If
     
     If ContainsOnly0(sOld) Then
   
       'Unicode-Problem, also kein LenB und co. verwenden:
       ReplaceBin0 replace, Text, Text, sOld, sNew, Start, Count
   
     Else
   
       'Groﬂ/Kleinschreibung unterscheiden:
       ReplaceBin replace, Text, Text, sOld, sNew, Start, Count, compare
   
     End If

   End If
End Function

Public Sub ReplaceDoMulti(ByRef Text As String, _
    ByRef sOld As String, ByRef sNew As String, _
    Optional ByVal compare As VbCompareMethod = vbBinaryCompare)
  
  Text = Join(Split(Text, sOld, , compare), sNew)
  
End Sub


' Attention: QuickReplace
' Tries to keeps the size by filling up the New String with spaces
' if Replace New String is bigger -> Data will grow => what is slow
' => Check is that kind of replace is appropriate before use!
Public Sub QuickReplace(Data$, ReplaceStr, ByVal NewStr$, Optional startPos& = 1, Optional Count& = 2147483647)
   StringFillUp NewStr, Len(ReplaceStr)
   
   ReplaceDo Data, ReplaceStr, NewStr, startPos, Count
'   '-
'   'Do Replacing
'    If Len(NewStr) > NewLen Then
'       'Extend String
'       Stop
'       ReplaceDo Data, oldValue, NewStr, NewPos
'    Else
'       Mid(Data, NewPos, NewLen) = BlockAlign_l(NewStr, NewLen)
'    End If

End Sub

Public Sub ReplaceDo(ByRef Text, _
    ByRef sOld, ByRef sNew, _
    Optional ByVal Start As Long = 1, _
    Optional ByRef Count As Long = 2147483647, _
    Optional ByVal compare As VbCompareMethod = vbBinaryCompare _
  )

  If LenB(sOld) = 0 Then

    'Suchstring ist leer: Nix machen!
    Count = 0

  ElseIf ContainsOnly0(sOld) Then

    'Unicode-Problem, also kein LenB und co. verwenden:
    ReplaceBin0 Text, Text, Text, sOld, sNew, Start, Count

  Else

    'Groﬂ/Kleinschreibung unterscheiden:
'    If InStr(Start, Text, sOld, vbBinaryCompare) Then
    ReplaceBin Text, Text, Text, sOld, sNew, Start, Count, compare
 '   Else Count = 0
'      End If

  End If

End Sub

Private Function ContainsOnly0(ByRef s) As Boolean

  Dim i As Long

  For i = 1 To Len(s)
    If Asc(Mid$(s, i, 1)) Then Exit Function
  Next i
  ContainsOnly0 = True

End Function


Private Static Sub ReplaceBin(ByRef result, _
    ByRef Text, ByRef search, _
    ByRef sOld, ByRef sNew, _
    ByVal Start As Long, ByRef Count As Long, _
    ByRef CompareMethode)

   Dim OldLen&, NewLen&
   OldLen = Len(sOld)
   NewLen = Len(sNew)
   
   Select Case NewLen
   
      Case OldLen  'einfaches ‹berschreiben:
         
         If CompareMethode = vbTextCompare Then
         
         'Convert to uppercase to do an case insensitve compare
          Dim U_Search$, U_sOld$
          U_Search = UCase$(search)
          U_sOld = UCase$(sOld)
          
            For Count = 1 To Count
              
            ' N‰chsten Treffer bestimmen:
              Start = InStr(Start, U_Search, U_sOld)
              
            ' Wenn kein weiter Treffer - Schleife verlassen
              If Start = 0 Then Exit For
              
            ' Patch Result
              Mid$(result, Start) = sNew
              
              Start = Start + OldLen

              
            Next Count
       
      Else
            
            For Count = 1 To Count
              
            ' N‰chsten Treffer bestimmen:
              Start = InStr(Start, search, sOld)
              
            ' Wenn kein weiter Treffer - Schleife verlassen
              If Start = 0 Then Exit For
              
              If sOld <> sNew Then
               ' Patch Result
                 Mid$(result, Start) = sNew
              End If
              
              Start = Start + OldLen

              
            Next Count
         End If
            
      Case Else 'Zerlegen und wiederzusammensetzen
         
       
       ' Bei SuchText Zerlegen
         Dim Text_Splited
         Text_Splited = Split(Text, sOld, , CompareMethode)
       
       ' ... Treffer Z‰hlen ...
         Count = UBound(Text_Splited)
         
       ' ... und wiederzusammensetzen :)
         result = Join(Text_Splited, sNew)
         
   End Select

End Sub

Private Static Sub OBSOLATED_ReplaceBin(ByRef result, _
    ByRef Text, ByRef search, _
    ByRef sOld, ByRef sNew As String, _
    ByVal Start As Long, ByRef Count As Long _
  )

  Dim TextLen As Long
  Dim OldLen As Long
  Dim NewLen As Long
  Dim ReadPos As Long
  Dim WritePos As Long
  Dim CopyLen As Long
  Dim Buffer As String
  Dim BufferLen As Long
  Dim BufferPosNew As Long
  Dim BufferPosNext As Long


  'Ersten Treffer bestimmen:
  If Start < 2 Then
    Start = InStrB(search, sOld)
  Else
    Start = InStrB(Start + Start - 1, search, sOld)
  End If
  If Start Then

    OldLen = LenB(sOld)
    NewLen = LenB(sNew)
    Select Case NewLen
    Case OldLen 'einfaches ‹berschreiben:

      result = Text
      For Count = 1 To Count
        MidB$(result, Start) = sNew
        Start = InStrB(Start + OldLen, search, sOld)
        If Start = 0 Then Exit Sub
      Next Count
      Exit Sub


    Case Is < OldLen 'Ergebnis wird k¸rzer:

      'Buffer initialisieren:
      TextLen = LenB(Text)
      If TextLen > BufferLen Then
        Buffer = Text
        BufferLen = TextLen
      End If

      'Ersetzen:
      ReadPos = 1
      WritePos = 1
      If NewLen Then

        'Einzuf¸genden Text beachten:
        For Count = 1 To Count
          CopyLen = Start - ReadPos
          If CopyLen Then
            BufferPosNew = WritePos + CopyLen
            MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
            MidB$(Buffer, BufferPosNew) = sNew
            WritePos = BufferPosNew + NewLen
          Else
            MidB$(Buffer, WritePos) = sNew
            WritePos = WritePos + NewLen
          End If
          ReadPos = Start + OldLen
          Start = InStrB(ReadPos, search, sOld)
          If Start = 0 Then Exit For
        Next Count

      Else

        'Einzuf¸genden Text ignorieren (weil leer):
        For Count = 1 To Count
          CopyLen = Start - ReadPos
          If CopyLen Then
            MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
            WritePos = WritePos + CopyLen
          End If
          ReadPos = Start + OldLen
          Start = InStrB(ReadPos, search, sOld)
          If Start = 0 Then Exit For
        Next Count

      End If

      'Ergebnis zusammenbauen:
      If ReadPos > TextLen Then
        result = LeftB$(Buffer, WritePos - 1)
      Else
        MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
        result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
      End If
      Exit Sub

    Case Else 'Ergebnis wird l‰nger:

      'Buffer initialisieren:
      TextLen = LenB(Text)
      BufferPosNew = TextLen + NewLen
      If BufferPosNew > BufferLen Then
        Buffer = Space$(BufferPosNew)
        BufferLen = LenB(Buffer)
      End If

      'Ersetzung:
      ReadPos = 1
      WritePos = 1
      For Count = 1 To Count
        CopyLen = Start - ReadPos
        If CopyLen Then
          'Positionen berechnen:
          BufferPosNew = WritePos + CopyLen
          BufferPosNext = BufferPosNew + NewLen

          'Ggf. Buffer vergrˆﬂern:
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = LenB(Buffer)
          End If

          'String "patchen":
          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
          MidB$(Buffer, BufferPosNew) = sNew
        Else
          'Position bestimmen:
          BufferPosNext = WritePos + NewLen

          'Ggf. Buffer vergrˆﬂern:
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = LenB(Buffer)
          End If

          'String "patchen":
          MidB$(Buffer, WritePos) = sNew
        End If
        WritePos = BufferPosNext
        ReadPos = Start + OldLen
        Start = InStrB(ReadPos, search, sOld)
        If Start = 0 Then Exit For
      Next Count

      'Ergebnis zusammenbauen:
      If ReadPos > TextLen Then
        result = LeftB$(Buffer, WritePos - 1)
      Else
        BufferPosNext = WritePos + TextLen - ReadPos
        If BufferPosNext < BufferLen Then
          MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
          result = LeftB$(Buffer, BufferPosNext)
        Else
          result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
        End If
      End If
      Exit Sub

    End Select

  Else 'Kein Treffer:
    result = Text
    Count = 0
  End If

End Sub

Private Static Sub ReplaceBin0(ByRef result, _
    ByRef Text, ByRef search, _
    ByRef sOld, ByRef sNew, _
    ByVal Start As Long, ByVal Count As Long _
  )

  Dim TextLen As Long
  Dim OldLen As Long
  Dim NewLen As Long
  Dim ReadPos As Long
  Dim WritePos As Long
  Dim CopyLen As Long
  Dim Buffer As String
  Dim BufferLen As Long
  Dim BufferPosNew As Long
  Dim BufferPosNext As Long

  'Ersten Treffer bestimmen:
  If Start < 2 Then
    Start = InStr(search, sOld)
  Else
    Start = InStr(Start, search, sOld)
  End If
  
  If Start Then
  
    OldLen = Len(sOld)
    NewLen = Len(sNew)
    Select Case NewLen
    Case OldLen 'einfaches ‹berschreiben:
    
      result = Text
      For Count = 1 To Count
        Mid$(result, Start) = sNew
        Start = InStr(Start + OldLen, search, sOld)
        If Start = 0 Then Exit Sub
      Next Count
      Exit Sub
    
    Case Is < OldLen 'Ergebnis wird k¸rzer:
    
      'Buffer initialisieren:
      TextLen = Len(Text)
      If TextLen > BufferLen Then
        Buffer = Text
        BufferLen = TextLen
      End If
      
      'Ersetzen:
      ReadPos = 1
      WritePos = 1
      If NewLen Then
      
        'Einzuf¸genden Text beachten:
        For Count = 1 To Count
          CopyLen = Start - ReadPos
          If CopyLen Then
            BufferPosNew = WritePos + CopyLen
            Mid$(Buffer, WritePos) = Mid$(Text, ReadPos, CopyLen)
            Mid$(Buffer, BufferPosNew) = sNew
            WritePos = BufferPosNew + NewLen
          Else
            Mid$(Buffer, WritePos) = sNew
            WritePos = WritePos + NewLen
          End If
          ReadPos = Start + OldLen
          Start = InStr(ReadPos, search, sOld)
          If Start = 0 Then Exit For
        Next Count
      
      Else
      
        'Einzuf¸genden Text ignorieren (weil leer):
        For Count = 1 To Count
          CopyLen = Start - ReadPos
          If CopyLen Then
            Mid$(Buffer, WritePos) = Mid$(Text, ReadPos, CopyLen)
            WritePos = WritePos + CopyLen
          End If
          ReadPos = Start + OldLen
          Start = InStr(ReadPos, search, sOld)
          If Start = 0 Then Exit For
        Next Count
      
      End If
      
      'Ergebnis zusammenbauen:
      If ReadPos > TextLen Then
        result = Left$(Buffer, WritePos - 1)
      Else
        Mid$(Buffer, WritePos) = Mid$(Text, ReadPos)
        result = Left$(Buffer, WritePos + Len(Text) - ReadPos)
      End If
      Exit Sub
    
    Case Else 'Ergebnis wird l‰nger:
    
      'Buffer initialisieren:
      TextLen = Len(Text)
      BufferPosNew = TextLen + NewLen
      If BufferPosNew > BufferLen Then
        Buffer = Space$(BufferPosNew)
        BufferLen = Len(Buffer)
      End If
      
      'Ersetzung:
      ReadPos = 1
      WritePos = 1
      For Count = 1 To Count
        CopyLen = Start - ReadPos
        If CopyLen Then
          'Positionen berechnen:
          BufferPosNew = WritePos + CopyLen
          BufferPosNext = BufferPosNew + NewLen
          
          'Ggf. Buffer vergrˆﬂern:
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = Len(Buffer)
          End If
          
          'String "patchen":
          Mid$(Buffer, WritePos) = Mid$(Text, ReadPos, CopyLen)
          Mid$(Buffer, BufferPosNew) = sNew
        Else
          'Position bestimmen:
          BufferPosNext = WritePos + NewLen
          
          'Ggf. Buffer vergrˆﬂern:
          If BufferPosNext > BufferLen Then
            Buffer = Buffer & Space$(BufferPosNext)
            BufferLen = Len(Buffer)
          End If
          
          'String "patchen":
          Mid$(Buffer, WritePos) = sNew
        End If
        WritePos = BufferPosNext
        ReadPos = Start + OldLen
        Start = InStr(ReadPos, search, sOld)
        If Start = 0 Then Exit For
      Next Count
      
      'Ergebnis zusammenbauen:
      If ReadPos > TextLen Then
        result = Left$(Buffer, WritePos - 1)
      Else
        BufferPosNext = WritePos + TextLen - ReadPos
        If BufferPosNext < BufferLen Then
          Mid$(Buffer, WritePos) = Mid$(Text, ReadPos)
          result = Left$(Buffer, BufferPosNext)
        Else
          result = Left$(Buffer, WritePos - 1) & Mid$(Text, ReadPos)
        End If
      End If
      Exit Sub
    
    End Select
  
  Else 'Kein Treffer:
    result = Text
  End If

End Sub
'////////////////////////////////////////////////////
'// BatchReplace
'//
'// it     =  data string
'// them   = "changeThis->toThis,andThis->toThat"
'//
'// Usage example:    BatchReplace " 1 - 2 ", "1->2,2->3"
'Public Function BatchReplace(ByRef it$, them, _
'            Optional compare As VbCompareMethod = vbTextCompare) As String
'
'    DoBatchReplace it, them, compare
'    BatchReplace = it
'
'End Function
'
'Public Sub DoBatchReplace(ByRef it$, them, _
'            Optional compare As VbCompareMethod = vbTextCompare)
'    Dim item
'    For Each item In Split(them, ",")
'
'      Dim SnR
'      SnR = Split(item, "->")
'
'      On Error Resume Next
'
'      'it = VBA.replace(it, SnR(0), SnR(1) _
'                              , , , compare)
'
'     ' Replace via String Split/Join
'      'it = Join( _
'               Split(it, SnR(0), , compare), _
'           SnR(1))
'
'
'    Next
'
'
'End Sub
