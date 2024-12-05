Attribute VB_Name = "mod_AU3_parser"
Option Explicit
Dim RE_HEXSTRING_Group$



'//////////////////////////////////////////////
'// BinaryToString
'
'// Replaces "BinaryToString(0x3031)" - > "02"
'// Returns: Number of Matches
Function BinaryToString(ClsDeobfu As ClsDeobfuscator, Data$, _
      FunctionName, Optional isStringReverse As Boolean) As Long
   
   
  
'   Debug.Assert FunctionName <> "A000021050005VNTGJ9T93X4N4WGE"
  
   If FunctionName Like "Execute" Then Exit Function
   
   Dim matchcount&
   
   Dim myRegExp As New RegExp
   With myRegExp

      .Global = True
      .IgnoreCase = True



      ' ,1 or ,2 or ,3 or ,4
      Dim RE_optional_Parm1$
      RE_optional_Parm1 = _
            RE_Group_NonCaptured( _
                 RE_WSpace(",", RE_Group("[1-4]")) _
            ) & "?"
      
      
      init_HEXSTRING_Group

      
      'BinaryToString("0x537472696E6752657665727365")
      .Pattern = RE_WSpace( _
            FunctionName, "\(", RE_HEXSTRING_Group, _
                RE_optional_Parm1, "\)")
                      
      Dim matches As MatchCollection
      Set matches = .Execute(Data)
      
      Dim counter&
      GUIEvent_ProcessBegin matches.Count
      
      Dim Match As Match
      For Each Match In matches
         
         GUIEvent_ProcessUpdate Inc(counter)
      
         With Match
         
            Dim IsPrintable As Boolean
            Dim BinData$
            On Error Resume Next
            
            BinData = _
               HexStringToString(.SubMatches(0), IsPrintable, .SubMatches(1))
               
            If Err = 0 Then
               
               If isStringReverse Then _
                  BinData = StrReverse(BinData)
                  
               If IsPrintable Then
                  log_verbose "Replacing: " & BinData & " <= " & .value
                  'ReplaceDo data, .value, EncodeUTF8(MakeAutoItString(BinData)), .FirstIndex, 1
                  'ReplaceDoMulti data, .value, EncodeUTF8(MakeAutoItString(BinData))
                  
                  ClsDeobfu.DeObfu_ReplaceStrings Data, .value, EncodeUTF8(BinData)
               Else
                  Log "Skipped replace(not printable): " & MakePrintable(BinData) & " <= " & .value
               End If
               
            End If
               
         End With

      Next
      
      matchcount = matches.Count
      
      GUIEvent_ProcessEnd

      
      
   End With
   
   
   BinaryToString = matchcount
   'Log MatchCount & " refs to '" & FunctionName & "'  transformed."
'   Log "'" & FunctionName & "' deleted.  " & MatchCount & " refs to it transformed."
   

End Function



Function GetFuncName(Func_Data) As String
   
   Dim myRegExp As New RegExp
   With myRegExp

      .IgnoreCase = True
      .Pattern = RE_WSpace("", _
            "Func", RE_Group("\w*"))

      Dim matches As MatchCollection
      Set matches = .Execute(Func_Data)
      If matches.Count < 1 Then
      'Err getting Func Name
'         Stop
      Else
         GetFuncName = matches(0).SubMatches(0)
      End If
   
   End With
   
End Function


Public Function Transform(Data$, FindCMD, FindPattern$, _
   Optional AppendToMatches = "", Optional PrependToMatches = "")
   Dim myRegExp As New RegExp
   With myRegExp
      .IgnoreCase = True
      .Global = True

      .Pattern = FindPattern
      
      Dim matches As MatchCollection
      Set matches = .Execute(Data)
      
      Dim counter&
      GUIEvent_ProcessBegin matches.Count

      Dim Match As Match
      For Each Match In matches
         
         GUIEvent_ProcessUpdate Inc(counter)
      
         With Match
         
            Dim RE_FirstGroup$
            RE_FirstGroup = PrependToMatches & .SubMatches(0) & AppendToMatches
            
            'Log "Replacing: " & FuncName & " <= " & .value
            'ReplaceDo data, .value, RE_FirstGroup, .FirstIndex, 1
            ReplaceDoMulti Data, .value, RE_FirstGroup
            
         End With
      Next
      
      Log matches.Count & " '" & FindCMD & "s' transformed."
      GUIEvent_ProcessEnd
      
   End With

End Function

Private Sub init_HEXSTRING_Group()
      ' Matches "0x22FF44", "22FF44"
      ' but not "0x0x22FF4"

      RE_HEXSTRING_Group = RE_WSpace( _
         RE_Quote & RE_Group_NonCaptured("0x") & "?", _
            RE_Group( _
               RE_Group_NonCaptured(RE_HEXDIGET) & "*" _
            ), _
         RE_Quote)
End Sub

Public Sub Optimise(Data$)
   Dim matchcount&
   
   Dim myRegExp As New RegExp
   With myRegExp

      .Global = True

      init_HEXSTRING_Group
      '----------
            
      'Call opti pass CALL
      'Call("StringIsInt", $TOGKRHFYFGPVTOVOX7) -> StringIsInt ( $TOGKRHFYFGPVTOVOX7))
      'Call("A000017181918TJOOX2I0B51GAUO977M") -> A000017181918TJOOX2I0B51GAUO977M()
      
            Transform Data, "Call", _
               RE_WSpace( _
                  "Call", "\(", _
                  RE_Quote, RE_Group(RE_AU3NAME), RE_Quote, _
                  ",?"), _
               "("
               
      '----------
            
      'Call opti pass Execute("0x7") -> 7
            Transform Data, "Execute", _
               RE_WSpace("Execute", "\(", RE_HEXSTRING_Group, "\)"), _
               "", "0x"
               
      '      .Pattern = RE_WSpace("Execute", "\(", RE_HEXSTRING_Group, "\)")
      '
      '      Set matches = .Execute(Data)
      '
      '      counter = 0
      '      GUIEvent_ProcessBegin matches.Count
      '
      '      For Each Match In matches
      '
      '         GUIEvent_ProcessUpdate Inc(counter)
      '
      '         With Match
      '
      '            Dim DecValue&
      '            DecValue = "&h" & .SubMatches(0)
      '
      '           ' Log "Replacing: " & DecValue & " <= " & .value
      '            ReplaceDo Data, .value, DecValue, .FirstIndex, 1
      '
      '         End With
      '      Next
      '      Log matches.Count & " 'Execute's transformed."
      '
      '      GUIEvent_ProcessEnd
      '---------

   
   End With
End Sub
