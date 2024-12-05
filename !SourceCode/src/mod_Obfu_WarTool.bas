Attribute VB_Name = "mod_Obfu_WarTool"
Option Explicit
' use to deobfuscate 'Wartool - Bots for Wartune and Legend Online'
' https://sites.google.com/site/sbotavc/



Public Sub obfu_WarTool(ClsDeobfu As ClsDeobfuscator, Data$)

   Dim FileName As New ClsFilename
   'FileName.FileName = InputBox("FileName:", "", Combo_Filename)
   FileName.FileName = FrmMain.Combo_Filename

'   Dim data$
'   data = FileLoad(FileName.FileName)
   
   Dim matches As MatchCollection
      
 ' Preclean
   Dim FunctionName$, matchcount&
   
   FunctionName = "BinaryToString"
   matchcount = BinaryToString(ClsDeobfu, Data$, FunctionName)
   
   Log matchcount & " refs to '" & FunctionName & "' transformed."

   
 '  SaveScriptData Data, True

  
 ' Split functions
   Dim Func As New tc_Func
   
   Dim functions As Collection
   Set functions = Func.splitFunc(Data)
   
  
   
 ' Find Function with 'StringReverse'
'   Dim Func_StringReverse
'   Func_StringReverse = ""
'
'   Dim item
'   For Each item In functions
'      If item Like "*StringReverse*" Then
'
'       ' Get FunctionName
'         Func_StringReverse = GetFuncName(item)
'
'         Exit For
'      End If
'   Next
   
   
   
 ' Find Functions using this StringReverse Function
  FunctionName = "StringReverse"
  matchcount = FindFuncWithDataAndDoReplace(ClsDeobfu, functions, _
      FunctionName, True, 0, 2)
      
  Log "'" & FunctionName & "' deleted.  " & matchcount & " refs to it transformed."
  

   FunctionName = "_HexToString"
   matchcount = FindFuncWithDataAndDoReplace(ClsDeobfu, functions, _
      FunctionName)
      
   Log "'" & FunctionName & "' deleted.  " & matchcount & " refs to it transformed."
   
  
   FunctionName = "BinaryToString"
  
   matchcount = FindFuncWithDataAndDoReplace(ClsDeobfu, functions, _
      "BinaryToString")
   Log "'" & FunctionName & "' deleted.  " & matchcount & " refs to it transformed."
  
   Data = Func.FuncJoin(functions)
'   SaveScriptData Data, True
  
   BinaryToString ClsDeobfu, Data$, "BinaryToString"
  
   '
 
   ' functions = Split(Data, "EndFunc", , vbTextCompare)
   
   
'      Dim myRegExp As New RegExp
'      With myRegExp
'
'         .Global = True
'         .IgnoreCase = False
'
'        ' http://regex101.com
'         .Pattern = RE_WSpace( _
'            RE_Group("\w*"), _
'            "\(", RE_Quote & "(?:0[xX])?" & _
'               RE_HEXDIGET & "*?" & RE_Quote, "\)")
'
'               Set matches = .Execute(Data)
'
'               Dim Match As Match
'
'               For Each Match In matches
'
'                     Dim FunctionName$
'                     FunctionName = Match.SubMatches(0)
'
'                     BinaryToString ClsDeobfu, Data, FunctionName
'            '   SaveScriptData data
'            '   qw
'
'
'               '      FunctionName = InputBox("FunctionName:", "", FunctionName)
'
'               Next
'
'      End With
      
'      FileName.Name = FileName.Name & "_Deobfu_unopti"
'
'    ' Save
'      FileSave FileName.FileName, data
'
'
'
'    Optimise data

'
''   If matches.Count Then
'      FileName.Name = FileName.Name & "_Deobfu"
'
'    ' Save
'      FileSave FileName.FileName, data
'
'
'       Log matches.Count & " replacements done."
'       Log "File save to: " & FileName.FileName
' '  Else
' '     Log "Nothing found."
' '  End If
      
      

End Sub

Private Function FindFuncWithDataAndDoReplace(ClsDeobfu As ClsDeobfuscator, functionsNameList As Collection, Findthis, _
         Optional isStringReverse As Boolean, _
         Optional RekLevel = 0, Optional RekMax = 1) As Long

   
   If RekLevel > RekMax Then Exit Function
   
   
   Dim bDoGUI As Boolean
   bDoGUI = (RekLevel = RekMax - 1)
   
   If bDoGUI Then GUIEvent_ProcessBegin functionsNameList.Count, 1
   
   Dim GUI_Counter&
   Dim matchcount&
   
   
   
   Dim item As tc_Func
   For Each item In functionsNameList
   
      If bDoGUI Then GUIEvent_ProcessUpdate Inc(GUI_Counter), 1
      If item.FuncData Like "*" & Findthis & "*" Then
      
       ' Get FunctionName
         Dim FunctionName$
         FunctionName = item.FuncName 'GetFuncName(item)
         
         If RekLevel > RekMax - 1 Then
          ' ... at RekMax Level do Search&Replace
            
            Dim tmpInOutParam$
            tmpInOutParam = item.FuncData
            
            Inc matchcount, _
               BinaryToString(ClsDeobfu, tmpInOutParam, Findthis, isStringReverse)
            
            item.FuncData = tmpInOutParam
         
            
         Else
         
          ' ... do Function Explore
            Inc matchcount, _
 _
               FindFuncWithDataAndDoReplace(ClsDeobfu, functionsNameList, _
                     FunctionName, isStringReverse, RekLevel + 1, RekMax)
                  
          ' Remove Function
            Dim FuncRemovedData$
            FuncRemovedData = vbCrLf & "; Funky funk " & FunctionName & " - Remove by Deobfucator"
            item.FuncData = FuncRemovedData
            
          ' QuickReplace Data, item & "EndFunc", _
          ' FuncRemovedData
                  
         End If
        
      End If
      
   Next
   

   If bDoGUI Then GUIEvent_ProcessEnd 1

   FindFuncWithDataAndDoReplace = matchcount
   
End Function
