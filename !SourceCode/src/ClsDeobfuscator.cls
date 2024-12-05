VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDeobfuscator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Dim DeObfu_FunctionNames As New Collection
Dim NewScript$

Sub Log(Text$)
   FrmMain.Log Text
End Sub

Sub log2(Text$)
   FrmMain.log2 Text
End Sub

Private Sub GUI_Show_Text(Text$)
   FrmMain.Txt_Script = Text
End Sub

Private Sub GUI_StatusBar_SetCurrentLine(ByVal Line&)

End Sub

Private Sub GUI_StatusBar_SetLines(ByVal Lines&)

End Sub



'Private Function Eval(Expr As String) As String
'   'Dim x  As MSScriptControl.ScriptControl
'   'Set x = New MSScriptControl.ScriptControl
'   Dim x As Object
'   Set x = CreateObject("MSScriptControl.ScriptControl") 'regsvr32.exe <WINXP>\SYSTEM32\msscript.ocx
'   x.Language = "VBScript"
'
'   Eval = x.Eval(Expr)
'
'   Set x = Nothing
'
'
'End Function
'
'Sub ReBase_Array(ByRef Arr, Dim_Min, Optional Dim_Max)
'
'   If IsMissing(Dim_Max) Then Dim_Max = UBound(Arr)
'
'   Dim NewArr As New Collection
'
'   Dim i&
'   For i = Dim_Min To Dim_Max
'      NewArr.Add Arr(i)
'   Next
'
'   Arr = CollectionToArray(NewArr)
'End Sub

Private Sub DeObfuscate_Chr_Encode()
   
'   If NewScript Like "*""[!""]*" Then
'      Log "Skipping Replacing 'Consts Globals & Dim' because there are strings in the script. Use the function rename to include these includes first."
'   Else
'     '=== Replace Dim Const and Globals
'      DeObfu_ReplaceGlobals NewScript
'      DeObfu_ReplaceDimsAndConst NewScript
'   End If


' --- Prepare regular expression objects & search pattern ---
   
   Dim Number$
   Number = "(\d{0,4})"
   
   Dim Operator$
   Operator = "([+-])"
   
 
   Dim CHR_Command$
 ' for CHR_Commands for example "CHR(142+8082)"
   CHR_Command = WSpace("CHR", "\(", Number, Operator, Number, "\)")
   
 
   Dim CHR_CommandElements$
 ' One or more CHR_Commands for example "CHR(157-125) & CHR(10+101)"
 ' used to separate & replace ChrStrings like
 ' 'CHR(157-125)' or
 ' 'CHR(157-125) & ""' or
 ' 'CHR(157-125) & CHR(10+101)' or
 ' 'CHR(157-125) & CHR(10+101) & ""

   CHR_CommandElements = WSpace(CHR_Command & _
                      "(", "&", CHR_Command, ")*" & _
                      "(", "&", """""){0,1}")
  
   Dim ChrStrings As New RegExp
   With ChrStrings
      .IgnoreCase = True
      .Global = False
      .Pattern = CHR_CommandElements
      
   End With
 
   Dim Chrs As New RegExp
 ' RegExp for separating Chrs like "CHR( '10' '+' '101' )"
   With Chrs
      .IgnoreCase = True
      .Global = True
      .Pattern = WSpace(CHR_Command)
   End With
 
' 'Break String into lines (but retain "_ "&CRLF as one line)
'   Dim Lines As New RegExp
'   With Lines
'       .IgnoreCase = True
'       .Global = True
'       .MultiLine = True
'       .Pattern = WSpace(".*", "[^_]", "\n")
'   End With
'   Dim Lines_Match As MatchCollection
'   Set Lines_Match = Lines.Execute(NewScript)
'
'


' --- Process Script ---

   Dim ChrStringNew As clsStrCat
   Set ChrStringNew = New clsStrCat
 
 ' Merge lines with _ at the end
   NewScript = Replace(NewScript, "_" & vbCrLf, "")

 ' Break into lines and process them
   Dim ScriptLines
   ScriptLines = Split(NewScript, vbCrLf)
   
   Dim Lines As Long
   Lines = UBound(ScriptLines)
   GUI_StatusBar_SetLines Lines
 
  'Go through all Lines
   Dim Line_idx
   For Line_idx = 0 To Lines
      Dim Line$
      Line = ScriptLines(Line_idx)
      GUI_StatusBar_SetLines Line_idx
'      GUI_Show_Text Line_idx & ": " & ScriptLines(Line_idx)
      
    
'    If Line_idx = 13245 Then Stop
    
    
    ' separate ChrString's
      Do While True
         Dim ChrStrings_Matches As MatchCollection
         Set ChrStrings_Matches = ChrStrings.Execute(Line)
    ' Exit (and process nxt line) if there are no ChrStrings
      If ChrStrings_Matches.Count = 0 Then Exit Do

        ChrStringNew.Clear
         
       'separate Chr's
        Dim Chrs_Match As Match
        For Each Chrs_Match In Chrs.Execute(ChrStrings_Matches(0))
           
          ' Concat Chrs to a ChrString
            With Chrs_Match
              
             ' "CHR(157-125)" -> ...
               Dim Chrs_op$, Chrs_arg1&, Chrs_arg2&
               Chrs_arg1 = .SubMatches(0)    'Chrs_arg1 = 157
               Chrs_op = .SubMatches(1)      'Chrs_op   = "-"
               Chrs_arg2 = .SubMatches(2)    'Chrs_arg2 = 125
            
            End With
              
          ' calc new Char
            Select Case Chrs_op
            Case "+"
               ChrStringNew.Concat (ChrW(Chrs_arg1 + Chrs_arg2))
            Case "-"
               ChrStringNew.Concat (ChrW(Chrs_arg1 - Chrs_arg2))
            End Select
           
         Next Chrs_Match
       
       ' replace ChrString in line
         ReplaceDo Line, CStr(ChrStrings_Matches(0)), MakeAutoItString(ChrStringNew.value)

       ' save modified Line
         ScriptLines(Line_idx) = Line
         
      Loop ' next ChrString in Line

      GUI_Show_Text Line_idx & ": " & ScriptLines(Line_idx)
      DoEventsSeldom
      
   Next Line_idx


  'Join lines back to a script
   NewScript = Join(ScriptLines, vbCrLf)

End Sub



'Private Sub OLD_DeObfuscate_Chr_Encode()
'
'
'
'
' ' Merge lines with _ at the end
' '  NewScript = Replace(NewScript, vbTab, "")
'   NewScript = Replace(NewScript, "&""""&", "&") '
'   NewScript = Replace(NewScript, "_" & vbCrLf, "")
'
' ' Break into lines and process them
'   Dim ScriptLines
'   ScriptLines = Split(NewScript, vbCrLf)
'
'   Dim Line_idx
'   For Line_idx = 0 To UBound(ScriptLines)
'    Dim Line$
'    Line = ScriptLines(Line_idx)
'
'' 'dbg
''  If Line_idx = 33900 Then Stop
'
'     'Seperated ChrStrings of each line
'      Dim ChrStrings
'      ChrStrings = Split(Line, "&""""")
'
'
'     'If contains ChrStrings...
'      If UBound(ChrStrings) > 0 Then
'
'         Dim ChrString_idx&
'         For ChrString_idx = 0 To UBound(ChrStrings)
'          Dim ChrString$
'          ChrString = ChrStrings(ChrString_idx)
'
'            Dim EncodedChars
'            EncodedChars = Split(ChrString, "CHR(")
'
'           'If contains chr's...
'            If UBound(EncodedChars) > 0 Then
'              '... simplify them to a String
'               Dim CHR_PreData$
'               CHR_PreData = EncodedChars(0)
'               EncodedChars(0) = ""
'
'               Dim Clean_ChrString$
'               Clean_ChrString = Join(EncodedChars, "CHR(")
'
'               On Error Resume Next
'               Dim Clean_String$
'               Clean_String = Eval(Clean_ChrString)
'               If Err = 0 Then
'                  ChrStrings(ChrString_idx) = CHR_PreData & Quote(Clean_String)
'               Else
'                  Log "Problem on Line: " & Line_idx & " (" & Err.Description & ") - Line skipped."
'               End If
'            End If
'
'         Next
'        'Rejoin Line
'         ScriptLines(Line_idx) = Join(ChrStrings, " ")
'      End If
'   Next
'
'   NewScript = Join(ScriptLines, vbCrLf)
'
'
'End Sub


Private Sub DeObfuscate_Encodeit()

 ' Merge lines with _ at the end
   NewScript = Replace(NewScript, "_" & vbCrLf, "")


 ' ======== Var Replace
   DeObfu_ReplaceDimsAndConst NewScript
   DeObfu_ReplaceGlobals NewScript
   DeObfu_ReplaceLocals NewScript



'==== Get Vectors
   Dim globs
   globs = Split(NewScript, "Global Const ")
   
   Dim FnDeCryHexToStr$
   FnDeCryHexToStr = Trim(strCrop1(globs(2), "=", "("))
   
   Dim b
   b = Split(globs(2), "=")
   
   Dim FnDeCryKeyVar$
   FnDeCryKeyVar = Trim(b(0))
   
   Dim FnDeCryHexToStrIn$
   FnDeCryHexToStrIn = Trim(strCrop1(globs(2), "'", "'"))
   
   Dim FnDeCryHexToStrKey, q
   q = Split(NewScript, FnDeCryHexToStr)(2)
   q = Split(q, "=")(1)
   'FnDeCryHexToStrKey = strCrop1(NewScript, FnDeCryHexToStr, "StringLen", , InStrRev(NewScript, FnDeCryHexToStr))
   FnDeCryHexToStrKey = Split(q, vbCrLf)(0)
   
   Dim AA&, BB&, Key&
   
   AA = ("&h" & Mid(FnDeCryHexToStrIn, 1, 2)) - FnDeCryHexToStrKey
   BB = ("&h" & Mid(FnDeCryHexToStrIn, 3, 2)) - FnDeCryHexToStrKey
   Key = Chr(AA) & Chr(BB)
   
  
   
  '==== Load String =====
  
   Dim i&, Line$, EncString$
   Dim VarNames As New Collection
   Dim VarValues As New Collection

   For i = 3 To UBound(globs)
      Dim tmpstr$
      tmpstr = globs(i)
      Line = strCropAndDelete(tmpstr, "", vbCrLf)
      globs(i) = tmpstr
   

      
      Dim Words
      Words = Split(Line, "'")
      EncString = Words(1)
      
      Dim VarName$
      VarName = Trim(strCrop1(Words(0), "$", "="))
         
      Dim DecString$
      DecString = Space(Len(EncString) \ 2)
      
      Dim pos&
      For pos = 1 To Len(EncString) Step 2
         Mid(DecString, (pos \ 2) + 1) = Chr(("&h" & Mid(EncString, pos, 2)) - Key)
      Next
      
      VarNames.Add VarName
      VarValues.Add DecString
      
   If tmpstr <> "" Then Exit For
'      log VarName & "   " & DecString
      
   Next
   
'  === Join Script & Remove EncodeIT "Global Const "
   Dim j&
   For j = i To UBound(globs)
      globs(j - i) = globs(j)
   Next
   
   ReDim Preserve globs(j - i - 1)
   NewScript = Join(globs, "Global Const ")
   GUI_Show_Text NewScript
   
   Log "Search'n'replace strings: "
   
'  ==== Replace Strings =====
   
   Dim ItemIndex&
   For ItemIndex = 1 To VarNames.Count
'       log VarNames(ItemIndex) & "   " & VarValues(ItemIndex)
      
      NewScript = Replace(NewScript, CStr("$" & VarNames(ItemIndex)), CStr("""" & Replace(VarValues(ItemIndex), """", """""") & """"))
      
      Log CStr("$" & VarNames(ItemIndex)) & "    " & """" & Replace(VarValues(ItemIndex), """", """""") & """"

   Next
   
   GUI_Show_Text NewScript
   
   Log "Processing function names and arguments:"
   
' ======== Functions

'  'Securitycheck last  functionname=FnDeCryHexToStr
'   Dim funcName$
'   funcName = Trim(Split(Funcs(UBound(Funcs)), "(")(0))
'   Debug.Assert (FnDeCryHexToStr <> funcName)

  Dim Funcs, FuncIdx
  Funcs = Split(NewScript, vbCrLf & "Func ")
  
' Remove last 3 functions added by EncodeIt
  ReDim Preserve Funcs(UBound(Funcs) - 3)
  
  
  For FuncIdx = LBound(Funcs) + 1 To UBound(Funcs)

   ' getName
   Dim funcName$
   funcName = Trim(Split(Funcs(FuncIdx), "(")(0))
   
'   Debug.Assert (FnDeCryHexToStr <> funcName) ' and (FuncIdx=UBound(Funcs)-2)

   ' add name for replace later
   DeObfu_FunctionNames.Add funcName

   ' seperate Body & rest
   Dim funcBody$, funcRest$, tmp
   tmp = Split(Funcs(FuncIdx), vbCrLf & "EndFunc")
   funcBody = tmp(0)
   funcRest = tmp(1)

   'replace
   DeObfu_ReplaceFnArgs funcBody, funcName
   ' writeback

   tmp(0) = funcBody
   Funcs(FuncIdx) = Join(tmp, vbCrLf & "EndFunc")


  Next

  NewScript = Join(Funcs, vbCrLf & "Func ")
   
  
  
  DeObfu_ReplaceFunctionNames NewScript
   
End Sub


Public Function DetectScriptObfuscated() As Long
'   FileName = "I:\!Cracks & Projects\AutoIt3\Obfuscator\smarty_Obfuscated.au3"
Log "Trying to DeObfuscate : " & FileName
   
 
 '=== Open & Test Obfu-Type ===
   
   File.Create FileName.FileName, False, False, True
   
  'Convert to Accii if it's unicode

   Dim UnicodeMarker%
   UnicodeMarker = File.intValue
   If UnicodeMarker = &HFEFF Then
      NewScript = StrConv(File.FixedString(-1), vbFromUnicode)
   Else
      File.Move -2
      NewScript = File.FixedString(-1)
   End If
 
 ' Test for if obfuscated ( find 'func A2C4558BC554')
' Option Compare Text
'   Dim UCaseScript$
'   UCaseScript = UCase(NewScript)
   If NewScript Like "*Func A2C4558BC554*" Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE
   ElseIf NewScript Like "*""4074656D70646972""*" Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE

   ElseIf NewScript Like "*4054656D70446972*" Then
      DetectScriptObfuscated = DE_OBFUSC_VANZANDE_VER24


   'Obfuscator v1.0.24.23  ( find 'func A1D00103002')
   ElseIf NewScript Like "*A1D00103002*" Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE
      
   ElseIf NewScript Like "*Int(99/3+15[*]100/4-13^2+81/3-17-245+99/3+15[*]100/4-13^2+81/3-17)*" Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_ENCODEIT
      
   ElseIf NewScript Like "*Chr(##?[-+]?##*" Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_CHR_ENCODE
      
   Else
      DetectScriptObfuscated = DE_OBFUSC_TYPE_NOT_OBFUSC
   End If
   
   
End Function


Public Sub DeObfuscate()

   ' Test for if obfuscated
   'my dilettantic deObfuscator
   Select Case DetectScriptObfuscated
      Case DE_OBFUSC_TYPE_NOT_OBFUSC
         File.CloseFile
         Err.Raise ERR_NO_OBFUSCATE_AUT, , "Script is NOT obfuscate by 'Jos van der Zande AutoIt3 Source Obfuscator v1.0.15 [July 1, 2007]', v1.0.14[June 16, 2007] or 'EncodeIt 2.0'"
         Exit Sub

      Case DE_OBFUSC_VANZANDE_VER24
         Log "Deobfuscating van Zande 1.0.24.22..."
         DeObfuscate_VanZande1_0_24
      
      Case DE_OBFUSC_TYPE_VANZANDE
         If InStr(1, NewScript, "lobal $Os") Then
            Log "Deobfuscating van Zande 1.0.15..."
            DeObfuscate_VanZande1_0_15
         ElseIf InStr(1, NewScript, ".tbl") Then
            Log "Deobfuscating van Zande 1.0.24..."
            DeObfuscate_VanZande1_0_15
         
         
         Else
            Log "Deobfuscating van Zande 1.0.14..."
            DeObfuscate_VanZande1_0_14
         End If
         
      Case DE_OBFUSC_TYPE_ENCODEIT
         Log "Deobfuscating Encodeit 2.0..."
         DeObfuscate_Encodeit
         
      Case DE_OBFUSC_TYPE_CHR_ENCODE
         Log "Deobfuscating Chr_Encode..."
         DeObfuscate_Chr_Encode
   End Select
   
   SaveScriptData NewScript
   Log "Deobfucation succeed."
End Sub

Function StringToHexString$(Txt_String)
   Dim tmpBuff As New StringReader
   tmpBuff = Txt_String
   StringToHexString = ValuesToHexString(tmpBuff, "")
End Function

Private Sub DeObfuscate_VanZande1_0_24()
' In v1.0.24.23 Seperator is wrapped up into "Execute(Binarytostring('0x"- 3 times
' Like this:
'FileInstall('test_org.au3.tbl',$AF123123,1)
'Global $Os = Execute(Binarytostring('0x457865637574652842696E61....
'                                       Execute(Binarytostring('0x457865637574652842696E6...
'                                                                 Execute(Binarytostring('0x537472696E6753706C6974...
'                                                                                           StringSplit(FileRead($AF123123...
   
   Dim tmp2$
   tmp2 = NewScript
   Do
      Dim tmp$
      tmp = strCrop1(tmp2, _
                  StringToHexString("Execute(Binarytostring('0x"), _
                  StringToHexString("'))"))
      
   If tmp = "" Then Exit Do
      tmp2 = HexStringToString(tmp)
   Loop While True
   
     
   'VanZande1_0_24
   Dim TBLStringSeperator$
   
   '0x537472696E6753706C69742846696C655265616428244146313233313233292C27 | 69354F | 272C3129
   '                                   StringSplit(FileRead($AF123123),' | i5O    | ',1)
   TBLStringSeperator = HexStringToString(strCrop1(tmp2, StringToHexString("StringSplit(FileRead($AF123123),'"), StringToHexString("',1)"), "4F"))
   Log "TBL seperator string: " & TBLStringSeperator
   DeObfuscate_VanZande1_0_15 TBLStringSeperator

End Sub


Private Sub DeObfuscate_VanZande1_0_15(Optional TBLStringSeperator = "O")


 ' Merge lines with _ at the end + Add vbCrLf after last EndFunc
   NewScript = Replace(NewScript, " _" & vbCrLf, "") ' & vbCrLf
   
''Correct case (may it's safer to let tidy do)
'   NewScript = Replace(NewScript, "Func ", "Func ", , , vbTextCompare)
'   NewScript = Replace(NewScript, "EndFunc ", "EndFunc ", , , vbTextCompare)
'   NewScript = Replace(NewScript, "IsDeclared", "IsDeclared", , , vbTextCompare)
'   NewScript = Replace(NewScript, "EndIf", "EndIf", , , vbTextCompare)

   
   Dim OSString$
   OSString = strCropWithSeek(NewScript, "($", "[")
   NewScript = Replace(NewScript, OSString, "Os", , , vbTextCompare)

   NewScript = Replace(NewScript, "$Os [", "$Os[", , , vbTextCompare)


'=== Replace Dim Const and Globals
BenchStart
   DeObfu_ReplaceGlobals NewScript
BenchEnd
' ReplaceDoMulti 49297 ;51078
' ReplaceDo 23188
' Empty 9719  9625  9641

   DeObfu_ReplaceDimsAndConst NewScript

   
   GUI_Show_Text NewScript

   Log "Okay. Obfucated script loaded and displayed."
   
'=== Prepare ===


    strCropAndDelete NewScript, "Global $", "StringTrimLeft(String(StringToBinary(""\"" & Hex(Random(256, 268435455, 1), 7))), 2)"

    Dim FnNameLoadTBL$
    FnNameLoadTBL$ = strCropAndDelete(NewScript, "global $Os" & vbCrLf, "()")
  
  
  
    Dim FnNameHexToString$
    FnNameHexToString = strCropWithSeek(NewScript, "=", "($", vbCrLf & "global ")
   
  
  
 Log "Extracting TBLFileName from Script."
  ' GetStringfile(*.tbl) from obfuscate script
    Dim StringTBLFileName As New ClsFilename
    StringTBLFileName.Path = FileName.Path
    StringTBLFileName.NameWithExt = Mid(strCrop(NewScript, "(", ".tbl"), 2) & ".tbl"
'    StringTBLFileName.NameWithExt = Mid(strCrop(NewScript, "FileInstall(", ".tbl"), 2) & ".tbl"
'    Dim line
'    For Each line In Split(NewScript, vbCrLf)
'    FileInstall('*.tbl",
'      StringTBLFileName.NameWithExt = strCrop(NewScript, "FileInstall('", "'")
'      Exit For
'    Next
    
 
 Log "Loading StringTBLFileName: " & StringTBLFileName.FileName & "..."
    
    Dim StringTBLFile As New FileStream
    StringTBLFile.Create StringTBLFileName.FileName
    
   'Read Stringtable
    Dim StringTBL
    ReDim StringTBL(1 To 1)
    
      Dim Line
      StringTBLFile.Position = 0
      For Each Line In Split(StringTBLFile.FixedString(-1), TBLStringSeperator)
         StringTBL(UBound(StringTBL)) = HexStringToString(Line)
         ReDim Preserve StringTBL(1 To UBound(StringTBL) + 1)
      Next
      StringTBLFile.CloseFile
      ReDim Preserve StringTBL(1 To UBound(StringTBL) - 1)


   
   
Log UBound(StringTBL) & " strings found."
 
  
  
  
Log "Restoring StringNames in ..."
  Dim LocalStringNames
  ReDim LocalStringNames(UBound(StringTBL))
  
  ' get local stringvars like
  'global $SA220010601E=A490000263F($Os[0x1]),$SA3A00204A08=A490000263F($Os[0x2]),$SA200030463A=A490000263F($Os[0x3])
   
   Dim MainIndexes As New Collection

   Dim Stat_NumLocal&, Stat_NumAtAll&
   Stat_NumAtAll = 0


BenchStart
  
   Dim EndToken$, NewScript2
   EndToken = "EndFunc" & vbCrLf
  
   NewScript2 = Split(NewScript, "Func ")
   Dim i&
   Dim iter: i = 0
   For iter = LBound(NewScript2) To UBound(NewScript2)
   Dim ScriptFunction$
   ScriptFunction = NewScript2(iter)
      
 ' Show ScriptFunction Before
   GUI_Show_Text ScriptFunction
      
   
    ' Get Current FunctionName
      Dim FunctionName$
      If i = 0 Then
         FunctionName = "Main"
         Log "FunctionName: " & FunctionName
         
      Else
        
       ' Extract FunctionName, Store it for later
         FunctionName = strCrop1(ScriptFunction, "", "(")
         DeObfu_FunctionNames.Add FunctionName
         Log "FunctionName: " & FunctionName
       
       ' === Replace Locals ====
         DeObfu_ReplaceLocals ScriptFunction
        
       ' === Replace FunctionArguments  ====
         DeObfu_ReplaceFnArgs ScriptFunction, FunctionName
         
         
        'Check for Mainfunction chunks between FUNC and ENDFUNC
         Dim afterEndFunc$
         afterEndFunc = Split(CStr(ScriptFunction), EndToken)(1)
         If Replace(CStr(afterEndFunc), vbCrLf, "") <> "" Then
           '... apply String Restore
            Dim index
            For Each index In MainIndexes
              DeObfu_ReplaceStrings ScriptFunction, CStr(LocalStringNames(index)), CStr(StringTBL(index))
            Next
         End If
'         GUI_Show_Text ScriptFunction

         
         
      End If
      
 ' Exit if FnNameLoadTBL is reached
   If FunctionName = FnNameLoadTBL Then Exit For
      
      


'   Debug.Assert "A2600102853" <> FunctionName

     
    'Delete/ Seperate 'Obfu-Header' in ScriptFunction
    'store deleted part in obfuSection for further processing
    'Example for an 'Obfu-Header'
    '   If Not Isdeclared("SSA110020160C") Then
    '      Global $SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])
    '      Global $SSA110020160C=0x01
    '   End If
     Dim obfuSection$
     
      If i = 0 Then
         obfuSection = ScriptFunction
      Else

'         obfuSection = strCropAndDelete(tmpstr, "If Not Isdeclared(", "EndIF" & vbCrLf)
         obfuSection = strCropAndDelete(ScriptFunction, "If Not IsDeclared(", "EndIf" & vbCrLf)
'         'maybe also triggered by some normal 'IsDeclared'
'         Debug.Assert InStr(1, ScriptFunction, "IsDeclared", vbTextCompare) = 0
      End If
         
    ' Statistic Counter for String in Current function
      Stat_NumLocal = 0
      
      Do
         
         
         'Process Local String Section
          Dim GlobalStringDef, CommandSeperateLines$
          CommandSeperateLines = strCropAndDelete(obfuSection, IIf(i = 0, "global ", "Global "), vbCrLf)
       '  Now CommandSeperateLines contains this:
       ' "$SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])"
      
      
    ' EXIT DO: if CommandSeperateLines dont contains sth like $Os[0x7] it's no obfu definition
      If InStr(1, CommandSeperateLines, "$Os[", vbTextCompare) = 0 Then Exit Do
          
         If i = 0 Then
            strCropAndDelete ScriptFunction, IIf(i = 0, "global ", "Global "), vbCrLf
         End If
          
          GlobalStringDef = Split(CommandSeperateLines, ",")
         'GlobalStringDef[0] = $SA3160B01D3E=A0000006239($Os[0x7])
         'GlobalStringDef[0] = $SA1D60C0011E=A0000006239($Os[0x8])
         '....
         
        'for all items in GlobalStringDef do...
         Dim LocalStringName
         For Each LocalStringName In GlobalStringDef
             
             'Sepertate GlobalStringDef further...
             Dim LocalStringName_Splitted
             LocalStringName_Splitted = Split(LocalStringName, "=")
             'LocalStringName_Splitted    with         $SA3160B01D3E=A0000006239($Os[0x7])
             'LocalStringName_Splitted[0] = $SA3160B01D3E
             'LocalStringName_Splitted[1] = A0000006239($Os[0x7])
             
             If LocalStringName_Splitted(1) Like "*$Os[[]0x*" Then
               index = CLng("&h" & strCrop1(LocalStringName_Splitted(1), "$Os[0x", "])"))
             Else
               index = CLng(strCrop1(LocalStringName_Splitted(1), "$Os[", "])"))
             End If
             
             LocalStringNames(index) = Trim(LocalStringName_Splitted(0))
            'LocalStringNames(&h7) = "$SA3160B01D3E"
            
             If i = 0 Then
               ' Store for later search&replace in mainfunction code between functions
                 MainIndexes.Add index
'                 Debug.Print LocalStringNames(index), StringTBL(index), index
             End If

          
              
           'Replace Strings
          ' In ScriptFunction code Replace all "$SA3160B01D3E" with for exsample "Wrong Name Entered!"
'            Debug.Print "Number(" & LocalStringNames(index) & ")", StringTBL(index)
            DeObfu_ReplaceStrings ScriptFunction, CStr(LocalStringNames(index)), CStr(StringTBL(index))

              
         Next

      Stat_NumLocal = Stat_NumLocal + UBound(GlobalStringDef)
      Loop While True
      
      Log "  Local strings: " & Stat_NumLocal
      Stat_NumAtAll = Stat_NumAtAll + Stat_NumLocal
    
   ' Show ScriptFunction after
     GUI_Show_Text ScriptFunction

    
    
    
    ' Save New Function
      NewScript2(i) = ScriptFunction
   
     
   i = i + 1
   Next
  
Log "Strings at all: " & Stat_NumAtAll
Log "strings in Tbl: " & UBound(StringTBL)
  
' Cut of rest after FnNameLoadTBL
  ReDim Preserve NewScript2(i - 1)
  
  
   NewScript = Join(NewScript2, vbCrLf & "Func ")
'  Stop
  
  
  '=== Replace Functionname
   DeObfu_ReplaceFunctionNames NewScript


'  '=== Replace Global Strings
'
'  log "Restoring global strings..."
'  For Each index In MainIndexes
'      DeObfu_ReplaceStrings NewScript, CStr(LocalStringNames(index)), StringTBL(index)
'  Next
'BenchEnd

  
  'Used String
'  Dim i&
   Log "UnAssigned/Unused strings [Index, Value]: "
   For i = LBound(StringTBL) To UBound(StringTBL)
      If LocalStringNames(i) = "" Then
         Log "  " & H16(i) & " " & StringTBL(i)
      End If
   Next
  
  File.CloseFile
  
  
 ' Clean up - delete *.tbl file
   If FrmMain.Chk_TmpFile.value = vbUnchecked Then
'      log "Deleting: " & StringTBLFileName.FileName
      FileDelete StringTBLFileName.FileName
   End If
  
  
  
End Sub



'Function ReplaceAndDetectChange(ByRef Text$, ParamArray Search_Replace()) As Boolean
'   Dim i&
'   For i = 0 To UBound(Search_Replace) Step 2
'
'      Dim LenText
'      LenText = Len(Text)
'
''      Text = Replace(Text, CStr(Search_Replace(i)), CStr(Search_Replace(i + 1)))
'      ReplaceDo Text, CStr(Search_Replace(i)), CStr(Search_Replace(i + 1))
'
''      ReplaceBin Text, Text, Text, CStr(Search_Replace(i)), CStr(Search_Replace(i + 1)), 1, 2147483647
'
'      ReplaceAndDetectChange = (LenText <> Len(Text))
'      If ReplaceAndDetectChange = True Then Exit For
'
'   Next
'
'
'End Function


Private Sub DeObfu_ReplaceFunctionNames(ByRef ScriptCode$)
   Dim FnNamePure, FnNameCounter&: FnNameCounter = 0
   For Each FnNamePure In DeObfu_FunctionNames
   
      Dim FnNameNew$
      FnNameNew = "Fn" & H16(FnNameCounter)
      
      ReplaceDo ScriptCode, CStr(FnNamePure), FnNameNew
      
      Inc FnNameCounter
   Next
End Sub
Private Sub DeObfu_ReplaceFnArgs(ByRef CodeSnipped$, ByRef FunctionName$)

  Dim ParamList$
  ParamList = strCrop1(CodeSnipped, FunctionName, vbCrLf)
  ParamList = strCrop1(CodeSnipped, "(", ")")
  
  Dim RawParams
  RawParams = Split(ParamList, ",")


   Dim FnArgument, FnNewName$, ArgCounter&
   ArgCounter = 0
   For Each FnArgument In RawParams
     
     FnNewName = "Arg"
     If FnArgument Like "*Const*" Then FnNewName = FnNewName & "C"
     If FnArgument Like "*ByRef*" Then FnNewName = FnNewName & "Ref"
     If FnArgument Like "*=*" Then
        FnNewName = FnNewName & "Opt"
        FnArgument = Split(FnArgument, "=")(0)
     End If
     FnNewName = FnNewName & H8(ArgCounter)
     
     
     Dim FnPureArgName$
     FnPureArgName = Trim(strCrop1(FnArgument, "$", ""))
'            tmpstr = CodeSnipped
     
     ReplaceDo CodeSnipped, CStr(FnPureArgName), FnNewName
'            CodeSnipped = tmpstr
     
     Inc ArgCounter
   Next

End Sub

Private Sub z_DeObfu_ReplaceX(ByRef CodeSnipped$, KeyWord$, Prefix$, Optional IgnoreLinesWith = "$Os")
   If IsMissing(IgnoreLinesWith) Then IgnoreLinesWith = "$Os"
   
   Dim LocalDefCommaSep, LocalDefCounter: LocalDefCounter = 0
   Dim LocalNameCounter&: LocalNameCounter = 0
   For Each LocalDefCommaSep In Split(CodeSnipped, KeyWord)
     'Skip first element
      If LocalDefCounter > 0 Then
         Dim LocalDefLine$
         LocalDefLine = Split(LocalDefCommaSep, vbCrLf)(0)
         If (LocalDefLine Like "*" & IgnoreLinesWith & "*") = False Then
         
           
           'crop off functions args (Arg1,arg2,arg3...) that also contains ','
            LocalDefLine = CropParenthesis(LocalDefLine)
            LocalDefLine = CropParenthesis(LocalDefLine, "[", "]")
            
            Dim LocalDef
            For Each LocalDef In Split(LocalDefLine, ",")

             ' Cut out Current Name
               Dim LocalDefPureName$
               LocalDefPureName = Split(LocalDef, "$")(1)
               LocalDefPureName = Split(LocalDefPureName, "=")(0)
               LocalDefPureName = Split(LocalDefPureName, ")")(0)
               ReplaceDo LocalDefPureName, vbTab, ""
               LocalDefPureName = Trim(Split(LocalDefPureName, "[")(0))

               
               
             ' Create New Name
               Dim LocalNameNew$: LocalNameNew = Prefix
               If LocalDef Like "*const*" Then LocalNameNew = LocalNameNew & "Const"
               If LocalDef Like "*[[]*" Then
                  LocalNameNew = LocalNameNew & "Arr"
               Else
                  LocalNameNew = LocalNameNew & "Var"
               End If
               LocalNameNew = LocalNameNew & H16(LocalNameCounter)
               
             ' Replace current with new name
               ReplaceDo CodeSnipped, LocalDefPureName, LocalNameNew
               'ReplaceDoMulti CodeSnipped, LocalDefPureName, LocalNameNew, vbTextCompare
               DoEventsSeldom
               
               Inc LocalNameCounter
            Next
         Else
         End If
      End If
      
      Inc LocalDefCounter
   Next


End Sub




Private Sub DeObfu_ReplaceGlobals(ByRef CodeSnipped$, Optional IgnoreLinesWith)
Log "Renaming Globals..."
   'z_DeObfu_ReplaceX CodeSnipped, "Global", "g", IgnoreLinesWith
'Global Const $A4A14104939 = $A2014201D1E

'Global $A4201D12945 =
'http://msdn.microsoft.com/en-us/library/1400241x.aspx
  Dim ObfuName$
  ObfuName = "[0123456789ABCDEF]{11}"
  Dim GlobalVar As New RegExp
  With GlobalVar
      .IgnoreCase = True
      .Global = False
      '.Pattern = WSpace("Global $" & ObfuName, "=")
      .Pattern = ObfuName
      
  End With
 

 'separate Chr's
  Dim GlobalVar_Match As Match
  For Each GlobalVar_Match In GlobalVar.Execute(NewScript)
     
    ' Concat Chrs to a ChrString
      With GlobalVar_Match
        Dim tmp
         tmp = .SubMatches(0)    'Chrs_arg1 = 157
      End With
        
   Next GlobalVar_Match
 
 ' replace ChrString in line
'   ReplaceDo Line, CStr(ChrStrings_Matches(0)), MakeAutoItString(ChrStringNew.value)

 ' save modified Line
 '  ScriptLines(Line_idx) = Line
   
'Loop ' next ChrString in Line




End Sub
Private Sub DeObfu_ReplaceDimsAndConst(ByRef CodeSnipped$, Optional IgnoreLinesWith)
Log "Renaming Dims..."
   z_DeObfu_ReplaceX CodeSnipped, "Dim", "gDim", IgnoreLinesWith
Log "Renaming Consts..."
   z_DeObfu_ReplaceX CodeSnipped, vbCrLf & "Const", "gConst", IgnoreLinesWith
   
   z_DeObfu_ReplaceX CodeSnipped, vbCrLf & "Global Const ", "glConst", IgnoreLinesWith
   
End Sub

Private Sub DeObfu_ReplaceLocals(ByRef CodeSnipped$)
'log "Renaming Locals..."
   z_DeObfu_ReplaceX CodeSnipped, "Local ", ""
End Sub

Private Sub DeObfu_ReplaceStrings(ByRef Data$, ByRef ReplaceStr$, ByRef NewStr$)
'log "Renaming Locals..."
 ' avoid Replacing ShellExecute
   ReplaceDo Data, "ShellExecute", "Shell_xecute"
      ReplaceDo Data, "Execute(" & ReplaceStr & ")", NewStr
   ReplaceDo Data, "Shell_xecute", "ShellExecute"
      ReplaceDo Data, "Number(" & ReplaceStr & ")", FormatNumber(NewStr), , , vbTextCompare
      ReplaceDo Data, ReplaceStr, """" & Replace(NewStr, """", """""") & """"
End Sub




Private Sub DeObfuscate_VanZande1_0_14()

   'Merge lines with _ at the end
    NewScript = Replace(NewScript, "_" & vbCrLf, "")



   GUI_Show_Text NewScript

   Log "Okay. Obfucated script loaded and displayed."
  
  
   Dim StringTBL As New Collection
  

  Dim LocalStringNames
  ReDim LocalStringNames(StringTBL.Count)
  
  ' get local stringvars like
  'global $SA220010601E=A490000263F($Os[0x1]),$SA3A00204A08=A490000263F($Os[0x2]),$SA200030463A=A490000263F($Os[0x3])
   
   Dim fnNameDecrypt$
   
   Dim MainIndexes As New Collection

   Dim Stat_NumLocal&, Stat_NumAtAll&
   Stat_NumAtAll = 0

  '=== Replace Dim Const and Globals
  DeObfu_ReplaceDimsAndConst NewScript
   
   
   
  ' global $SA60D010561B=A0300004831("41415354")
  '                      ^^^^^^^^^^^-> fnStringDecrypt
      Dim NewScript2
      NewScript2 = Split(NewScript, "Func ", , vbTextCompare)
      Dim fnStringDecrypt$
    ' fnStringDecrypt is the Last FunctionName
      fnStringDecrypt = strCrop1(NewScript2(UBound(NewScript2) - 1), "", "(")
   
  'Ignore all lines with "A0300004831()" ( <-fnStringDecrypt ) to speed up replace
  'lines with "A0300004831()" will be deleted later anyway
'!!!   DeObfu_ReplaceGlobals NewScript, fnStringDecrypt


Log "Restoring StringNames in ..."
  
  
   Dim i&
   NewScript2 = Split(NewScript, "Func ", , vbTextCompare)
  
   
   Dim iter: i = 0
   For iter = LBound(NewScript2) To UBound(NewScript2)
   Dim ScriptFunction$
   ScriptFunction = NewScript2(iter)
 
    ' Get Current FunctionName
      Dim FunctionName$
      
    ' is first function chunk
      If i = 0 Then
         FunctionName = "Main"
      
      Else
        
         FunctionName = strCrop1(ScriptFunction, "", "(")
'         Debug.Assert FunctionName <> "Stoper"
       
       'if FunctionName was Obfuscated Store it for later
       If FunctionName Like "A??????????" Then DeObfu_FunctionNames.Add FunctionName
   
      
        ' === Replace Locals ====
          DeObfu_ReplaceLocals ScriptFunction
         
        ' === Replace FunctionArguments  ====
          DeObfu_ReplaceFnArgs ScriptFunction, FunctionName
         
      End If '//is first function chunk
      
 ' Exit if FnNameLoadTBL is reached
'   If FunctionName = FnNameLoadTBL Then Exit For
      
      

   Log "FunctionName: " & FunctionName
   'Debug.Assert "A5E30E02006" <> FunctionName

     
    'Delete/ Seperate 'Obfu-Header' in ScriptFunction
    'store deleted part in obfuSection for further processing
    'Example for an 'Obfu-Header'
    '   If Not Isdeclared("SSA110020160C") Then
    '      Global $SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])
    '      Global $SSA110020160C=0x01
    '   End If
     Dim obfuSection$
     
      If i > 0 Then
      
       ' Find out if it look like this
       ' If Not Isdeclared("SSA1810601658") Then
       ' Global $SA0...
       ' EndIF
       '
       ' OR
       '
       ' If Not Isdeclared("SSA1810601658") Then Global $SA0...
       ' <without EndIF>
         Dim isIf_EndIF_Type As Boolean
         On Error Resume Next
         Dim If_Not_Isdeclared_Chunks
         If_Not_Isdeclared_Chunks = Split(ScriptFunction, "If Not Isdeclared(", , vbTextCompare)(1)
         If Err = 0 Then
            On Error GoTo 0
            If_Not_Isdeclared_Chunks = Split(If_Not_Isdeclared_Chunks, vbCrLf)(0)
            isIf_EndIF_Type = 0 = StrComp(Right(If_Not_Isdeclared_Chunks, 5), " Then", vbTextCompare)
            
            
            
            If isIf_EndIF_Type Then
               obfuSection = strCropAndDelete(ScriptFunction, "If Not Isdeclared(", "EndIF" & vbCrLf)
'            Else
'               obfuSection = strCropAndDelete(ScriptFunction, "If Not Isdeclared(", vbCrLf)
            End If
            
            
         Else
            On Error GoTo 0
            obfuSection = ""
         End If
      End If
   
         
    ' Statistic Counter for String in Current function
      Stat_NumLocal = 0
      
      Do
         
         'Process Local String Section
          Dim GlobalStringDef, CommandSeperateLines$
          
          If i = 0 Then
            CommandSeperateLines = strCropAndDelete(ScriptFunction, "global ", vbCrLf)
            ReplaceDo CommandSeperateLines, "Const ", "", , , vbTextCompare
          Else
          
            If isIf_EndIF_Type = False Then
                           'If_Not_Isdeclared_Chunks
               obfuSection = strCropAndDelete(ScriptFunction, "If Not Isdeclared(", vbCrLf) & vbCrLf
               
               CommandSeperateLines = strCropAndDelete(obfuSection, "Global Const ", vbCrLf)
            Else
               CommandSeperateLines = strCropAndDelete(obfuSection, "Global ", vbCrLf)
            End If

            
          End If
       
       '  Now CommandSeperateLines contains this:
       ' "$SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])"
      
      
    ' EXIT DO: if CommandSeperateLines dont contains sth like $Os[0x7] it's no obfu definition
'      If InStr(CommandSeperateLines, "Const ") Then Exit Do
      If InStr(CommandSeperateLines, "=0x01") Then Exit Do
      If CommandSeperateLines = "" Then Exit Do
      If InStr(CommandSeperateLines, fnStringDecrypt) = 0 Then Exit Do


          
          GlobalStringDef = Split(CommandSeperateLines, ",")
         'GlobalStringDef[0] = $SA3160B01D3E=A0000006239($Os[0x7])
         'GlobalStringDef[0] = $SA1D60C0011E=A0000006239($Os[0x8])
         '....
         
        'for all items in GlobalStringDef do...
         Dim LocalStringName
         For Each LocalStringName In GlobalStringDef
             
             'Sepertate GlobalStringDef further...
             Dim LocalStringName_Splitted
             LocalStringName_Splitted = Split(LocalStringName, "=")
             'LocalStringName_Splitted    with         $SA3160B01D3E=A0000006239($Os[0x7])
             'LocalStringName_Splitted[0] = $SA3160B01D3E
             'LocalStringName_Splitted[1] = A0000006239($Os[0x7])
             
             ReDim Preserve LocalStringNames(1 To UBound(LocalStringNames) + 1)
             
             Dim index
             index = UBound(LocalStringNames)
             'CLng("&h" & strCrop1(LocalStringName_Splitted(1), "$Os[0x", "])"))
             LocalStringNames(index) = Trim(LocalStringName_Splitted(0))
            'LocalStringNames(&h7) = "$SA3160B01D3E"
            
             StringTBL.Add HexStringToString(Trim(strCrop1(LocalStringName_Splitted(1), """", """")))
            
            
            
             If i = 0 Then
               ' Store for later search&replace in mainfunction code between functions
                 MainIndexes.Add index
'                 Debug.Print LocalStringNames(index), StringTBL(index), index
             End If

              
           'Replace Local Strings
          ' In ScriptFunction code Replace all "$SA3160B01D3E" with for example "Wrong Name Entered!"
          DeObfu_ReplaceStrings ScriptFunction, CStr(LocalStringNames(index)), StringTBL(index)
'            tmpstr = Replace(tmpstr, "Execute(" & LocalStringNames(index) & ")", StringTBL(index))
'            tmpstr = Replace(tmpstr, "Number(" & LocalStringNames(index) & ")", FormatNumber(StringTBL(index)))
'            tmpstr = Replace(tmpstr, CStr(LocalStringNames(index)), """" & Replace(StringTBL(index), """", """""") & """")
            
            DoEventsSeldom
              
         Next
         


      Stat_NumLocal = Stat_NumLocal + UBound(GlobalStringDef)
      Loop While True
      
      Log "  Local strings: " & Stat_NumLocal
      Stat_NumAtAll = Stat_NumAtAll + Stat_NumLocal
    
    
    ' Save New Function
      NewScript2(i) = ScriptFunction
   
      GUI_Show_Text ScriptFunction
   
     
   Inc i
   Next
  
  
  
  
Log "Strings at all: " & Stat_NumAtAll
Log "strings in Tbl: " & StringTBL.Count
  
' Cut of last Three Functions that where for decrypting
  ReDim Preserve NewScript2(UBound(NewScript2) - 3)
  
  
   NewScript = Join(NewScript2, vbCrLf & "Func ")
'  Stop
  
  '=== Replace Functionname
   DeObfu_ReplaceFunctionNames NewScript


  '=== Replace Global Strings
  
  Log "Restoring global strings..."
  For Each index In MainIndexes
'           For index = 1 To StringTBL.Count
'Debug.Assert LocalStringNames(index) <> "$SA5C9AB03457"
      DeObfu_ReplaceStrings NewScript, CStr(LocalStringNames(index)), StringTBL(index)

  Next
                                 
  
  'Used String
'  Dim i&
   Log "UnAssigned/Unused strings [Index, Value]: "
   For i = 1 To StringTBL.Count
      If LocalStringNames(i) = "" Then
         Log "  " & H16(i) & " " & StringTBL(i)
      End If
   Next
   

  File.CloseFile
  
End Sub



Private Function FormatNumber$(Expr$)
   On Error GoTo err_FormatNumber
     'Set default incase of error
      FormatNumber = Expr
'      If IsNumeric(Expr) Then
         
        'Beautify Numbers
         If (Expr > 65535) Or (Expr < 0) Then
            FormatNumber = "0x" & H32(Expr)
            
         ElseIf Expr > 10 Then
            FormatNumber = "0x" & H16(Expr)
            
         Else
            '0...10'
'            FormatNumber = Expr
            
         End If
         
'      Else
'         FormatNumber = Expr
'      End If

err_FormatNumber:
End Function
