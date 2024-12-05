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
Option Explicit

Dim DeObfu_FunctionNames As New Collection
Dim NewScript$

Sub log(text$)
   FrmMain.log text
End Sub

Sub log2(text$)
   FrmMain.log2 (text)
End Sub

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
  
   Dim i&, line$, EncString$
   Dim VarNames As New Collection
   Dim VarValues As New Collection

   For i = 3 To UBound(globs)
      Dim tmpstr$
      tmpstr = globs(i)
      line = strCropAndDelete(tmpstr, "", vbCrLf)
      globs(i) = tmpstr
   

      
      Dim Words
      Words = Split(line, "'")
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
   FrmMain.Txt_Script = NewScript
   
   log "Search'n'replace strings: "
   
'  ==== Replace Strings =====
   
   Dim ItemIndex&
   For ItemIndex = 1 To VarNames.Count
'       log VarNames(ItemIndex) & "   " & VarValues(ItemIndex)
      
      NewScript = Replace(NewScript, CStr("$" & VarNames(ItemIndex)), CStr("""" & Replace(VarValues(ItemIndex), """", """""") & """"))
      
      log CStr("$" & VarNames(ItemIndex)) & "    " & """" & Replace(VarValues(ItemIndex), """", """""") & """"

   Next
   
   FrmMain.Txt_Script = NewScript
   
   log "Processing function names and arguments:"
   
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


Public Function DetectScriptObfuscated() As Integer
'   FileName = "I:\!Cracks & Projects\AutoIt3\Obfuscator\smarty_Obfuscated.au3"
log "Trying to DeObfuscate : " & FileName
   
 
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
   If InStr(NewScript, "Func A2C4558BC554") Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE
      
   ElseIf InStr(NewScript, "Int(99/3+15*100/4-13^2+81/3-17-245+99/3+15*100/4-13^2+81/3-17)") Then
      DetectScriptObfuscated = DE_OBFUSC_TYPE_ENCODEIT
      
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
         
      Case DE_OBFUSC_TYPE_VANZANDE
         If InStr(1, NewScript, "lobal $Os") Then
            log "Deobfuscating van Zande 1.0.15..."
            DeObfuscate_VanZande1_0_15
         Else
            log "Deobfuscating van Zande 1.0.14..."
            DeObfuscate_VanZande1_0_14
         End If
         
      Case DE_OBFUSC_TYPE_ENCODEIT
         log "Deobfuscating Encodeit 2.0..."
         DeObfuscate_Encodeit
   
   End Select
   
   SaveScriptData NewScript
   log "Deobfucation succeed."
End Sub


Private Sub DeObfuscate_VanZande1_0_15()

'=== Replace Dim Const and Globals
   DeObfu_ReplaceDimsAndConst NewScript
'   DeObfu_ReplaceGlobals NewScript
   
   FrmMain.Txt_Script = NewScript

   log "Okay. Obfucated script loaded and displayed."
   
'=== Prepare ===

    Dim FnNameLoadTBL$
    FnNameLoadTBL$ = strCropAndDelete(NewScript, "global $Os" & vbCrLf, "()")
  
  
    Dim FnNameHexToString$
    FnNameHexToString = strCropWithSeek(NewScript, "=", "($", vbCrLf & "global ")
   
  
 log "Extracting TBLFileName from Script."
  ' GetStringfile(*.tbl) from obfuscate script
    Dim StringTBLFileName As New ClsFilename
    StringTBLFileName.Path = FileName.Path
    StringTBLFileName.NameWithExt = strCrop(NewScript, "FileInstall('", "',Execute(")
 
 log "Loading StringTBLFileName: " & StringTBLFileName.FileName & "..."
    
    Dim StringTBLFile As New FileStream
    StringTBLFile.Create StringTBLFileName.FileName
    
   'Read Stringtable
    Dim StringTBL As New Collection
    Dim line
    StringTBLFile.Position = 0
    For Each line In Split(StringTBLFile.FixedString(-1), "O")
       StringTBL.Add HexStringToString(line)
    Next
    StringTBLFile.CloseFile
   
   
log StringTBL.Count & " strings found."
 
  
  
  
log "Restoring StringNames in ..."
  Dim LocalStringNames
  ReDim LocalStringNames(StringTBL.Count)
  
  ' get local stringvars like
  'global $SA220010601E=A490000263F($Os[0x1]),$SA3A00204A08=A490000263F($Os[0x2]),$SA200030463A=A490000263F($Os[0x3])
   
   Dim MainIndexes As New Collection

   Dim Stat_NumLocal&, Stat_NumAtAll&
   Stat_NumAtAll = 0

  
   Dim EndToken$, NewScript2
   EndToken = "EndFunc" & vbCrLf
  
   NewScript2 = Split(NewScript, "Func ")
   Dim i&
   Dim ScriptFunction: i = 0
   For Each ScriptFunction In NewScript2
      
   
    ' Get Current FunctionName
      Dim FunctionName$
      If i = 0 Then
         FunctionName = "Main"
      Else
        'Check for Mainfunction chunks between FUNC and ENDFUNC
         Dim afterEndFunc$
         afterEndFunc = Split(CStr(ScriptFunction), EndToken)(1)
         If Replace(CStr(afterEndFunc), vbCrLf, "") <> "" Then
           '... apply search& replace here
           
           Dim index
           For Each index In MainIndexes
'           For index = 1 To StringTBL.Count
            ' Restore String
              Dim tmpstr$
              tmpstr = ScriptFunction
              tmpstr = Replace(tmpstr, "Execute(" & LocalStringNames(index) & ")", StringTBL(index))
              tmpstr = Replace(tmpstr, "Number(" & LocalStringNames(index) & ")", StringTBL(index))
              tmpstr = Replace(tmpstr, CStr(LocalStringNames(index)), """" & Replace(StringTBL(index), """", """""") & """")
              ScriptFunction = tmpstr

           Next
           
           
           
         End If
         FunctionName = strCrop1(ScriptFunction, "", "(")
         
          ' Store for later
         DeObfu_FunctionNames.Add FunctionName
   
        tmpstr = ScriptFunction
       
       ' === Replace Locals ====
         DeObfu_ReplaceLocals tmpstr
        
       ' === Replace FunctionArguments  ====
         DeObfu_ReplaceFnArgs tmpstr, FunctionName
         
         ScriptFunction = tmpstr
         
         
      End If
      
 ' Exit if FnNameLoadTBL is reached
   If FunctionName = FnNameLoadTBL Then Exit For
      
      

   log "FunctionName: " & FunctionName
   'Debug.Assert "A5E30E02006" <> FunctionName

     
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
         tmpstr = ScriptFunction
         obfuSection = strCropAndDelete(tmpstr, "If Not Isdeclared(", "EndIF" & vbCrLf)
         ScriptFunction = tmpstr
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
      If InStr(CommandSeperateLines, "$Os[0x") = 0 Then Exit Do
          
          If i = 0 Then ScriptFunction = obfuSection
          
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
             
             index = CLng("&h" & strCrop1(LocalStringName_Splitted(1), "$Os[0x", "])"))
             LocalStringNames(index) = LocalStringName_Splitted(0)
            'LocalStringNames(&h7) = "$SA3160B01D3E"
            
             If i = 0 Then
               ' Store for later search&replace in mainfunction code between functions
                 MainIndexes.Add index
'                 Debug.Print LocalStringNames(index), StringTBL(index), index
             End If

          
            tmpstr = ScriptFunction
              
           'Replace Strings
          ' In ScriptFunction code Replace all "$SA3160B01D3E" with for exsample "Wrong Name Entered!"
            tmpstr = Replace(tmpstr, "Execute(" & LocalStringNames(index) & ")", StringTBL(index))
            tmpstr = Replace(tmpstr, "Number(" & LocalStringNames(index) & ")", StringTBL(index))
            tmpstr = Replace(tmpstr, CStr(LocalStringNames(index)), """" & Replace(StringTBL(index), """", """""") & """")
            
            ScriptFunction = tmpstr
              
         Next

      Stat_NumLocal = Stat_NumLocal + UBound(GlobalStringDef)
      Loop While True
      
      log "  Local strings: " & Stat_NumLocal
      Stat_NumAtAll = Stat_NumAtAll + Stat_NumLocal
    
    ' Save New Function
      NewScript2(i) = ScriptFunction
   
     
   i = i + 1
   Next
  
log "Strings at all: " & Stat_NumAtAll
log "strings in Tbl: " & StringTBL.Count
  
' Cut of rest after FnNameLoadTBL
  ReDim Preserve NewScript2(i - 1)
  
  
   NewScript = Join(NewScript2, vbCrLf & "Func ")
'  Stop
  
  
  '=== Replace Functionname
   DeObfu_ReplaceFunctionNames NewScript


  
  'Used String
'  Dim i&
   log "UnAssigned/Unused strings [Index, Value]: "
   For i = 1 To StringTBL.Count
      If LocalStringNames(i) = "" Then
         log "  " & H16(i) & " " & StringTBL(i)
      End If
   Next
  
  File.CloseFile
  
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
         If Not (LocalDefLine Like "*" & IgnoreLinesWith & "*") Then
         
           
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
log "Renaming Globals..."
   z_DeObfu_ReplaceX CodeSnipped, "Global", "g", IgnoreLinesWith
End Sub
Private Sub DeObfu_ReplaceDimsAndConst(ByRef CodeSnipped$)
log "Renaming Dims..."
   z_DeObfu_ReplaceX CodeSnipped, "Dim", "gDim"
log "Renaming Consts..."
   z_DeObfu_ReplaceX CodeSnipped, vbCrLf & "Const", "gConst"
End Sub

Private Sub DeObfu_ReplaceLocals(ByRef CodeSnipped$)
'log "Renaming Locals..."
   z_DeObfu_ReplaceX CodeSnipped, "Local", ""
End Sub




Private Sub DeObfuscate_VanZande1_0_14()

   'Merge lines with _ at the end
    NewScript = Replace(NewScript, "_" & vbCrLf, "")



'  log "Running 'Tidy.exe " & FileName.NameWithExt & "' to Fix Upper&LowerCase Spelling."
'
'
'  File.Create FileName.FileName & ".tmp", True, True, False
'  File.FixedString(-1) = NewScript
''  File.CloseFile
'
'  Dim cmdline$, parameters$, Logfile$
'  cmdline = App.Path & "\Tidy\Tidy.exe"
'  parameters = """" & FileName & ".tmp" & """ /proper 1" ' /KeepNVersions=1
'  log cmdline & " " & parameters
'  ShellEx cmdline, parameters
'
'
''  File.Create FileName.FileName & ".tmp"
'  File.Position = 0
'  NewScript = File.FixedString(-1)
'  File.CloseFile

'Stop


   FrmMain.Txt_Script = NewScript

   log "Okay. Obfucated script loaded and displayed."
  
  
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
   DeObfu_ReplaceGlobals NewScript, fnStringDecrypt


log "Restoring StringNames in ..."
  
  
   Dim i&
   NewScript2 = Split(NewScript, "Func ", , vbTextCompare)
  
   
   
   Dim ScriptFunction: i = 0
   For Each ScriptFunction In NewScript2
 
    ' Get Current FunctionName
      Dim FunctionName$
      
    ' is first function chunk
      If i = 0 Then
         FunctionName = "Main"
      Else
        'Check for Mainfunction chunks between FUNC and ENDFUNC
         Dim afterEndFunc$
         afterEndFunc = Split(CStr(ScriptFunction), "EndFunc" & vbCrLf, , vbTextCompare)(1)
         If Replace(CStr(afterEndFunc), vbCrLf, "") <> "" Then
           '... apply search& replace here
           
           Dim index
           For Each index In MainIndexes
'           For index = 1 To StringTBL.Count
            ' Restore String

              Dim tmpstr$
              tmpstr = ScriptFunction
              tmpstr = Replace(tmpstr, "Execute(" & LocalStringNames(index) & ")", StringTBL(index))
              tmpstr = Replace(tmpstr, "Number(" & LocalStringNames(index) & ")", StringTBL(index))
              tmpstr = Replace(tmpstr, CStr(LocalStringNames(index)), """" & Replace(StringTBL(index), """", """""") & """")
              ScriptFunction = tmpstr

           Next
                                 
         End If '//Mainfunction
         
         
         FunctionName = strCrop1(ScriptFunction, "", "(")
         
          ' Store for later
         DeObfu_FunctionNames.Add FunctionName
   
        tmpstr = ScriptFunction
       
       ' === Replace Locals ====
         DeObfu_ReplaceLocals tmpstr
        
       ' === Replace FunctionArguments  ====
         DeObfu_ReplaceFnArgs tmpstr, FunctionName
         
         ScriptFunction = tmpstr
         
         
      End If
      
 ' Exit if FnNameLoadTBL is reached
'   If FunctionName = FnNameLoadTBL Then Exit For
      
      

   log "FunctionName: " & FunctionName
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
            
            
            tmpstr = ScriptFunction
            
            If isIf_EndIF_Type Then
               obfuSection = strCropAndDelete(tmpstr, "If Not Isdeclared(", "EndIF" & vbCrLf)
'            Else
'               obfuSection = strCropAndDelete(tmpstr, "If Not Isdeclared(", vbCrLf)
            End If
            
            ScriptFunction = tmpstr
            
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
            tmpstr = ScriptFunction
            CommandSeperateLines = strCropAndDelete(tmpstr, "global ", vbCrLf)
            ScriptFunction = tmpstr
            ReplaceDo CommandSeperateLines, "Const ", "", , , vbTextCompare
          Else
          
            If isIf_EndIF_Type = False Then
                           'If_Not_Isdeclared_Chunks
               tmpstr = ScriptFunction
               obfuSection = strCropAndDelete(tmpstr, "If Not Isdeclared(", vbCrLf) & vbCrLf
               ScriptFunction = tmpstr
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
             index = UBound(LocalStringNames)
             'CLng("&h" & strCrop1(LocalStringName_Splitted(1), "$Os[0x", "])"))
             LocalStringNames(index) = LocalStringName_Splitted(0)
            'LocalStringNames(&h7) = "$SA3160B01D3E"
            
             StringTBL.Add HexStringToString(Trim(strCrop1(LocalStringName_Splitted(1), """", """")))
            
            
            
             If i = 0 Then
               ' Store for later search&replace in mainfunction code between functions
                 MainIndexes.Add index
'                 Debug.Print LocalStringNames(index), StringTBL(index), index
             End If

          
            tmpstr = ScriptFunction
              
           'Replace Strings
          ' In ScriptFunction code Replace all "$SA3160B01D3E" with for exsample "Wrong Name Entered!"
            tmpstr = Replace(tmpstr, "Execute(" & LocalStringNames(index) & ")", StringTBL(index))
            tmpstr = Replace(tmpstr, "Number(" & LocalStringNames(index) & ")", StringTBL(index))
            tmpstr = Replace(tmpstr, CStr(LocalStringNames(index)), """" & Replace(StringTBL(index), """", """""") & """")
            
            ScriptFunction = tmpstr
            
            DoEventsSeldom
              
         Next

      Stat_NumLocal = Stat_NumLocal + UBound(GlobalStringDef)
      Loop While True
      
      log "  Local strings: " & Stat_NumLocal
      Stat_NumAtAll = Stat_NumAtAll + Stat_NumLocal
    
    ' Save New Function
      NewScript2(i) = ScriptFunction
   
      FrmMain.Txt_Script = ScriptFunction
   
     
   Inc i
   Next
  
  
  
  
log "Strings at all: " & Stat_NumAtAll
log "strings in Tbl: " & StringTBL.Count
  
' Cut of last Three Functions that where for decrypting
  ReDim Preserve NewScript2(UBound(NewScript2) - 3)
  
  
   NewScript = Join(NewScript2, vbCrLf & "Func ")
'  Stop
  
  '=== Replace Functionname
   DeObfu_ReplaceFunctionNames NewScript



  
  'Used String
'  Dim i&
   log "UnAssigned/Unused strings [Index, Value]: "
   For i = 1 To StringTBL.Count
      If LocalStringNames(i) = "" Then
         log "  " & H16(i) & " " & StringTBL(i)
      End If
   Next
   

  File.CloseFile
  
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



