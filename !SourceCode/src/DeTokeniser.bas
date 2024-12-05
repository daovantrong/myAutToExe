Attribute VB_Name = "DeTokeniser"
Option Explicit
Const AUTOIT_SourceCodeLine_MAXLEN& = 4096

Const whiteSpaceTerminal$ = " "
Const ExcludePreWhiteSpaceTerminal$ = "(["
Const ExcludePostWhiteSpaceTerminal$ = ")]."

Const TokenFile_RequiredInputExtensions = ".tok .mem"

Dim Atom$, SourceCodeLine$
Dim bAddWhiteSpace As Boolean

Sub DeToken()
   BenchStart
   
   With File
    
      Log "Trying to DeTokenise: " & FileName.FileName
      
      If InStr(TokenFile_RequiredInputExtensions, FileName.Ext) = 0 Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! Required FileExtension for Tokenfiles: '" & TokenFile_RequiredInputExtensions & "'" & vbCrLf & _
         "Rename this file manually to show that this should be detokenied."
      End If
      
      
'      If FrmMain.chk_NoDeTokenise.Value = vbChecked Then
'         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! Enable DeTokenise in Options to use it." & FileName.FileName
'
'      End If
      
      .Create FileName.FileName, False, False, True
      
      
   On Error GoTo DeToken_Err
      .Position = 0
      
      Dim Lines&
      Lines = .longValue
      FL "Code Lines: " & Lines & "   0x" & H32(Lines)
      
    ' File shouldn't start with MZ 00 00 -> ExeFile
    ' &HDFEFFF -> Unicodemarker
      If ((Lines And 65535) = &H5A4D) Or (Lines = &HDFEFF) Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "That's no Au3-TokenFile."
      
      ElseIf ((Lines And &H7FFFFFF) > &H3BFEFF) Then
         'It's highly unlikly that there are more that 16 Mio lines in a Sourcefile
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "This seem to be no Au3-TokenFile."
      End If
      
      
            
      
      FrmMain.List_Source.Clear
      FrmMain.List_Source.Visible = True
   
      
      
      Dim Cmd&
      Dim size&

      Dim SourceCode ' As New Collection
      Dim SourceCodeLineCount&
      ReDim SourceCode(1 To Lines):     SourceCodeLineCount = 1:
      Dim TokenCount&: TokenCount = 0
      
      Dim RawString As StringReader: Set RawString = New StringReader
      Dim DecodeString As StringReader: Set DecodeString = New StringReader

      SourceCodeLine = ""
      Do
   
         Atom = ""
         
         If (SourceCodeLineCount > Lines) Then
            Exit Do
         End If
         
         
       ' Default
         bAddWhiteSpace = False
         
       ' Read Token
         Cmd = .ByteValue
         Inc TokenCount
         
       ' Log it ''" & Chr(Cmd) & "'
         FL_verbose "Token: " & H8(Cmd) & "      (Line: " & SourceCodeLineCount & "  TokenCount: " & TokenCount & ")"
'         If RangeCheck(SourceCodeLineCount, 8716, 8714) Then
''            Stop
'            If FrmMain.Chk_verbose <> vbChecked Then FrmMain.Chk_verbose = vbChecked
'         Else
'           If FrmMain.Chk_verbose <> vbUnchecked Then FrmMain.Chk_verbose = vbUnchecked
'         End If
         
         
         Select Case Cmd
         
'------- Numbers -----------
         Case &H0 To &HF
            '&H5
            Dim int32$
            int32 = .longValue
            Atom = int32
            FL_verbose "Int32: 0x" & H32(int32) & "   " & int32
            
'            Debug.Assert Cmd = 5
         
         Case &H10 To &H1F
            Dim int64 As Currency
            int64 = .int64Value
            'int64 = H32(.longValue)
            'int64 = H32(.longValue) & int64
            'Replace 123,45 -> 12345
            Atom = Replace(CStr(int64), ",", "")
            
            FL_verbose "int64: " & int64
            
            Debug.Assert Cmd = &H10
         
         Case &H20 To &H2F
            
           'Get DoubleValue
            Dim Double_$
            Double_ = .DoubleValue
           
           'Replace 123,11 -> 123.11
            Atom = Replace(CStr(Double_), ",", ".")
         
            FL_verbose "64Bit-float: " & Double_
         
            Debug.Assert Cmd = &H20
         

'------- Strings -----------
         Case &H30 To &H3F
            
           'Get StrLength and load it
            size = .longValue
            FL_verbose "StringSize: " & H32(size)
            RawString = .FixedStringW(size)
           
           'XorDecode String
            Dim pos&, XorKey_l As Byte, XorKey_h As Byte
            
            XorKey_l = (size And &HFF)
            XorKey_h = ((size \ &H100) And &HFF) ' 2^8 = 256
            
            Dim tmpBuff() As Byte
            tmpBuff = RawString
            
            For pos = LBound(tmpBuff) To UBound(tmpBuff) Step 2
               tmpBuff(pos) = tmpBuff(pos) Xor XorKey_l
               tmpBuff(pos + 1) = tmpBuff(pos + 1) Xor XorKey_h
'               DecodeString = tmpBuff
               
               'If 0 = (pos Mod &H8000) Then DoEvents
            Next
            
            DecodeString = tmpBuff
            
'Comment out due to bad performance
'            RawString.Position = 0
'            DecodeString = Space(RawString.Length \ 2)
'            Do Until RawString.EOS
'               DecodeString.int8 = RawString.int8 Xor Size
'               If Not (RawString.EOS) Then Debug.Assert RawString.int8 = 0
'            Loop
            
            
'------- Commands -----------
            Select Case Cmd
            
            Case &H30 'BlockElement (FUNC, IF...) and the Rest of 42 Elements: "AND OR NOT IF THEN ELSE ELSEIF ENDIF WHILE WEND DO UNTIL FOR NEXT TO STEP IN EXITLOOP CONTINUELOOP SELECT CASE ENDSELECT SWITCH ENDSWITCH CONTINUECASE DIM REDIM LOCAL GLOBAL CONST FUNC ENDFUNC RETURN EXIT BYREF WITH ENDWITH TRUE FALSE DEFAULT ENUM NULL"
               FL_verbose """" & DecodeString.Data & """   Type: BlockElement"
               
               Atom = DecodeString
               bAddWhiteSpace = True
              
              'LineBreak after and before 'Functions'
               If Atom = "ENDFUNC" Then
                  Atom = Atom & vbCrLf
               ElseIf Atom = "FUNC" Then
                  Atom = vbCrLf & Atom
               End If

            
            Case &H31 'FunctionCall with params
               Atom = DecodeString
               FL_verbose """" & DecodeString.Data & """   Type: AutoItFunction"
               
            Case &H32 'Macro
               Atom = "@" & DecodeString
               FL_verbose """" & DecodeString.Data & """   Type: Macro"
            
            Case &H33 'Variable
               Atom = "$" & DecodeString
               FL_verbose """" & DecodeString.Data & """   Type: Variable"
            
            Case &H34 'FunctionCall
               Atom = DecodeString
               FL_verbose """" & DecodeString.Data & """   Type: UserFunction"
            
            Case &H35 'Property
               Atom = "." & DecodeString
               FL_verbose """" & DecodeString.Data & """   Type: Property"
            
            Case &H36 'UserString
               
               Atom = MakeAutoItString(DecodeString.Data)
               FL_verbose """" & DecodeString.Data & """   Type: UserString"
            
            Case &H37 '# PreProcessor
               Atom = DecodeString
               bAddWhiteSpace = True
               FL_verbose """" & DecodeString.Data & """   Type: PreProcessor"
            
            
            Case Else
               'Unknown StringToken
               Stop
            End Select
            
 '           log String(40, "_")
         
'------- Operators -----------
         Case &H40 To &H56
'            Atom = Choose((Cmd - &H40 + 1), ",", "=", ">", "<", "<>", ">=", "<=", "(", ")", "+", "-", "/", "", "&", "[", "]", "==", "^", "+=", "-=", "/=", "*=", "&=")
         '                     Au3Manual AcciChar
            
            Select Case Cmd
               Case &H40: Atom = ","  '        2C
               Case &H41: Atom = "="  ' 1  13  3D
               Case &H42: Atom = ">"  ' 16     3E
               Case &H43: Atom = "<"  ' 18     3C
               Case &H44: Atom = "<>" ' 15     3C
               Case &H45: Atom = ">=" ' 17     3E
               Case &H46: Atom = "<=" ' 19     3C
               Case &H47: Atom = "("  '        28
               Case &H48: Atom = ")"  '        29
               Case &H49: Atom = "+": ' 7      2B
               Case &H4A: Atom = "-": ' 8      2D
               Case &H4B: Atom = "/"  ' 10     2F
               Case &H4C: Atom = "*": ' 9      2A
               Case &H4D: Atom = "&"  ' 11     26
               Case &H4E: Atom = "["  '        5B
               Case &H4F: Atom = "]"  '        5D
               Case &H50: Atom = "==" ' 14     3D
               Case &H51: Atom = "^"  ' 12     5E
               Case &H52: Atom = "+=" '2       2B
               Case &H53: Atom = "-=" '3       2D
               Case &H54: Atom = "/=" '5       2F
               Case &H55: Atom = "*=" '4       2A
               Case &H56: Atom = "&=" '6       26
            End Select
            FL_verbose """" & Atom & """   Type: operator" '   AddWhiteSpace=" & bAddWhiteSpace
            
'------- EOL -----------
         Case &H7F
            'Execute
            
            
            SourceCodeLine = RTrim$(SourceCodeLine)
            
            LogSourceCodeLine SourceCodeLine
            
            log_verbose ">>>  " & SourceCodeLine
            log_verbose String(80, "_")
            log_verbose ""
 
          ' Test Length
            Dim SourceCodeLine_Len&
            SourceCodeLine_Len = Len(SourceCodeLine)
            If SourceCodeLine_Len >= AUTOIT_SourceCodeLine_MAXLEN Then
               Log "WARNING: SourceCodeLine: " & SourceCodeLineCount & " is " & _
               SourceCodeLine_Len - AUTOIT_SourceCodeLine_MAXLEN & " chars longer than " & _
               AUTOIT_SourceCodeLine_MAXLEN & " - Please remove some spaces manually to make it shorter."
            End If
          
          
          ' Add SourceCodeLine to SourceCode
            SourceCode(SourceCodeLineCount) = SourceCodeLine
            Inc SourceCodeLineCount
            
            SourceCodeLine = ""

         Case Else
            
           'Unknown Token
           Log "ERROR: Unknown Token: " & Cmd & " at " & H32(.Position)
           Exit Do
           'qw
           Stop
           

         End Select
         
         'Debug.Assert SourceCodeLineCount < 1021
         
         If bAddWhiteSpace Then
           'Add to SourceLine
            'SourceCodeLine = SourceCodeLine & Atom & AddWhiteSpace
            'SourceCodeLine = SourceCodeLine & Atom & whiteSpaceTerminal
'            If SourceCodeLine = "" Then
'               SourceCodeLine = Atom & whiteSpaceTerminal
'            Else
               SourceCodeLine = SourceCodeLine & whiteSpaceTerminal & Atom & whiteSpaceTerminal
'            End If
            
         Else
           'Add to SourceLine
            SourceCodeLine = SourceCodeLine & Atom
         End If
         
         DoEventsVerySeldom

      Loop Until .EOF
    
Err.Clear
DeToken_Err:
Select Case Err
   Case 0
   Case Else
     Dim ErrText$
     ErrText = "ERROR: " & Err.Description & vbCrLf & _
      "FileOffset: " & H32(.Position) & vbCrLf & _
      "when detokising script line: " & SourceCodeLineCount & vbCrLf & SourceCodeLine
     Log ErrText
     MsgBox ErrText, vbCritical, "Unexpected Error during detokenising"
     
     Resume DeToken_Finally
End Select
DeToken_Finally:
   .CloseFile
  End With
  
BenchEnd
  
  
  
  If FrmMain.Chk_TmpFile = vbUnchecked Then
     Log "Keep TmpFile is unchecked => Deleting '" & FileName.NameWithExt & "'"
     FileDelete (FileName)
  End If
  
  FileName.Ext = ".au3"
  
  
'   If bUnicodeEnable Then
      Dim ScriptData$
      ScriptData = Join(SourceCode, vbCrLf)

'      Dim FileName_UTF16 As New ClsFilename
'      FileName_UTF16.FileName = FileName.FileName
'
'      FileName_UTF16.Name = FileName.Name & "_UTF16"
'      FrmMain.Log "Saving UTF16-Script to: " & FileName_UTF16.FileName
'
'      File.Create FileName_UTF16.FileName, True, False, False
'      File.Position = 0
'      File.FixedString(-1) = UTF16_BOM & ScriptData
'      File.setEOF
'      File.CloseFile
'
'   End If
  
  FrmMain.Log "Converting Unicode to UTF8, since Tidy don't support unicode."
  SaveScriptData UTF8_BOM & EncodeUTF8(ScriptData)
   
  Log "Token expansion succeed."
   
  FrmMain.List_Source.Visible = False


End Sub



Private Sub LogSourceCodeLine(TextLine$)
   If FrmMain.Chk_verbose.value = vbChecked Then
   
      On Error Resume Next
      With FrmMain.List_Source
         .AddItem TextLine
       
       ' Process windows messages (=Refresh display)
         If Rnd < 0.01 Then
             ' Scroll to last item
            .ListIndex = .ListCount - 1
         End If
         
      End With
   End If
End Sub
'Handle UserString with Quotes...
Function MakeAutoItString(RawString$)
             
   ' HasDoubleQuote ?
     If InStr(RawString, """") <> 0 Then
        
      ' HasSingleQuote ?
        If InStr(RawString, "'") <> 0 Then
         ' Scenario3: " This is a 'Example' on correct "Quoting" String "
           MakeAutoItString = """" & Replace(RawString, """", """""") & """"
        Else
         ' Scenario2: " This is a "Example". "
           MakeAutoItString = "'" & RawString & "'"
        End If
     Else
      ' ' Scenario1: " ExampleString "
        MakeAutoItString = """" & RawString & """"
     End If
     

End Function

' Add WhiteSpace Seperator to SourceCodeLine
Function AddWhiteSpace$()
   
   'No WhiteSpace at the Beginning
   If SourceCodeLine = "" Then Exit Function
   
   Dim LastChar$
   LastChar = Right(SourceCodeLine, 1)
   
   Dim NextChar$
   NextChar = Left(Atom, 1)
   
   'Don'Append WhiteSpace in cases like this :
   '"@CMDLIND ["   or   "@CMDLIND [0" <-"].."
   '         (^-PreCase)                (^-PostCase)
   If InStr(1, ExcludePreWhiteSpaceTerminal, LastChar) Or _
      InStr(1, ExcludePostWhiteSpaceTerminal, NextChar) Then
'      Stop
   ElseIf whiteSpaceTerminal <> LastChar Then
         AddWhiteSpace = whiteSpaceTerminal
   End If
   
End Function





Private Sub FL_verbose(Text)
   FrmMain.FL_verbose Text
End Sub
Private Sub log_verbose(TextLine$)
   FrmMain.log_verbose TextLine$
End Sub

Private Sub FL(Text)
   FrmMain.FL Text
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Private Sub Log(TextLine$)
   FrmMain.Log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Private Sub log_clear()
   FrmMain.log_clear
End Sub

