Attribute VB_Name = "GetCamo"
Option Explicit

Dim CamoGetDoDebug As Boolean
Dim CamoPattern As New clsStrCat
Dim myRegExp  As New RegExp
Dim filedata As New StringReader
Dim pattern_DebugInfo$


'Converts a string to Hex reg expression
Private Function z_szToGREPHex(isUniCode As Boolean, ParamArray Args())
   
   Dim Texts
   If UBound(Args) = 0 Then
      Texts = Args(0)
   Else
      Texts = Args
   End If
   
   Dim seperator$
   seperator = "\x"
   
   Dim ret As New clsStrCat
   Dim Data() As Byte
   
   Dim Text
   For Each Text In Texts
      Data = Text & IIf(isUniCode, vbNullChar, "")
   
      Dim i&
      For i = LBound(Data) To UBound(Data) Step IIf(isUniCode, 1, 2)
         ' ... due to unsafe Unicode convert
         If Not (isUniCode) Then Debug.Assert Data(i + 1) = 0
         ret.Concat seperator & H8(Data(i))
      Next
   Next
   
   z_szToGREPHex = ret.value
End Function
Public Function szToUnicodeGREPHex(ParamArray Texts() As Variant)
   szToUnicodeGREPHex = z_szToGREPHex(True, Texts)
End Function

Public Function sToGREPHex(ParamArray Texts())
   sToGREPHex = z_szToGREPHex(False, Texts)
End Function

'Public Function REszToGREPHex(Text_RE, Optional Seperator = "\x")
'   Dim Texts()
'   For Each item In Text_RE
'
'      Dim RE
'      RE = False
'      RE = item(2)
'
'      If RE Then
'         Text = RE
'
'      Else
'
'         Dim Text
'         Text = item(1)
'
'      End If
'
'
'      ArrayAdd Texts
'
'   Next
'   szToUnicodeGREPHex = szToGREPHex(Texts)
'
'End Function


Sub pattern_New()
   CamoPattern.Clear
End Sub


Sub pattern_SimpleAdd(SearchText As String)
   CamoPattern.Concat SearchText
   
   CamoGet_DebugCheck
End Sub

Sub pattern_AddNullBytesPadding()
   pattern_Add vbNullChar, "*", False
End Sub



Sub pattern_Add(SearchText As String, Optional quantifier = "", Optional isUniCode = True)
   Dim RE As String
   If isUniCode Then
      RE = szToUnicodeGREPHex(SearchText)
   Else
      RE = sToGREPHex(SearchText)
   End If
   
   pattern_SimpleAdd RE_Group_NonCaptured(RE) + quantifier
   
   
   CamoGet_DebugCheck
End Sub


Sub pattern_Getter(HowManyChars&)
   pattern_SimpleAdd RE_Group(RE_AnyCharRepeat(HowManyChars, HowManyChars))
End Sub


Function pattern_GetMatches(CamoPattern) As Match

   myRegExp.Pattern = CamoPattern
   
   Dim matches As MatchCollection
   Set matches = myRegExp.execute(filedata.Data)
   
   If matches.Count > 1 Then
      Log pattern_DebugInfo & " - is ambigious it has " & matches.Count & " matches"
   End If
   
   If matches.Count >= 1 Then
      Set pattern_GetMatches = matches(0)
   Else
      log_verbose pattern_DebugInfo & " - no match found."
   End If
   
End Function




Function SubMatchOffset(Match As Match, SubMatchIndex&) As Long
   With Match
      Dim Offset&
      Offset = InStr(.value, .SubMatches(SubMatchIndex))
      If RangeCheck(Offset, .Length) Then
         SubMatchOffset = Offset + .FirstIndex - 1
      Else
        'Outside range - normally this should not happen
         Stop
      End If

   End With

End Function
Sub CamoGet_DebugCheck()
   If CamoGetDoDebug Then
      myRegExp.Pattern = CamoPattern
      Dim matches  As MatchCollection
      Set matches = myRegExp.execute(filedata.Data)
      If matches.Count = 0 Then
            Debug.Print "CamoGet: Pattern'" + pattern_DebugInfo + "' (" & Len(myRegExp.Pattern) & ") failed."
            Debug.Print myRegExp.Pattern
            Stop
            'CamoPattern.Clear '!!! just for convient debugging
      End If
   End If
End Sub



Public Sub CamoGet()


   CamoGetDoDebug = False 'True

   'Dim FileName$
   'Frm_Options.Txt_GetCamoFileName = "C:\Tools\MATE\tes\Unc3nZureD\DecompileME.exe_0x401000-0xbc000.bin"
   
   log_verbose "GetCamo's: LoadingFile: " & Frm_Options.Txt_GetCamoFileName
   

   filedata = FileLoad(Frm_Options.Txt_GetCamoFileName)
    
   
   'If Not CamoGetDoDebug Then _
   'On Error Resume Next


' .rdata
'                                            25 00 30 00  32 00               % 0 2
'     64 00 00 00  72 00 62 00  00 00 00 00  77 00 2B 00  62 00   d   r b     w + b
'     00 00 45 41  30 36 00 00  00 00 25 30  32 58 00 00  00 00     EA06    %02X
'     41 55 33 21  00 00 00 00  61 00 75 00  74 00 00 00  2A 00   AU3!    a u t   *
'     00 00 77 00  62 00 00 00  00 00 46 49  4C 45 00 00  00 00     w b     FILE
'     41 00 42 00  53 00 00 00                                    A B S

'Newer Version
'000B22E0                            30 00 32 00 64 00 00 00           0 2 d
'000B22F0   44 00 65 00 66 00 61 00  75 00 6C 00 74 00 00 00   D e f a u l t
'000B2300   77 00 2B 00 62 00 00 00  45 41 30 36 00 00 00 00   w + b   EA06
'000B2310   61 00 75 00 74 00 00 00  77 00 62 00 00 00 00 00   a u t   w b
'000B2320   46 49 4C 45 00 00 00 00  57 6F 77 36 34 52 65 76   FILE    Wow64Rev
'
'3.3.8.1
'00000000   00 00 00 72 00 62 00 00      r b
'00000008   00 00 00 77 00 2B 00 62      w + b
'00000010   00 00 00 45 41 30 36 00      EA06
'00000018   00 00 00 25 30 32 58 00      %02X
'00000020   00 00 00 41 55 33 21 00      AU3!
'00000028   00 00 00 61 00 75 00 74      a u t
'00000030   00 00 00 2A 00 00 00 77      *   w
'00000038   00 62 00 00 00 00 00 46    b     F
'00000040   49 4C 45 00 00 00 00 41   ILE    A
'00000048   00 42 00 53 00 00 00 63    B S   c
'00000050   00 6C 00 6F 00 73 00 65    l o s e
' See MATE/Doc/regexp.htm
    
    

   

   myRegExp.IgnoreCase = False
   myRegExp.Global = False
   myRegExp.MultiLine = True
   Dim Match  As Match
   
   pattern_New
   
   pattern_DebugInfo = ".text#2 AU3! for validation"
   
   '  8D8C24 A4000000        LEA     ECX, [ESP+A4]
   '  E8 C27DFFFF            CALL    004046CE
   '  817C24 60 41553321     CMP     [DWORD ESP+60], 21335541
   '  0F84 AF670600          JE      004730C9
   ' pattern_SimpleAdd "\xE8...\xFF"
   pattern_SimpleAdd "\x81\x7C\x24."
   pattern_Getter 4
   pattern_SimpleAdd "\x0F"
   
   Set Match = pattern_GetMatches(CamoPattern)
   If Not Match Is Nothing Then
      Dim AU3_SubTypeText$
      AU3_SubTypeText = Match.SubMatches(0)
      
      Dim AU3_SubType$
      If AU3_SubTypeText <> AU3_SubType Then
         Log "GetCamo ERROR: AU3_SubType that I got from .text is different from the on in the .data section."
         Log "AU3_SubType_Text: " & ToHexStr(AU3_SubTypeText) & " <> AU3_SubType_data: " & ToHexStr(AU3_SubType)
      Else
         log_verbose ("Alternative AU3_SubType in .text matches with the one in .data")
      End If
      
   End If



   pattern_New
   
  '#1 AU3_SubType
   pattern_DebugInfo = ".rdata#1 EA06"
   
   pattern_Add "%02d", "?"
   pattern_Add "rb", "?"
   pattern_Add ""
   pattern_Add "w+b"
                  
   pattern_Getter 4
   pattern_AddNullBytesPadding
                  
  '#2 AU3_Type ("AU3!")
   pattern_DebugInfo = ".rdata#2 AU3!"
   
   pattern_Add "%02X", "?", False
   pattern_AddNullBytesPadding
   pattern_Getter 4
   pattern_AddNullBytesPadding

  '#3 AU3_ResTypeFile
   pattern_DebugInfo = ".rdata#3 FILE"
   pattern_Add "aut"
   pattern_Add "*", "?"
   pattern_Add "wb"
   pattern_AddNullBytesPadding
   pattern_Getter 4
   
 ' Do #1 EA06 , #2 AU3! and #3 FILE
   Set Match = pattern_GetMatches(CamoPattern)
   
   If Not Match Is Nothing Then
       
      Dim mymatch As SubMatches
      'Set mymatch = Match.SubMatches
      
      
      'Dim AU3_SubType$
      AU3_SubType = Match.SubMatches(0)
         
      Dim AU3_Type$
      AU3_Type = Match.SubMatches(1)
      
      Dim AU3_ResTypeFile$
      AU3_ResTypeFile = Match.SubMatches(2)
      

   End If
   
   '  alternative AU3_SubType is also present in .text section
   '  use it to validate ...
   
   
      With Frm_Options
         .Txt_AU3_SubType_hex = ToHexStr(AU3_SubType)
         log_verbose H32(SubMatchOffset(Match, 0)) & " ->  Found  AU3_SubType: " & .txt_AU3_SubType
         
         .txt_AU3_Type_hex = ToHexStr(AU3_Type)
         log_verbose H32(SubMatchOffset(Match, 1)) & " ->  Found  AU3_Type : " & .txt_AU3_Type
         
         .txt_AU3_ResTypeFile_hex = ToHexStr(AU3_ResTypeFile)
         log_verbose H32(SubMatchOffset(Match, 2)) & " ->  Found  AU3_ResTypeFile :" & .txt_AU3_ResTypeFile
      
      End With
   
   
   
'_____________________________________________________________________________

'.data
'00002088   37 BE 0B B4 A1 8E 0C C3   7¾ ´¡Ž Ã
'00002090   1B DF 05 5A 8D EF 02 2D    ß Z ï -
'00002098   28 58 49 00 00 00 00 00   (XI
'000020A0   1C 58 49 00 01 00 00 00    XI
'000020A8   10 58 49 00 02 00 00 00    XI
'000020B0   00 58 49 00 03 00 00 00    XI
'...
'000020F8   3C 57 49 00 0C 00 00 00   <WI
'00002100   E8 59 49 00 01 00 00 00   èYI
'00002108   28 58 49 00 00 00 00 00   (XI
'00002110   1C 58 49 00 01 00 00 00    XI
'00002118   10 58 49 00 02 00 00 00    XI
'00002120   00 58 49 00 03 00 00 00    XI
'00002128   E0 57 49 00 04 00 00 00   àWI
'...
'00002168   3C 57 49 00 0C 00 00 00   <WI
'00002170   28 58 49 00 00 00 00 00   (XI
'00002178   1C 58 49 00 01 00 00 00    XI
'00002180   10 58 49 00 02 00 00 00    XI
'...
'000021D0   3C 57 49 00 0C 00 00 00   <WI
'000021D8   28 58 49 00 00 00 00 00   (XI
'000021E0   1C 58 49 00 01 00 00 00    XI
'000021E8   10 58 49 00 02 00 00 00    XI
'000021F0   00 58 49 00 03 00 00 00    XI
'000021F8   E0 57 49 00 04 00 00 00   àWI
'00002200   C8 57 49 00 05 00 00 00   ÈWI
'00002208   B8 57 49 00 06 00 00 00   ¸WI
'00002210   A4 57 49 00 07 00 00 00   ¤WI
'00002218   8C 57 49 00 08 00 00 00   ŒWI
'00002220   70 57 49 00 09 00 00 00   pWI
'00002228   58 57 49 00 0A 00 00 00   XWI
'00002230   48 57 49 00 0B 00 00 00   HWI
'00002238   3C 57 49 00 0C 00 00 00   <WI
'00002240   99 4C 53 0A 86 D6 48 7D   ™LS †ÖH}
'00002248   A3 48 4B BE 98 6C 4A A9   £HK¾˜lJ©
'00002250   80 00 00 00 00 00 00 00   €           <= 80-Bytes long  NullByteArray
'00002258   00 00 00 00 00 00 00 00
   pattern_New
   pattern_DebugInfo = ".data#4 £HK..."
   
 ' myRegExp has a stupid bug - it doesn't matches \x00 !!!
 ' ^-so I used '.' instead
   pattern_SimpleAdd ("\x01...\x02...\x03...\x03...........\x07") '\xBE '\x8E
 
'   pattern_SimpleAdd ("\x37.\x0B\xB4\xA1.\x0C\xC3\x1B\xDF\x05\x5A\x8D\xEF\x02\x2D") '\xBE '\x8E
 '  pattern_SimpleAdd RE_Group_NonCaptured(RE_AnyCharRepeat(429, 429))  '("\x99\x4C\x53\x0A\x86\xD6\x48\x7D")
 
   Set Match = pattern_GetMatches(CamoPattern)
   
   If Not Match Is Nothing Then
      filedata.Position = Match.FirstIndex
      log_verbose H32(filedata.Position) & " ->  AU3_Signature seek backwards position..." ' & .txt_AU3Sig

      
      Dim AU3Sig_Hex$
      
    ' Now find begin of 80-Bytes long  NullByteArray...
      filedata.bSearchBackward = True
      filedata.FindByte &H80
'      If filedata.FindString(Chr(&HE) & String(&H44, vbNullChar) & Chr(&HF) & String(3, vbNullChar)) Then
'         AU3Sig_Hex = filedata.FixedString(8)
'      End If
      
      filedata.bSearchBackward = False
      
      If filedata.int32 = 0 Then
         filedata.Move -4
         
         
   '    ' Subpattern
   '      pattern_New
   '      pattern_SimpleAdd ("\x01..." & _
   '                      "\x02..." & _
   '                      "\x03...")
   '      myRegExp.Pattern = CamoPattern
   '
   '
   '
   '      filedata.DisableAutoMove = True
   '
   '    ' get 1KB tmp buffer
   '      Dim filedataTmpBuff As New StringReader
   '      filedataTmpBuff.Data = filedata.FixedString(1024)
   '      Set Match = Nothing
   '      Set Match = myRegExp.Execute(filedataTmpBuff.Data)(0)
   '      filedata.DisableAutoMove = False
   '
   '    ' Seek to 80 00 00 ...
   '      filedataTmpBuff.bSearchBackward = True
   '      filedataTmpBuff.Position = Match.FirstIndex
   '      filedataTmpBuff.FindByte &H80
   '
       ' Seek to au3sig start
         filedata.Move -1 - 2 * 8
          
       ' Read AU3_Signature
         Dim hex1 As New StringReader
         hex1.Data = filedata.FixedString(8)
         
         Dim hex2 As New StringReader
         hex2.Data = filedata.FixedString(8)
         
         AU3Sig_Hex = ValuesToHexString(hex2) & ValuesToHexString(hex1)
         AU3Sig_Hex = RTrim(AU3Sig_Hex)
         
         
         Frm_Options.txt_AU3Sig_Hex = AU3Sig_Hex
         
         log_verbose H32(filedata.Position - 16) & " ->  Found  AU3_Signature: " & AU3Sig_Hex
         
         Frm_Options.Chk_NormalSigScan.value = vbChecked
         
      Else
         log_verbose "Find AU3_Signature failed!"
      
         Frm_Options.Chk_NormalSigScan.value = vbUnchecked

      End If
      
      
   End If
'---------------------------------------------------

 ' 18EE
   pattern_New
   pattern_DebugInfo = ".text#1 0x18EE FILE_DecryptionKey..."

   
   pattern_SimpleAdd "\xE8...\xFF"
   pattern_SimpleAdd (".\xC4.")      ' \x83    ADD     ESP, 10
   pattern_SimpleAdd ("\x68")                ' PUSH
   pattern_Getter 4  '("\x11\x2B\x04\x7F")   '        18EE
   pattern_SimpleAdd ("\x6A\x04")            ' PUSH    4
   pattern_SimpleAdd ("\x8D")                ' LEA     EDX, [EBP-C]  '\x8D\x54\x24 or  8D55 F4
                                          '52  PUSH    EDX

   Set Match = pattern_GetMatches(CamoPattern)
   
   If Not Match Is Nothing Then
   
      Dim tmpstr As New StringReader
      tmpstr.Data = Match.SubMatches(0)
   
      Frm_Options.txt_FILE_DecryptionKey = H32(tmpstr.int32)
      log_verbose H32(SubMatchOffset(Match, 0)) & " ->  Found  AU3_ResourceTypeFILE: " & Frm_Options.txt_FILE_DecryptionKey

   End If
    
'---------------------------------------------------
   pattern_New
   pattern_DebugInfo = ".text#11 0x99F2 FileInst_LenNew..."

   pattern_SimpleAdd "\xE8...\xFF"
   pattern_SimpleAdd (".\xC4.")  ' \x83ADD     ESP, 10
   pattern_SimpleAdd ("\x68")                ' PUSH 99f2
   pattern_Getter 4                              '0x99f2
   pattern_SimpleAdd ("\x6A\x10")            ' PUSH    10
   pattern_SimpleAdd ("\x8D")                ' LEA     EDX, [EBP-C]    '\x54\x24")        '8D55 F4
                                        '  52  PUSH    EDX

   Set Match = pattern_GetMatches(CamoPattern)
   If Not Match Is Nothing Then
   
      tmpstr = Match.SubMatches(0)
      Frm_Options.txtXORKey_MD5PassphraseHashText_DataNew = H32(tmpstr.int32)
   End If
 
 
 
'---------------------------------------------------
   pattern_New
   pattern_DebugInfo = ".text#2 0xB33F FileInst_DataNew..."
 
  ' pattern_SimpleAdd ("\x8B\x06")             ' MOV     EAX, [ESI]
   pattern_SimpleAdd ("\x50")                  ' PUSH    EAX
   pattern_SimpleAdd ("\x81\xF7")              ' XOR     EDI,
   pattern_Getter 4       '("\xBC\xAD\x00\x00")             '0ADBC
   pattern_SimpleAdd ("\x8D\x1C\x3F")          ' LEA     EBX, [EDI+EDI]
   pattern_SimpleAdd ("\x53")                  ' PUSH    EBX
   pattern_SimpleAdd ("\x8D") '\x4C\x24\x38"   ' LEA     ECX, [ESP+38]
                 '00457313    8D  8D  E0FDFFFF   LEA     ECX, [EBP-220]
   pattern_SimpleAdd RE_AnyCharRepeat(3, 5)
  
   pattern_SimpleAdd ("\x6A\x01")            ' PUSH    1
   pattern_SimpleAdd ("\x51")                ' PUSH    ECX
   pattern_SimpleAdd ("\xE8...\xFF")         ' CALL    004151B0FC   \xFF")     '  CALL    004151B0
   pattern_SimpleAdd (".\xC4.")         ' \x83 ADD     ESP, 20
   pattern_SimpleAdd ("\x81\xC7")            '  ADD     EDI,
   pattern_Getter 4                 '("\x3F\xB3\x00\x00") '0B33F
   pattern_SimpleAdd ("\x57")                '  PUSH    EDI
   pattern_SimpleAdd ("\x53")                '  PUSH    EBX
   pattern_SimpleAdd ("\x8D...") '\x54\x24." '  LEA     EDX, [ESP+28]
'  pattern_SimpleAdd ("\x52")                '  PUSH    EDX
   myRegExp.Pattern = CamoPattern
   
   Set Match = pattern_GetMatches(CamoPattern)
   If Not Match Is Nothing Then
      
      tmpstr = Match.SubMatches(0)
      Frm_Options.txtSrcFile_FileInst_LenNew = H32(tmpstr.int32)
      
      tmpstr = Match.SubMatches(1)
      Frm_Options.txtSrcFile_FileInst_DataNew = H32(tmpstr.int32)
      
       log_verbose H32(SubMatchOffset(Match, 0)) & " ->  Found  FileInst_New Data&Len"
   End If
 
'---------------------------------------------------
   pattern_New
   pattern_DebugInfo = ".text#3 0xF479 CompiledPathName..."
 
                                   '8B7C24 28       MOV     EDI, [ESP+28]
      'pattern_SimpleAdd ("\x8B\x16")             ' MOV     EDX, [ESI]
       pattern_SimpleAdd ("\x52")                 ' PUSH    EDX
       pattern_SimpleAdd ("\x81\xF7")             ' XOR     EDI,
       pattern_Getter 4                  '("\x20\xF8\x00\x00")  0F820
       pattern_SimpleAdd ("\x8D\x1C\x3F")         ' LEA     EBX, [EDI+EDI]
       pattern_SimpleAdd ("\x53")                 ' PUSH    EBX
       pattern_SimpleAdd ("\x8D") '\x44\x24.")    ' LEA     EAX, [ESP+40]
       pattern_SimpleAdd RE_AnyCharRepeat(3, 5)
       pattern_SimpleAdd ("\x6A\x01")             ' PUSH    1
       pattern_SimpleAdd ("\x50")                 ' PUSH    EAX
       pattern_SimpleAdd ("\xE8...\xFF")          ' CALL    004151B0
       pattern_SimpleAdd (".\xC4.")           '\x83 ADD     ESP, 28
       pattern_SimpleAdd ("\x81\xC7")            '  ADD     EDI,
       pattern_Getter 4        '("\x79\xF4\x00\x00")             0F479
       pattern_SimpleAdd ("\x57")                  ' PUSH    EDI
   
   Set Match = pattern_GetMatches(CamoPattern)
   If Not Match Is Nothing Then

      tmpstr = Match.SubMatches(0)
      Frm_Options.txtCompiledPathName_LenNew = H32(tmpstr.int32)
      
      tmpstr = Match.SubMatches(1)
      Frm_Options.txtCompiledPathName_DataNew = H32(tmpstr.int32)
      
      filedata.Position = SubMatchOffset(Match, 0)
      log_verbose H32(filedata.Position) & " ->  Found  CompiledPathName Data&Len"
   End If
 
'---------------------------------------------------
'       pattern_SimpleAdd ("\xE8...\xFF")            'E8 11DDFBFF     CALL    0041527B
'       pattern_SimpleAdd ("\x8B.\x08")            '\x8B\x46\x08 MOV     EAX, [ESI+8]
'                   '                 8B4E 08         MOV     ECX, [ESI+8]
'
'       pattern_SimpleAdd (".\xC4\x10")         '\x83 ADD     ESP, 10
'       pattern_SimpleAdd RE_AnyCharRepeat(1, 2) ' ("\x05") ADD     EAX,   | 81C1 77240000   ADD     ECX, 2477
'       pattern_Getter                             '("\x77\x24\x00\x00") ' 2477
'       pattern_SimpleAdd (".") ' ("\x50")            'PUSH    EAX
'       pattern_SimpleAdd ("\x57")                    'PUSH    EDI
'       pattern_SimpleAdd (".") '\x55")               'PUSH    EBP
'       pattern_SimpleAdd ("\xE8...\xFF")

'004578A3    C2 0C00         RETN    0C
'004578A6    8B56 08         MOV     EDX, [ESI+8]
'004578A9    81C2 77240000   ADD     EDX, 2477
'004578AF    52              PUSH    EDX
'004578B0    8D8424 BC060000 LEA     EAX, [ESP+6BC]
'004578B7    50              PUSH    EAX
'004578B8    E8 90F6FEFF     CALL    00446F4D
'

'2477
   pattern_New
   pattern_DebugInfo = ".text#7 0x2477 DecryptionKey_New..."
   
       pattern_SimpleAdd ("\xC2..")                'C2 0C00         RETN    0C
                        ' "‹... =>  "\x8b...
       pattern_SimpleAdd ("‹.\x08")                '8B56 08         MOV     EDX, [ESI+8]
       pattern_SimpleAdd RE_AnyCharRepeat(1, 2)   '  05   77240000  ADD     EAX, 2477 |
                                                 '   81C1 77240000  ADD     ECX, 2477
       pattern_Getter 4 '("\x77\x24\x00\x00") ' 2477
       pattern_SimpleAdd "."                        ' ("\x50")                     'PUSH    EAX | 51              PUSH    ECX
       pattern_SimpleAdd "\x8D" & RE_AnyCharRepeat(5, 6) '  8D8D E0FDFFFF   LEA     ECX, [EBP-220]
       pattern_SimpleAdd "."                       '  "\x50"       '50              PUSH    EAX | EDX
       pattern_SimpleAdd "\xE8...\xFF"
       pattern_SimpleAdd "\x33\xC0"
   
   Set Match = pattern_GetMatches(CamoPattern)
   If Not Match Is Nothing Then
      
      tmpstr = Match.SubMatches(0)
      Frm_Options.txtData_DecryptionKey_New = H32(tmpstr.int32)
      
      filedata.Position = SubMatchOffset(Match, 0)
      log_verbose H32(filedata.Position) & " ->  Found  DecryptionKey: "
      
      
   End If
  
  
'---------------------------------------------------------
'00402851    803408 2F       XOR     [BYTE EAX+ECX], 2F
'00402855    41              INC     ECX
'00402856    3B4D 10         CMP     ECX, [EBP+10]
'00402859  ^ 75 F6           JNZ     SHORT 00402851
'0040285B    E9 A1020000     JMP     00402B01
'80 34 08 2F 41 3B 4D 10 75 F6 E9 A1 02 00 00

   pattern_New
   pattern_DebugInfo = ".text#8 0x2F optional extra XORCryptkey..."
   
   pattern_SimpleAdd (".\x34\x08" & RE_Group(RE_AnyChar))                  '803408 2F       XOR     [BYTE EAX+ECX], 2F
   pattern_SimpleAdd ("\x41\x3B\x4D\x10\x75") '\xF6\xE9\xA1\x02\x00\x00")
       
       
       
   Set Match = pattern_GetMatches(CamoPattern)
   
   Dim XORCryptkey&
   If Not Match Is Nothing Then XORCryptkey = Asc(Match.SubMatches(0))
   If XORCryptkey Then
      
      XORCryptkey = Asc(Match.SubMatches(0))
   
   
      filedata.Position = Match.FirstIndex
      Log H32(filedata.Position) & " ->  " & _
          "XORCryptkey: " & H8x(XORCryptkey) _
          & "    as char '" & Match.SubMatches(0) & "'"
      Log "Custom ReadFileHook with XORCryptkey found !!!"
      
    ' Xor & save as *.a3x
      FileName.Ext = "a3x"
      FileSave FileName.FileName, _
         SimpleXor(filedata.Data, XORCryptkey)
      Log "XOR'ed whole file and saved it to " & FileName.FileName
      
      MsgBox "Press Ok to reload " & FileName.NameWithExt & " now ! ", vbInformation, "Xor decrypt done. "
      
    ' Open File
      FrmMain.Combo_Filename = FileName.FileName
      
   End If

 
'---------------------------------------------------
   
  Frm_Options.CommitChanges
 
End Sub

Public Function ToHexStr(Data As String) As String
   Dim tmp As New StringReader
   tmp.Data = Data
   ToHexStr = RTrim(ValuesToHexString(tmp))
End Function


Public Function SimpleXor(ScriptData$, ByVal Xor_Key&) As String
   
      
      Dim tmpBuff() As Byte
      tmpBuff = DecodeUnicode(ScriptData)
      Dim tmpByte As Byte
      
      Dim StrCharPos&
      For StrCharPos = 0 To UBound(tmpBuff)
         tmpByte = tmpBuff(StrCharPos)
         tmpByte = (tmpByte Xor Xor_Key) And &HFF
         tmpBuff(StrCharPos) = tmpByte
      
         If 0 = (StrCharPos Mod &H8000) Then myDoEvents
         
      Next
      
      SimpleXor = EncodeUnicode(tmpBuff)
      
      
End Function

