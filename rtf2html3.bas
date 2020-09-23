Attribute VB_Name = "rtf2html"
Option Explicit

Private strCurPhrase As String
Private strHTML As String
Private Codes() As String
Private NextCodes() As String
Private CodesBeg() As String         'beginning codes
Private NextCodesBeg() As String     'beginning codes for next text
Private CodesTmp() As String         'temp stack for copying
Private CodesTmpBeg() As String      'temp stack for copying beg

Public strCR As String           'string to use for CRs - blank if +CR not chosen in options
Private strBeforeText As String
Private strBeforeText2 As String
Private strBeforeText3 As String
Private gPlain As Boolean            'true if all codes shouls be popped before next text
Private strColorTable() As String    'table of colors
Private lColors As Long              '# of colors
Private strFontTable() As String     'table of fonts
Private lFonts As Long               '# of fonts
Private strEOL As String             'string to include before <br>
Private lSkipWords As Long           'number od words to skip from current
Private gBOL As Boolean              'a <br> was inserted but no non-whitespace text has been inserted

Private strFont As String
Private strTable As String
Private strFontColor As String     'current font color for setting up fontstring
Private strFontSize As String      'current font size for setting up fontstring
Private lFontSize As Long

Function ClearCodes()
    ReDim Codes(0)
    ReDim NextCodes(0)
    ReDim CodesBeg(0)
    ReDim NextCodesBeg(0)
End Function


Function ClearFont()
    strFont = ""
    strTable = ""
    strFontColor = ""
    strFontSize = ""
    lFontSize = 0
End Function

Function Codes2NextTill(strCode As String)
    

    Dim l As Long

    l = UBound(Codes)
    While Codes(l) <> strCode And l >= 0
        l = l - 1
    Wend
    CodesBeg(l) = ""
    l = l + 1
    While l <= UBound(Codes)
        PushNext (Codes(l))
        PushNextBeg (CodesBeg(l))
        CodesBeg(l) = ""
        l = l + 1
    Wend
End Function

Function GetColorTable(strSecTmp As String, strColorTable() As String)
    'get color table data and fill in strColorTable array
    Dim lColors As Long
    Dim lBOS As Long
    Dim lEOS As Long
    Dim strTmp As String
    
    lBOS = InStr(strSecTmp, "\colortbl")
    ReDim strColorTable(0)
    lColors = 1
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strSecTmp, ";}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strSecTmp, "\red")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strColorTable(lColors)
                strTmp = Trim(Hex(Mid(strSecTmp, lBOS + 4, 1) & IIf(IsNumeric(Mid(strSecTmp, lBOS + 5, 1)), Mid(strSecTmp, lBOS + 5, 1), "") & IIf(IsNumeric(Mid(strSecTmp, lBOS + 6, 1)), Mid(strSecTmp, lBOS + 6, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strSecTmp, "\green")
                strTmp = Trim(Hex(Mid(strSecTmp, lBOS + 6, 1) & IIf(IsNumeric(Mid(strSecTmp, lBOS + 7, 1)), Mid(strSecTmp, lBOS + 7, 1), "") & IIf(IsNumeric(Mid(strSecTmp, lBOS + 8, 1)), Mid(strSecTmp, lBOS + 8, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strSecTmp, "\blue")
                strTmp = Trim(Hex(Mid(strSecTmp, lBOS + 5, 1) & IIf(IsNumeric(Mid(strSecTmp, lBOS + 6, 1)), Mid(strSecTmp, lBOS + 6, 1), "") & IIf(IsNumeric(Mid(strSecTmp, lBOS + 7, 1)), Mid(strSecTmp, lBOS + 7, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strSecTmp, "\red")
                lColors = lColors + 1
            Wend
        End If
    End If
End Function

Function GetFontTable(strSecTmp As String, strFontTable() As String)
    'get font table data and fill in strFontTable array
    Dim lFonts As Long
    Dim lBOS As Long
    Dim lEOS As Long
    Dim strTmp As String
    
    lBOS = InStr(strSecTmp, "\fonttbl")
    ReDim strFontTable(0)
    lFonts = 0
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strSecTmp, ";}}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strSecTmp, "\f0")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strFontTable(lFonts)
                While ((Mid(strSecTmp, lBOS, 1) <> " ") And (lBOS <= lEOS))
                    lBOS = lBOS + 1
                Wend
                lBOS = lBOS + 1
                strTmp = Mid(strSecTmp, lBOS, InStr(lBOS, strSecTmp, ";") - lBOS)
                strFontTable(lFonts) = strFontTable(lFonts) & strTmp
                lBOS = InStr(lBOS, strSecTmp, "\f" & (lFonts + 1))
                lFonts = lFonts + 1
            Wend
        End If
    End If
End Function


Function InNext(strTmp) As Boolean
    Dim gTmp As Boolean
    Dim l As Long
    
    l = 1
    gTmp = False
    While l <= UBound(NextCodes) And Not gTmp
        If NextCodes(l) = strTmp Then gTmp = True
        l = l + 1
    Wend
    InNext = gTmp
End Function

Function InCodes(strTmp) As Boolean
    Dim gTmp As Boolean
    Dim l As Long
    
    l = 1
    gTmp = False
    While l <= UBound(Codes) And Not gTmp
        If Codes(l) = strTmp And Len(CodesBeg(l)) > 0 Then gTmp = True
        l = l + 1
    Wend
    InCodes = gTmp
End Function


Function NabNextLine(strRTF As String) As String
    Dim l As Long
    
    l = InStr(strRTF, vbCrLf)
    If l = 0 Then l = Len(strRTF)
    NabNextLine = TrimAll(Left(strRTF, l))
    If l = Len(strRTF) Then
        strRTF = ""
    Else
        strRTF = TrimAll(Mid(strRTF, l))
    End If
End Function

Function NabNextWord(strLine As String) As String
    Dim l As Long
    Dim lvl As Integer
    Dim gEndofWord As Boolean
    Dim gInCommand As Boolean    'current word is command instead of plain word
    
    gInCommand = False
    l = 0
    lvl = 0
    'strLine = TrimifCmd(strLine)
    If Left(strLine, 1) = "}" Then
        strLine = Mid(strLine, 2)
        NabNextWord = "}"
        GoTo finally
    End If
    While Not gEndofWord
        l = l + 1
        If l >= Len(strLine) Then
            If l = Len(strLine) Then l = l + 1
            gEndofWord = True
        ElseIf InStr("\{}", Mid(strLine, l, 1)) Then
            If l = 1 And Mid(strLine, l, 1) = "\" Then gInCommand = True
            If Mid(strLine, l + 1, 1) <> "\" And l > 1 And lvl = 0 Then
                gEndofWord = True
            End If
        ElseIf Mid(strLine, l, 1) = " " And lvl = 0 And gInCommand Then
            gEndofWord = True
        End If
    Wend
    
    If l = 0 Then l = Len(strLine)
    NabNextWord = Left(strLine, l - 1)
    While Len(NabNextWord) > 0 And InStr("{}", Right(NabNextWord, 1))
        NabNextWord = Left(NabNextWord, Len(NabNextWord) - 1)
    Wend
    While Len(NabNextWord) > 0 And InStr("{}", Left(NabNextWord, 1))
        NabNextWord = Right(NabNextWord, Len(NabNextWord) - 1)
    Wend
    strLine = Mid(strLine, l)
    If Left(strLine, 1) = " " Then strLine = Mid(strLine, 2)
finally:
End Function

Function NabSection(strRTF As String, lPos As Long) As String
    'grab section surrounding lPos, strip section out of strRTF and return it
    Dim lBOS As Long         'beginning of section
    Dim lEOS As Long         'ending of section
    Dim strChar As String
    Dim lLev As Long         'level of brackets/parens
    Dim lRTFLen As Long
    
    lRTFLen = Len(strRTF)
    
    lBOS = lPos
    strChar = Mid(strRTF, lBOS, 1)
    lLev = 1
    While lLev > 0
        lBOS = lBOS - 1
        If lBOS <= 0 Then
            lLev = lLev - 1
        Else
            strChar = Mid(strRTF, lBOS, 1)
            If strChar = "}" Then
                lLev = lLev + 1
            ElseIf strChar = "{" Then
                lLev = lLev - 1
            End If
        End If
    Wend
    lBOS = lBOS - 1
    If lBOS < 1 Then lBOS = 1
    
    lEOS = lPos
    strChar = Mid(strRTF, lEOS, 1)
    lLev = 1
    While lLev > 0
        lEOS = lEOS + 1
        If lEOS >= lRTFLen Then
            lLev = lLev - 1
        Else
            strChar = Mid(strRTF, lEOS, 1)
            If strChar = "{" Then
                lLev = lLev + 1
            ElseIf strChar = "}" Then
                lLev = lLev - 1
            End If
        End If
    Wend
    lEOS = lEOS + 1
    If lEOS > lRTFLen Then lEOS = lRTFLen
    NabSection = Mid(strRTF, lBOS + 1, lEOS - lBOS - 1)
    strRTF = Mid(strRTF, 1, lBOS) & Mid(strRTF, lEOS)
    strRTF = rtf2html_replace(strRTF, vbCrLf & vbCrLf, vbCrLf)
End Function

Function Next2Codes()
    'move codes from pending ("next") stack to current stack
    Dim lNumCodes As Long
    Dim l As Long
    
    If UBound(NextCodes) > 0 Then
        lNumCodes = UBound(Codes)
        ReDim Preserve Codes(lNumCodes + UBound(NextCodes))
        ReDim Preserve CodesBeg(lNumCodes + UBound(NextCodes))
        For l = 1 To UBound(NextCodes)
            Codes(lNumCodes + l) = NextCodes(l)
            CodesBeg(lNumCodes + l) = NextCodesBeg(l)
        Next l
        ReDim NextCodes(0)
        ReDim NextCodesBeg(0)
    End If
End Function

Function Codes2Next()
    'move codes from "current" stack to pending ("next") stack
    Dim lNumCodes As Long
    Dim l As Long
    
    If UBound(Codes) > 0 Then
        lNumCodes = UBound(NextCodes)
        ReDim Preserve NextCodes(lNumCodes + UBound(Codes))
        ReDim Preserve NextCodesBeg(lNumCodes + UBound(Codes))
        For l = 1 To UBound(Codes)
            NextCodes(lNumCodes + l) = Codes(l)
            NextCodesBeg(lNumCodes + l) = CodesBeg(l)
        Next l
        ReDim Codes(0)
        ReDim CodesBeg(0)
    End If
End Function

Function ParseFont(strColor As String, strSize As String) As String
    Dim strTmpFont As String
    
    strTmpFont = "<font"
    If strColor <> "" Then
       strTmpFont = strTmpFont & " color=""" & strColor & """"
    End If
    If strSize <> "" And strSize <> "2" Then
        strTmpFont = strTmpFont & " size=" & strSize
    End If
    strTmpFont = strTmpFont & ">"
    ParseFont = strTmpFont
End Function

Function PopCode() As String
    If UBound(Codes) > 0 Then
        PopCode = Codes(UBound(Codes))
        ReDim Preserve Codes(UBound(Codes) - 1)
    End If
End Function

Function GetAllCodes() As String
    Dim strTmp As String
    Dim l As Long
    
    strTmp = ""
    If UBound(Codes) > 0 Then
        For l = UBound(Codes) To 1 Step -1
            strTmp = strTmp & Codes(l)
        Next l
    End If
    GetAllCodes = strTmp
End Function

Function GetAllNextCodes() As String
    Dim strTmp As String
    Dim l As Long
    
    strTmp = ""
    If UBound(NextCodes) > 0 Then
        For l = 1 To UBound(NextCodes)
            strTmp = strTmp & NextCodes(l)
        Next l
    End If
    GetAllNextCodes = strTmp
End Function

Function GetAllCodesBeg() As String
    Dim strTmp As String
    Dim l As Long
    
    strTmp = ""
    If UBound(CodesBeg) > 0 Then
        For l = 1 To UBound(CodesBeg)
            strTmp = strTmp & CodesBeg(l)
        Next l
    End If
    GetAllCodesBeg = strTmp
End Function

Function GetAllNextCodesBeg() As String
    Dim strTmp As String
    Dim l As Long
    
    strTmp = ""
    If UBound(NextCodesBeg) > 0 Then
        For l = 1 To UBound(NextCodesBeg)
            strTmp = strTmp & NextCodesBeg(l)
        Next l
    End If
    GetAllNextCodesBeg = strTmp
End Function


Function PopCodeBeg() As String
    If UBound(CodesBeg) > 0 Then
        PopCodeBeg = CodesBeg(UBound(CodesBeg))
        ReDim Preserve CodesBeg(UBound(CodesBeg) - 1)
    End If
End Function

Function PopTmp() As String
    If UBound(CodesTmp) > 0 Then
        PopTmp = CodesTmp(UBound(CodesTmp))
        ReDim Preserve CodesTmp(UBound(CodesTmp) - 1)
    End If
End Function

Function PopTmpBeg() As String
    If UBound(CodesTmp) > 0 Then
        PopTmpBeg = CodesTmpBeg(UBound(CodesTmpBeg))
        ReDim Preserve CodesTmpBeg(UBound(CodesTmpBeg) - 1)
    End If
End Function

Function PopNext() As String
    If UBound(NextCodes) > 0 Then
        PopNext = NextCodes(UBound(NextCodes))
        ReDim Preserve NextCodes(UBound(NextCodes) - 1)
    End If
End Function

Function PopNextBeg() As String
    If UBound(NextCodesBeg) > 0 Then
        PopNextBeg = NextCodesBeg(UBound(NextCodesBeg))
        ReDim Preserve NextCodesBeg(UBound(NextCodesBeg) - 1)
    End If
End Function

Function ProcessWord(strWord As String)
    
    
    Dim l As Long
    
    
    Dim strTableAlign As String    'current table alignment for setting up tablestring
    Dim strTableWidth As String    'current table width for setting up tablestring
    
    If lSkipWords > 0 Then
        lSkipWords = lSkipWords - 1
        Exit Function
    End If
    If Left(strWord, 1) = "\" Or Left(strWord, 1) = "{" Or Left(strWord, 1) = "}" Then
        Select Case Left(strWord, 2)
        Case "}"
            For l = 1 To UBound(CodesBeg)
                CodesBeg(l) = ""
            Next l
            ClearFont
        Case "\b"    'bold
            If strWord = "\b" Then
                If Codes(UBound(Codes)) <> "</b>" Or (Codes(UBound(Codes)) = "</b>" And CodesBeg(UBound(Codes)) = "") Then
                    PushNext ("</b>")
                    PushNextBeg ("<b>")
                End If
            ElseIf strWord = "\bullet" Then
            ElseIf strWord = "\b0" Then    'bold off
                If InCodes("</b>") Then
                    Codes2NextTill ("</b>")
                ElseIf InNext("</b>") Then
                    RemoveFromNext ("</b>")
                End If
            End If
        Case "\c"
            If strWord = "\cf0" Then    'color font off
                If InCodes("</font>") Then
                    Codes2NextTill ("</font>")
                ElseIf InNext("</font>") Then
                    RemoveFromNext ("</font>")
                End If
            ElseIf Left(strWord, 3) = "\cf" And IsNumeric(Mid(strWord, 4)) Then  'color font
                'get color code
                l = Val(Mid(strWord, 4))
                If l <= UBound(strColorTable) And l > 0 Then
                    strFontColor = "#" & strColorTable(l)
                End If
                
                'insert color
                If strFontColor <> "#" Then
                    strFont = ParseFont(strFontColor, strFontSize)
                    If InNext("</font>") Then
                        ReplaceInNextBeg "</font>", strFont
                    ElseIf InCodes("</font>") Then
                        PushNext ("</font>")
                        PushNextBeg (strFont)
                        Codes2NextTill "</font>"
                    Else
                        PushNext ("</font>")
                        PushNextBeg (strFont)
                    End If
                End If
            End If
        Case "\f"
            If Left(strWord, 3) = "\fs" And IsNumeric(Mid(strWord, 4)) Then  'font size
                l = Val(Mid(strWord, 4))
                lFontSize = Int((l / 6) - 0)    'calc to convert RTF to HTML sizes
                If lFontSize > 8 Then lFontSize = 8
                If lFontSize < 1 Then lFontSize = 1
                strFontSize = Trim(lFontSize)
                'insert size
                If strFontSize <> "2" And strFontSize <> "" Then
                    strFont = ParseFont(strFontColor, strFontSize)
                    If InNext("</font>") Then
                        ReplaceInNextBeg "</font>", strFont
                    ElseIf InCodes("</font>") Then
                        PushNext ("</font>")
                        PushNextBeg (strFont)
                        Codes2NextTill "</font>"
                    Else
                        PushNext ("</font>")
                        PushNextBeg (strFont)
                    End If
                End If
            End If
        Case "\i"
            If strWord = "\i" Then 'italics
                If Codes(UBound(Codes)) <> "</i>" Or (Codes(UBound(Codes)) = "</i>" And CodesBeg(UBound(Codes)) = "") Then
                    PushNext ("</i>")
                    PushNextBeg ("<i>")
                End If
            ElseIf strWord = "\i0" Then 'italics off
                If InCodes("</i>") Then
                    Codes2NextTill ("</i>")
                ElseIf InNext("</i>") Then
                    RemoveFromNext ("</i>")
                End If
            End If
        Case "\l"
            'If strWord = "\listname" Then
            '    lSkipWords = 1
            'End If
        Case "\p"
            If strWord = "\par" Then
                strBeforeText2 = strBeforeText2 & strEOL & "<br>" & strCR
                gBOL = True
                'If Len(strBOL) > 0 Then
                '    PushNext ("</li>")
                '    PushNextBeg ("<li>")
                'End If
            ElseIf strWord = "\pard" Then
                For l = 1 To UBound(CodesBeg)
                    CodesBeg(l) = ""
                Next l
                ClearFont
            ElseIf strWord = "\plain" Then
                For l = 1 To UBound(CodesBeg)
                    CodesBeg(l) = ""
                Next l
                ClearFont
            ElseIf strWord = "\pnlvlblt" Then 'bulleted list
                'If Codes(UBound(Codes)) = "</u>" Then
                '    strTmp = PopCode
                '    strTmp = PopCodeBeg
                'End If
                'PushNext ("</ul>")
                'PushNextBeg ("<ul>")

                'strBOS = "<UL>"
                'strBOL = "<li>"
                'strEOL = "</li>"
                'strEOP = "</UL>"
            ElseIf strWord = "\pntxta" Then 'numbered list?
                lSkipWords = 1
            ElseIf strWord = "\pntxtb" Then 'numbered list?
                lSkipWords = 1
            End If
        Case "\q"
            If strWord = "\qc" Then    'centered
                strTableAlign = "center"
                strTableWidth = "100%"
                If InNext("</td></tr></table>") Then
                    '?
                Else
                    strTable = "<table width=" & strTableWidth & "><tr><td align=""" & strTableAlign & """>"
                End If
                If InNext("</td></tr></table>") Then
                    ReplaceInNextBeg "</td></tr></table>", strTable
                ElseIf InCodes("</td></tr></table>") Then
                    PushNext ("</td></tr></table>")
                    PushNextBeg (strTable)
                    Codes2NextTill "</td></tr></table>"
                Else
                    PushNext ("</td></tr></table>")
                    PushNextBeg (strTable)
                End If
            ElseIf strWord = "\qr" Then    'right justified
                strTableAlign = "right"
                strTableWidth = "100%"
                If InNext("</td></tr></table>") Then
                    '?
                Else
                    strTable = "<table width=" & strTableWidth & "><tr><td align=""" & strTableAlign & """>"
                End If
                If InNext("</td></tr></table>") Then
                    ReplaceInNextBeg "</td></tr></table>", strTable
                ElseIf InCodes("</td></tr></table>") Then
                    PushNext ("</td></tr></table>")
                    PushNextBeg (strTable)
                    Codes2NextTill "</td></tr></table>"
                Else
                    PushNext ("</td></tr></table>")
                    PushNextBeg (strTable)
                End If
            End If
        Case "\s"
            'If strWord = "\snext0" Then    'style
            '    lSkipWords = 1
            'End If
        Case "\u"
            If strWord = "\ul" Then    'underline
                If Codes(UBound(Codes)) <> "</u>" Or (Codes(UBound(Codes)) = "</u>" And CodesBeg(UBound(Codes)) = "") Then
                    PushNext ("</u>")
                    PushNextBeg ("<u>")
                End If
            ElseIf strWord = "\ulnone" Then    'stop underline
                If InCodes("</u>") Then
                    Codes2NextTill ("</u>")
                ElseIf InNext("</u>") Then
                    RemoveFromNext ("</u>")
                End If
            End If
        End Select
    Else
        If Len(strWord) > 0 Then
            If Trim(strWord) = "" Then
                If gBOL Then strWord = rtf2html_replace(strWord, " ", "&nbsp;")
                strCurPhrase = strCurPhrase & strBeforeText3 & strWord
            Else
                strBeforeText = strBeforeText & GetAllCodes
                Next2Codes
                strBeforeText3 = GetAllCodesBeg
                RemoveBlanks
                
                strCurPhrase = strCurPhrase & strBeforeText
                strBeforeText = ""
                strCurPhrase = strCurPhrase & strBeforeText2
                strBeforeText2 = ""
                strCurPhrase = strCurPhrase & strBeforeText3 & strWord
                strBeforeText3 = ""
                gBOL = False
            End If
        End If
    End If
    'MsgBox (strWord)
End Function


Function PushCode(strCode As String)
    ReDim Preserve Codes(UBound(Codes) + 1)
    Codes(UBound(Codes)) = strCode
End Function

Function PushTmp(strCode As String)
    ReDim Preserve CodesTmp(UBound(CodesTmp) + 1)
    CodesTmp(UBound(CodesTmp)) = strCode
End Function

Function PushTmpBeg(strCode As String)
    ReDim Preserve CodesTmpBeg(UBound(CodesTmpBeg) + 1)
    CodesTmpBeg(UBound(CodesTmpBeg)) = strCode
End Function

Function PushCodeBeg(strCode As String)
    ReDim Preserve CodesBeg(UBound(CodesBeg) + 1)
    CodesBeg(UBound(CodesBeg)) = strCode
End Function

Function PushNext(strCode As String)
    If Len(strCode) > 0 Then
        ReDim Preserve NextCodes(UBound(NextCodes) + 1)
        NextCodes(UBound(NextCodes)) = strCode
    End If
End Function

Function PushNextBeg(strCode As String)
    ReDim Preserve NextCodesBeg(UBound(NextCodesBeg) + 1)
    NextCodesBeg(UBound(NextCodesBeg)) = strCode
End Function

Function RemoveBlanks()
    Dim l As Long
    Dim lOffSet As Long
    
    l = 1
    lOffSet = 0
    While l <= UBound(CodesBeg) And l + lOffSet <= UBound(CodesBeg)
        If CodesBeg(l) = "" Then
            lOffSet = lOffSet + 1
        Else
            l = l + 1
        End If
        If l + lOffSet <= UBound(CodesBeg) Then
            Codes(l) = Codes(l + lOffSet)
            CodesBeg(l) = CodesBeg(l + lOffSet)
        End If
    Wend
    If lOffSet > 0 Then
        ReDim Preserve Codes(UBound(Codes) - lOffSet)
        ReDim Preserve CodesBeg(UBound(CodesBeg) - lOffSet)
    End If
End Function


Function RemoveFromNext(strRem As String)
    Dim l As Long
    Dim m As Long
    
    l = 1
    While l < UBound(NextCodes)
        If NextCodes(l) = strRem Then
            For m = l To UBound(NextCodes) - 1
                NextCodes(m) = NextCodes(m + 1)
                NextCodesBeg(m) = NextCodesBeg(m + 1)
            Next m
            l = m
        Else
            l = l + 1
        End If
    Wend
    ReDim Preserve NextCodes(UBound(NextCodes) - 1)
    ReDim Preserve NextCodesBeg(UBound(NextCodesBeg) - 1)
End Function

Function rtf2html_replace(ByVal strIn As String, ByVal strRepl As String, ByVal strWith As String) As String
    'replace all instances of strRepl in strIn with strWith
    Dim I As Integer
    
    If ((Len(strRepl) = 0) Or (Len(strIn) = 0)) Then
        rtf2html_replace = strIn
        Exit Function
    End If
    I = InStr(strIn, strRepl)
    While I <> 0
        strIn = Left(strIn, I - 1) & strWith & Mid(strIn, I + Len(strRepl))
        I = InStr(I + Len(strWith), strIn, strRepl)
    Wend
    rtf2html_replace = strIn
End Function

Function ReplaceInNextBeg(strCode As String, strWith As String)
    Dim l As Long
    
    l = 1
    While l <= UBound(NextCodes) And NextCodes(l) <> strCode
        l = l + 1
    Wend
    If NextCodes(l) = strCode Then
        NextCodesBeg(l) = strWith
    End If
End Function

Function rtf2html(strRTF As String, Optional strOptions As String) As String
    'Version 3.01b04
    'Copyright Brady Hegberg 2000
    '  I'm not licensing this software but I'd appreciate it if
    '  you'd to consider it to be under an lgpl sort of license
    
    'More information can be found at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html
    
    'Converts Rich Text encoded text to HTML format
    'if you find some text that this function doesn't
    'convert properly please email the text to
    'bradyh@bitstream.net
    
    'Options:
    '+H              add an HTML header and footer
    '+G              add a generator Metatag
    '+T="MyTitle"    add a title (only works if +H is used)
    '+CR             add a carraige return after all <br>s
    '+I              keep html codes intact
    
    Dim strHTML As String
    Dim strRTFTmp As String
    Dim lBOS As Long                 'beginning of section                'end of section
   
    
    
    
    
    
    
    
    
    Dim gHTML As Boolean             'true if html codes should be left intact
    Dim strSecTmp As String          'temporary section buffer
    Dim strWordTmp As String         'temporary word buffer
    Dim strEndText As String         'ending text
    
    Dim strHtmlBody As String
  
    
    ClearCodes
    strHTML = ""
    gPlain = False
    gBOL = True
    
    'setup +CR option
    If InStr(strOptions, "+CR") <> 0 Then strCR = vbCrLf Else strCR = ""
    'setup +HTML option
    If InStr(strOptions, "+I") <> 0 Then gHTML = True Else gHTML = False

    strRTFTmp = TrimAll(strRTF)

    If Left(strRTFTmp, 1) = "{" And Right(strRTFTmp, 1) = "}" Then strRTFTmp = Mid(strRTFTmp, 2, Len(strRTFTmp) - 2)
    
    'setup color table
    lBOS = InStr(strRTFTmp, "\colortbl")
    If lBOS > 0 Then
        strSecTmp = NabSection(strRTFTmp, lBOS)
        GetColorTable strSecTmp, strColorTable()
    End If
    
    'setup font table
    lBOS = InStr(strRTFTmp, "\fonttbl")
    If lBOS > 0 Then
        strSecTmp = NabSection(strRTFTmp, lBOS)
        GetFontTable strSecTmp, strFontTable()
    End If
    
    'setup stylesheets
    lBOS = InStr(strRTFTmp, "\stylesheet")
    If lBOS > 0 Then
        strSecTmp = NabSection(strRTFTmp, lBOS)
        'ignore stylesheets for now
    End If
    
    'setup info
    lBOS = InStr(strRTFTmp, "\info")
    If lBOS > 0 Then
        strSecTmp = NabSection(strRTFTmp, lBOS)
        'ignore info for now
    End If
    
    'list table
    lBOS = InStr(strRTFTmp, "\listtable")
    If lBOS > 0 Then
        strSecTmp = NabSection(strRTFTmp, lBOS)
        'ignore info for now
    End If
    
    'list override table
    lBOS = InStr(strRTFTmp, "\listoverridetable")
    If lBOS > 0 Then
        strSecTmp = NabSection(strRTFTmp, lBOS)
        'ignore info for now
    End If

    While Len(strRTFTmp) > 0
        strSecTmp = NabNextLine(strRTFTmp)
        While Len(strSecTmp) > 0
            strWordTmp = NabNextWord(strSecTmp)
            If Len(strWordTmp) > 0 Then ProcessWord strWordTmp
        Wend
    Wend
    
    'get any remaining codes in stack
    Next2Codes
    strEndText = strEndText & GetAllCodes
    strBeforeText2 = rtf2html_replace(strBeforeText2, "<br>", "")
    strBeforeText2 = rtf2html_replace(strBeforeText2, vbCrLf, "")
    strCurPhrase = strCurPhrase & strBeforeText & strBeforeText2 & strEndText
    strBeforeText = ""
    strBeforeText2 = ""
    strBeforeText3 = ""
    strHTML = strHTML & strCurPhrase
    strCurPhrase = ""
    
    Dim strTitel As String
    
    If InStr(strOptions, "+T=") > 0 Then
        strTitel = GetTitel(0, "+T=", strOptions)
    End If
    
    If InStr(strOptions, "+H") > 0 Then
        strHtmlBody = strHtmlBody + "<HTML>" + strCR
        strHtmlBody = strHtmlBody + "<HEAD>" + strCR
        strHtmlBody = strHtmlBody + "<TITLE>" + strTitel + "</TITLE>" + strCR
        strHtmlBody = strHtmlBody + "</HEAD>" + strCR
        strHtmlBody = strHtmlBody + "<BODY bgcolor=" + "white" + " text=" + "black" + ">" + strCR
    
        rtf2html = strHtmlBody + strHTML + "</BODY>" + strCR + "</HTML>"
    Else
        rtf2html = strHTML
    End If
End Function

Public Function GetTitel(intPosition As Long, SearchStr As String, ByRef strarray As String) As String

  Dim strTemp As String
  Dim strValue As String
  Dim Counter As Integer
  Dim StartPosi As Integer
  

    On Error GoTo error

    
    StartPosi = InStr(LCase$(strarray), SearchStr) + Len(SearchStr)

    Do
        strValue = strValue + strTemp
        strTemp = Mid$(strarray, StartPosi + Counter, 1)
        Counter = Counter + 1
    Loop Until strTemp = vbCrLf Or Counter = Len(strarray)

    'Remove the ""
    If Left$(strValue, 1) = Chr$(34) Then strValue = Right$(strValue, Len(strValue) - 1)
    If Right$(strValue, 1) = Chr$(34) Then strValue = Left$(strValue, Len(strValue) - 1)

    GetTitel = Replace(strValue, " ", "")

Exit Function

error:
    GetTitel = ""

End Function

Function ShowCodes()
    Dim strTmp As String
    Dim l As Long
    
    strTmp = "Codes: "
    For l = 1 To UBound(Codes)
        strTmp = strTmp & Codes(l) & ", "
    Next l
    strTmp = strTmp & vbCrLf & "BegCodes: "
    For l = 1 To UBound(CodesBeg)
        strTmp = strTmp & CodesBeg(l) & ", "
    Next l
    strTmp = strTmp & vbCrLf & "NextCodes: "
    For l = 1 To UBound(NextCodes)
        strTmp = strTmp & NextCodes(l) & ", "
    Next l
    strTmp = strTmp & vbCrLf & "NextBegCodes: "
    For l = 1 To UBound(NextCodesBeg)
        strTmp = strTmp & NextCodesBeg(l) & ", "
    Next l
    MsgBox (strTmp)
End Function

Function TrimAll(ByVal strTmp As String) As String
    Dim l As Long
    
    strTmp = Trim(strTmp)
    l = Len(strTmp) + 1
    While l <> Len(strTmp)
        l = Len(strTmp)
        If Right(strTmp, 1) = vbCrLf Then strTmp = Left(strTmp, Len(strTmp) - 1)
        If Left(strTmp, 1) = vbCrLf Then strTmp = Right(strTmp, Len(strTmp) - 1)
        If Right(strTmp, 1) = vbCr Then strTmp = Left(strTmp, Len(strTmp) - 1)
        If Left(strTmp, 1) = vbCr Then strTmp = Right(strTmp, Len(strTmp) - 1)
        If Right(strTmp, 1) = vbLf Then strTmp = Left(strTmp, Len(strTmp) - 1)
        If Left(strTmp, 1) = vbLf Then strTmp = Right(strTmp, Len(strTmp) - 1)
    Wend
    TrimAll = strTmp
End Function

Function HTMLCode(strRTFCode As String) As String
    'given rtf code return html code
    Select Case strRTFCode
    Case "00"
        HTMLCode = "&nbsp;"
    Case "a9"
        HTMLCode = "&copy;"
    Case "b4"
        HTMLCode = "&acute;"
    Case "ab"
        HTMLCode = "&laquo;"
    Case "bb"
        HTMLCode = "&raquo;"
    Case "a1"
        HTMLCode = "&iexcl;"
    Case "bf"
        HTMLCode = "&iquest;"
    Case "c0"
        HTMLCode = "&Agrave;"
    Case "e0"
        HTMLCode = "&agrave;"
    Case "c1"
        HTMLCode = "&Aacute;"
    Case "e1"
        HTMLCode = "&aacute;"
    Case "c2"
        HTMLCode = "&Acirc;"
    Case "e2"
        HTMLCode = "&acirc;"
    Case "c3"
        HTMLCode = "&Atilde;"
    Case "e3"
        HTMLCode = "&atilde;"
    Case "c4"
        HTMLCode = "&Auml;"
    Case "e4"
        HTMLCode = "<FONT SIZE=""-1""><SUP>TM</SUP></FONT>"
    Case "c5"
        HTMLCode = "&Aring;"
    Case "e5"
        HTMLCode = "&aring;"
    Case "c6"
        HTMLCode = "&AElig;"
    Case "e6"
        HTMLCode = "&aelig;"
    Case "c7"
        HTMLCode = "&Ccedil;"
    Case "e7"
        HTMLCode = "&ccedil;"
    Case "d0"
        HTMLCode = "&ETH;"
    Case "f0"
        HTMLCode = "&eth;"
    Case "c8"
        HTMLCode = "&Egrave;"
    Case "e8"
        HTMLCode = "&egrave;"
    Case "c9"
        HTMLCode = "&Eacute;"
    Case "e9"
        HTMLCode = "&eacute;"
    Case "ca"
        HTMLCode = "&Ecirc;"
    Case "ea"
        HTMLCode = "&ecirc;"
    Case "cb"
        HTMLCode = "&Euml;"
    Case "eb"
        HTMLCode = "&euml;"
    Case "cc"
        HTMLCode = "&Igrave;"
    Case "ec"
        HTMLCode = "&igrave;"
    Case "cd"
        HTMLCode = "&Iacute;"
    Case "ed"
        HTMLCode = "&iacute;"
    Case "ce"
        HTMLCode = "&Icirc;"
    Case "ee"
        HTMLCode = "&icirc;"
    Case "cf"
        HTMLCode = "&Iuml;"
    Case "ef"
        HTMLCode = "&iuml;"
    Case "d1"
        HTMLCode = "&Ntilde;"
    Case "f1"
        HTMLCode = "&ntilde;"
    Case "d2"
        HTMLCode = "&Ograve;"
    Case "f2"
        HTMLCode = "&ograve;"
    Case "d3"
        HTMLCode = "&Oacute;"
    Case "f3"
        HTMLCode = "&oacute;"
    Case "d4"
        HTMLCode = "&Ocirc;"
    Case "f4"
        HTMLCode = "&ocirc;"
    Case "d5"
        HTMLCode = "&Otilde;"
    Case "f5"
        HTMLCode = "&otilde;"
    Case "d6"
        HTMLCode = "&Ouml;"
    Case "f6"
        HTMLCode = "&ouml;"
    Case "d8"
        HTMLCode = "&Oslash;"
    Case "f8"
        HTMLCode = "&oslash;"
    Case "d9"
        HTMLCode = "&Ugrave;"
    Case "f9"
        HTMLCode = "&ugrave;"
    Case "da"
        HTMLCode = "&Uacute;"
    Case "fa"
        HTMLCode = "&uacute;"
    Case "db"
        HTMLCode = "&Ucirc;"
    Case "fb"
        HTMLCode = "&ucirc;"
    Case "dc"
        HTMLCode = "&Uuml;"
    Case "fc"
        HTMLCode = "&uuml;"
    Case "dd"
        HTMLCode = "&Yacute;"
    Case "fd"
        HTMLCode = "&yacute;"
    Case "ff"
        HTMLCode = "&yuml;"
    Case "de"
        HTMLCode = "&THORN;"
    Case "fe"
        HTMLCode = "&thorn;"
    Case "df"
        HTMLCode = "&szlig;"
    Case "a7"
        HTMLCode = "&sect;"
    Case "b6"
        HTMLCode = "&para;"
    Case "b5"
        HTMLCode = "&micro;"
    Case "a6"
        HTMLCode = "&brvbar;"
    Case "b1"
        HTMLCode = "&plusmn;"
    Case "b7"
        HTMLCode = "&middot;"
    Case "a8"
        HTMLCode = "&uml;"
    Case "b8"
        HTMLCode = "&cedil;"
    Case "aa"
        HTMLCode = "&ordf;"
    Case "ba"
        HTMLCode = "&ordm;"
    Case "ac"
        HTMLCode = "&not;"
    Case "ad"
        HTMLCode = "&shy;"
    Case "af"
        HTMLCode = "&macr;"
    Case "b0"
        HTMLCode = "&deg;"
    Case "b9"
        HTMLCode = "&sup1;"
    Case "b2"
        HTMLCode = "&sup2;"
    Case "b3"
        HTMLCode = "&sup3;"
    Case "bc"
        HTMLCode = "&frac14;"
    Case "bd"
        HTMLCode = "&frac12;"
    Case "be"
        HTMLCode = "&frac34;"
    Case "d7"
        HTMLCode = "&times;"
    Case "f7"
        HTMLCode = "&divide;"
    Case "a2"
        HTMLCode = "&cent;"
    Case "a3"
        HTMLCode = "&pound;"
    Case "a4"
        HTMLCode = "&curren;"
    Case "a5"
        HTMLCode = "&yen;"
    Case "85"
        HTMLCode = "..."
    End Select
End Function

Function TrimifCmd(ByVal strTmp As String) As String
    Dim l As Long
    
    l = 1
    While Mid(strTmp, l, 1) = " "
        l = l + 1
    Wend
    If Mid(strTmp, l, 1) = "\" Or Mid(strTmp, l, 1) = "{" Then
        strTmp = Trim(strTmp)
    Else
        If Left(strTmp, 1) = " " Then strTmp = Mid(strTmp, 2)
        strTmp = RTrim(strTmp)
    End If
    TrimifCmd = strTmp
End Function


