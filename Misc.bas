Attribute VB_Name = "Misc"
'--------------------------------------------------------------------------
'
' Author:       90 % Sebastian Fahrenkrog (contact@wirdesignen.de)
' DateCreated:  16.06.2002
' Description:  misc stuff and declarations for the vbMime class
'
' ModuleType:   bas
'
'--------------------------------------------------------------------------

Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Mails() As Mail

Public Type Attachments
    Name    As String
    Data()  As String
End Type

Public Type Mail
    Header           As String
    from             As String
    To               As String
    Date             As String
    Subject          As String
    Message          As String 'Plain Text Message
    HTMLMessage      As String 'HTML Message Part
    Size             As Long
    AttachedFiles    As Integer
    Attachments() As Attachments
End Type

Public strlines()          As String
Public strLine()           As String
Public tmpAttachmntStr     As String
Public AttachmentCounter   As Integer

'Declarations for very fast String Array Routines
Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (dest As Any, Source As Any, _
        ByVal numBytes As Long)
Declare Sub ZeroMemory Lib "kernel32" Alias _
        "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

'Base64 Class
Public pbBuffer1() As Byte
Public pbBuffer2() As Byte
Public ptSpan()    As String

'Class for the multi language support
Global cLanguage As New clsLanguagePack

'Prevent the showing of the right click Internet Explorer window
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
Public Const WM_RBUTTONUP = &H205
Public Const WH_MOUSE = 7

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type
    
Public gLngMouseHook As Long
    
Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, mhs As MOUSEHOOKSTRUCT) As Long
Dim strBuffer As String
Dim strClassName As String
Dim lngResult As Long

If (nCode >= 0 And wParam = WM_RBUTTONUP) Then

        'Preinitialize string
        strBuffer = Space(255)
        
       ' lngBufferLen = Len(strBuffer)
        
        'This is the string that holds the class name that we are looking for
        strClassName = "Internet Explorer_Server"
        
        'Debug.Print strClassName
        
        'Get the classname for the Window that has been clicked, making sure something is returned
        'If the function returns 0, it has failed
        lngResult = GetClassName(mhs.hwnd, strBuffer, Len(strBuffer))
                
        'Debug.Print Left$(strBuffer, lngResult)
                
        If lngResult > 0 Then

            'Check to see if the class of the window we clicked on is the same as above
            If Left$(strBuffer, lngResult) = strClassName Then
                
                'Value is the same. Squash the command
                MouseHookProc = 1
                
                Exit Function
                
            End If
            
        End If

    End If

MouseHookProc = CallNextHookEx(gLngMouseHook, nCode, wParam, mhs)
End Function

Public Function CheckExistence(Pclist As ComboBox, Data As String) As Boolean
Dim Counter As Integer

For Counter = 0 To Pclist.ListCount
    If Pclist.List(Counter) = Data Then
        CheckExistence = True
        Exit Function
    End If
Next Counter

End Function

Function SaveIni(KeySection As String, strKey As String, KeyValue As String)

  Dim lngResult As Long
  Dim strFilename

    strFilename = App.Path & "\Pop3Popper.ini" 'Declare your ini file !
    lngResult = WritePrivateProfileString(KeySection, strKey, KeyValue, strFilename)

    SaveIni = lngResult

End Function

Function LoadIni(KeySection As String, strKey As String)
    
    Dim lngResult As Long
    Dim strFilename As String
    Dim strResult As String * 100
    Dim KeyValue As String
    
    strFilename = App.Path & "\Pop3Popper.ini" 'Declare your ini file !
    
    lngResult = GetPrivateProfileString(KeySection, _
                strKey, "", strResult, Len(strResult), _
                strFilename)
    
    If lngResult = 0 Then
        'An error has occurred
        LoadIni = ""
    Else
        KeyValue = Trim(strResult)
        KeyValue = Replace(KeyValue, Chr(0), "")
        LoadIni = KeyValue
    End If
    
End Function

'Glue several lines together that belong together
Public Function UnfoldArray(fromLine As Long, toLine As Long, ByRef FoldedArray() As String) As String()
Dim Counter As Integer, UCounter As Integer
Dim strHeader As String
Dim TempArray() As String

On Error GoTo error


'Extract only the Mime Headers
ReDim TempArray(toLine - fromLine)

For Counter = fromLine To toLine
    TempArray(UCounter) = FoldedArray(Counter)
    UCounter = UCounter + 1
Next

strHeader = Join(TempArray, vbCrLf)

'Hmm I try to unfold the Mail Header...
strHeader = Replace(strHeader, vbCrLf + Chr$(9), " ")
strHeader = Replace(strHeader, vbCrLf + Chr$(11), " ")
strHeader = Replace(strHeader, vbCrLf + Chr$(32), " ")
strHeader = Replace(strHeader, vbCrLf + Chr$(255), " ")

UnfoldArray = Split(strHeader, vbCrLf)

error:

End Function

'Returns the Line that contains a String (reversed for speed reasons)
Public Function RevfindEmptyLine(ByRef strLine() As String) As Long
Dim Counter As Long
Dim TmpLngt As Long
Dim TmpString As String

On Error GoTo error

TmpLngt = UBound(strLine)
Counter = TmpLngt


    Do
        Counter = Counter - 1
            
        TmpString = strLine(Counter + 1)
        
       
        
            If TmpString = "" Then
                RevfindEmptyLine = Counter + 1
                Exit Function
            
        End If
    
            
    Loop Until Counter = 0
    
error:
    RevfindEmptyLine = -1
End Function

'Finds a line that only contain one Crlf
Public Function findEmptyLine(intPosition As Long, ByRef strlines() As String) As Long

  Dim Counter As Long
  Dim TmpLngt As Long
  Dim strTemp As String

    On Error GoTo error

    If intPosition < 0 Then
        findEmptyLine = -1
        Exit Function
    End If

    TmpLngt = UBound(strlines)

    Do
        Counter = Counter + 1
        strTemp = strlines(intPosition + Counter - 1)
    Loop Until Counter = TmpLngt Or Len(strTemp) = 0

    If strlines(intPosition + Counter - 1) = "" Then
        findEmptyLine = intPosition + Counter - 1
      Else
error:
        findEmptyLine = -1
    End If

End Function

'Returns the Line of an array that contains a String
Public Function findLine(intPosition As Long, SearchStr As String, strlines() As String, Optional IgnoreInstrWord As Boolean) As Long

  Dim Counter As Long
  Dim TmpLngt As Long
  Dim TmpLngt2 As Long

    On Error GoTo error

    TmpLngt = UBound(strlines)
    Counter = Counter + intPosition

    If Counter >= TmpLngt Then
        GoTo error
    End If

    Do
        Counter = Counter + 1

        Select Case IgnoreInstrWord
            Case False
                TmpLngt2 = InStrWord(strlines(Counter - 1), SearchStr)
            Case True
                TmpLngt2 = InStr(strlines(Counter - 1), SearchStr)
        End Select
        

        If TmpLngt2 > 0 Then
            findLine = Counter - 1

            Exit Function
        End If

    Loop Until Counter = TmpLngt

error:
    findLine = -1

End Function

'Get the Value from an E-Mail Header
Public Function GetInfo(intPosition As Long, SearchStr As String, ByRef strlines() As String) As String

  Dim strTemp As String
  Dim strValue As String
  Dim Counter As Integer
  Dim StartPosi As Integer
  Dim Counter2 As Integer
  Dim strarray As String

    On Error GoTo error

    strarray = strlines(intPosition)
    StartPosi = InStr(LCase$(strarray), SearchStr) + Len(SearchStr)

    Do
        strValue = strValue + strTemp
        strTemp = Mid$(strarray, StartPosi + Counter, 1)
        Counter = Counter + 1
        Counter2 = Len(strarray)
    Loop Until strTemp = vbCrLf Or Counter = Counter2

    'Remove the ""
    If Left$(strValue, 1) = Chr$(34) Then strValue = Right$(strValue, Len(strValue) - 1)
    If Right$(strValue, 1) = Chr$(34) Then strValue = Left$(strValue, Len(strValue) - 1)

    GetInfo = Replace(strValue, " ", "")

Exit Function

error:
    GetInfo = ""

End Function

'Returns the Line that contains a String (reversed for speed reasons)*
Public Function RevfindLine(SearchStr As String, ByRef strlines() As String) As Long

  Dim Counter As Long
  Dim TmpLngt As Long
  Dim TmpString As String

    On Error GoTo error

    TmpLngt = UBound(strlines)
    Counter = TmpLngt

    Do
        Counter = Counter - 1

        TmpString = strlines(Counter + 1)

        If InStr(TmpString, SearchStr) > 0 Then
            RevfindLine = Counter + 1
            Exit Function
        End If

    Loop Until Counter = 0

error:
    RevfindLine = -1

End Function

'Checks if a string contains a special seperated word
Public Function InStrWord( _
                          ByRef Text As String, _
                          ByRef Word As String _
                          ) As Long

  'Deklarationen:

  Dim WordLen As Long
  Dim TextEnd As Long
  Dim OK As Boolean

    WordLen = Len(Word)
    If WordLen = 0 Then
        Exit Function
    End If

    TextEnd = Len(Text) - WordLen + 1

    InStrWord = InStr(1, Text, Word, vbTextCompare)
    Do While InStrWord

        If InStrWord = 1 Then
            OK = True
          Else
            OK = IsWordSep(Mid$(Text, InStrWord - 1, 1))
        End If

        'Ggf. Zeichen hinter dem Wort checken:
        If OK And (InStrWord < TextEnd) Then
            OK = IsWordSep(Mid$(Text, InStrWord + WordLen, 1))
        End If

        'Treffer zurückgeben oder weitersuchen:
        If OK Then
            Exit Do
        End If

        InStrWord = InStr(InStrWord + WordLen, Text, Word, vbTextCompare)

    Loop

End Function

'Returns true if a char is a known seperator
Public Function IsWordSep(ByVal Char As String) As Boolean

    If Char = " " Or Char = vbCr Or Char = vbLf Or Char = vbTab Or Char = Chr$(34) Or Char = vbCrLf Or Char = "-" Then
        IsWordSep = True
    End If

End Function



'**************************************************************************************
'Replace function
'
'Author: unknown
'
'Desc:
'
'this functions are a lot faster than the original functions and usefull
'for VB5 User
''**************************************************************************************

Public Function Replace(ByRef Text As String, _
                        ByRef sOld As String, ByRef sNew As String, _
                        Optional ByVal Start As Long = 1, _
                        Optional ByVal Count As Long = 2147483647, _
                        Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
                        ) As String

    If LenB(sOld) Then

        If Compare = vbBinaryCompare Then
            ReplaceBin Replace, Text, Text, _
                       sOld, sNew, Start, Count
          Else
            ReplaceBin Replace, Text, LCase$(Text), _
                       LCase$(sOld), sNew, Start, Count
        End If

      Else 'Suchstring ist leer:
        Replace = Text
    End If

End Function

Private Static Sub ReplaceBin(ByRef Result As String, _
                ByRef Text As String, ByRef Search As String, _
                ByRef sOld As String, ByRef sNew As String, _
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
        Start = InStrB(Search, sOld)
      Else
        Start = InStrB(Start + Start - 1, Search, sOld)
    End If
    If Start Then

        OldLen = LenB(sOld)
        NewLen = LenB(sNew)
        Select Case NewLen
          Case OldLen 'einfaches Überschreiben:

            Result = Text
            For Count = 1 To Count
                MidB$(Result, Start) = sNew
                Start = InStrB(Start + OldLen, Search, sOld)
                If Start = 0 Then
                    Exit Sub
                End If
            Next Count
            Exit Sub

          Case Is < OldLen 'Ergebnis wird kürzer:

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

                'Einzufügenden Text beachten:
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
                    Start = InStrB(ReadPos, Search, sOld)
                    If Start = 0 Then
                        Exit For
                    End If
                Next Count

              Else

                'Einzufügenden Text ignorieren (weil leer):
                For Count = 1 To Count
                    CopyLen = Start - ReadPos
                    If CopyLen Then
                        MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
                        WritePos = WritePos + CopyLen
                    End If
                    ReadPos = Start + OldLen
                    Start = InStrB(ReadPos, Search, sOld)
                    If Start = 0 Then
                        Exit For
                    End If
                Next Count

            End If

            'Ergebnis zusammenbauen:
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
              Else
                MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                Result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
            End If
            Exit Sub

          Case Else 'Ergebnis wird länger:

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

                    'Ggf. Buffer vergrößern:
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

                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If

                    'String "patchen":
                    MidB$(Buffer, WritePos) = sNew
                End If
                WritePos = BufferPosNext
                ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
                If Start = 0 Then Exit For
            Next Count

            'Ergebnis zusammenbauen:
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
              Else
                BufferPosNext = WritePos + TextLen - ReadPos
                If BufferPosNext < BufferLen Then
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                    Result = LeftB$(Buffer, BufferPosNext)
                  Else
                    Result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
                End If
            End If
            Exit Sub

        End Select

      Else
        Result = Text
    End If

End Sub

Public Sub MoveStringArray(Source() As String, dest() As String, firstEl As Long, lastEL As Long)

  Dim numBytes As Long


On Error GoTo error

    numBytes = (lastEL - firstEl + 1) * 4
    ' start with a fresh new array
    '(it clears all its descriptors)
    ReDim dest(0 To lastEL - firstEl) As String
    ' copy all the descriptors from source() to dest()
    CopyMemory ByVal VarPtr(dest(0)), _
               ByVal VarPtr(Source(firstEl)), numBytes
    ' manually clear all the descriptors in source()
    ZeroMemory ByVal VarPtr(Source(firstEl)), numBytes

error:
End Sub



':) Ulli's VB Code Formatter V2.12.7 (19.06.2002 23:13:06) 48 + 401 = 449 Lines
