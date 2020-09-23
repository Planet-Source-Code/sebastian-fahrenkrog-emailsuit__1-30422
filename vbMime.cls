VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
'
' Author:       Sebastian Fahrenkrog (contact@wirdesignen.de)
' DateCreated:  16.06.2002
' Description:  Mime class contains all functions to decode E-Mails with
'               Attachments
' ModuleType:   Class
'
'--------------------------------------------------------------------------

Option Explicit
Option Base 0

'Base64
Private Const CHAR_EQUAL As Byte = 61
Private Const CHAR_CR As Byte = 13
Private Const CHAR_LF As Byte = 10

Private m_ReverseIndex1(0 To 255) As Byte
Private m_ReverseIndex2(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex3(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex4(0 To 255) As Byte

'Mime
Private m_strMessageText     As String
Private m_strMessageBody     As String
Private m_strHeaders         As String

'Pop3 Class
'Dim intMailSelected As Integer

Private Enum POP3States
    POP3_Connect
    POP3_USER
    POP3_PASS
    POP3_STAT
    Pop3_retr
    Pop3_dele
    POP3_QUIT
End Enum

Private m_State       As POP3States
Private m_strPop3Host As String
Private m_strUsername As String
Private m_strPassword As String
Private bolDelMail As Boolean
Private pbExitImmediately As Boolean
Private bRaiseTimeOutError As Boolean
Private pbConnected As Boolean

Private intMessages          As Integer
Private intCurrentMessage    As Integer
Private strBuffer            As String
Private DataPointer&
Private Const BlockSize = 2048
' Class Events
Private WithEvents Pop3sck As CSocket
Attribute Pop3sck.VB_VarHelpID = -1
Public Event ReceivedSuccesful()
Public Event MimeFailed(Explanation As String)
Public Event Pop3Status(Status As String)
Public Event Progress(PercentComplete As Long)
'For WaitUntilTrue()
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim vbBase64 As New base64



'==========
' Class_Initialize;
' initializes codec tables.
'==========
Private Sub Class_Initialize()

  'Initiate Winsock

    Set Pop3sck = New CSocket
    

    'initialize the base64 table
    Dim I As Long

    'Setup the encodeing and decoding lookup arrays.
    'Essentially we speed up the routine by pre-shifting
    'the data so it only needs combined with And and Or.
    
    m_ReverseIndex4(65) = 0 'Asc("A")
    m_ReverseIndex4(66) = 1 'Asc("B")
    m_ReverseIndex4(67) = 2 'Asc("C")
    m_ReverseIndex4(68) = 3 'Asc("D")
    m_ReverseIndex4(69) = 4 'Asc("E")
    m_ReverseIndex4(70) = 5 'Asc("F")
    m_ReverseIndex4(71) = 6 'Asc("G")
    m_ReverseIndex4(72) = 7 'Asc("H")
    m_ReverseIndex4(73) = 8 'Asc("I")
    m_ReverseIndex4(74) = 9 'Asc("J")
    m_ReverseIndex4(75) = 10 'Asc("K")
    m_ReverseIndex4(76) = 11 'Asc("L")
    m_ReverseIndex4(77) = 12 'Asc("M")
    m_ReverseIndex4(78) = 13 'Asc("N")
    m_ReverseIndex4(79) = 14 'Asc("O")
    m_ReverseIndex4(80) = 15 'Asc("P")
    m_ReverseIndex4(81) = 16 'Asc("Q")
    m_ReverseIndex4(82) = 17 'Asc("R")
    m_ReverseIndex4(83) = 18 'Asc("S")
    m_ReverseIndex4(84) = 19 'Asc("T")
    m_ReverseIndex4(85) = 20 'Asc("U")
    m_ReverseIndex4(86) = 21 'Asc("V")
    m_ReverseIndex4(87) = 22 'Asc("W")
    m_ReverseIndex4(88) = 23 'Asc("X")
    m_ReverseIndex4(89) = 24 'Asc("Y")
    m_ReverseIndex4(90) = 25 'Asc("Z")
    m_ReverseIndex4(97) = 26 'Asc("a")
    m_ReverseIndex4(98) = 27 'Asc("b")
    m_ReverseIndex4(99) = 28 'Asc("c")
    m_ReverseIndex4(100) = 29 'Asc("d")
    m_ReverseIndex4(101) = 30 'Asc("e")
    m_ReverseIndex4(102) = 31 'Asc("f")
    m_ReverseIndex4(103) = 32 'Asc("g")
    m_ReverseIndex4(104) = 33 'Asc("h")
    m_ReverseIndex4(105) = 34 'Asc("i")
    m_ReverseIndex4(106) = 35 'Asc("j")
    m_ReverseIndex4(107) = 36 'Asc("k")
    m_ReverseIndex4(108) = 37 'Asc("l")
    m_ReverseIndex4(109) = 38 'Asc("m")
    m_ReverseIndex4(110) = 39 'Asc("n")
    m_ReverseIndex4(111) = 40 'Asc("o")
    m_ReverseIndex4(112) = 41 'Asc("p")
    m_ReverseIndex4(113) = 42 'Asc("q")
    m_ReverseIndex4(114) = 43 'Asc("r")
    m_ReverseIndex4(115) = 44 'Asc("s")
    m_ReverseIndex4(116) = 45 'Asc("t")
    m_ReverseIndex4(117) = 46 'Asc("u")
    m_ReverseIndex4(118) = 47 'Asc("v")
    m_ReverseIndex4(119) = 48 'Asc("w")
    m_ReverseIndex4(120) = 49 'Asc("x")
    m_ReverseIndex4(121) = 50 'Asc("y")
    m_ReverseIndex4(122) = 51 'Asc("z")
    m_ReverseIndex4(48) = 52 'Asc("0")
    m_ReverseIndex4(49) = 53 'Asc("1")
    m_ReverseIndex4(50) = 54 'Asc("2")
    m_ReverseIndex4(51) = 55 'Asc("3")
    m_ReverseIndex4(52) = 56 'Asc("4")
    m_ReverseIndex4(53) = 57 'Asc("5")
    m_ReverseIndex4(54) = 58 'Asc("6")
    m_ReverseIndex4(55) = 59 'Asc("7")
    m_ReverseIndex4(56) = 60 'Asc("8")
    m_ReverseIndex4(57) = 61 'Asc("9")
    m_ReverseIndex4(43) = 62 'Asc("+")
    m_ReverseIndex4(47) = 63 'Asc("/")

    'Calculate the other arrays.
    For I = 0 To 255
        If m_ReverseIndex4(I) <> 0 Then
            m_ReverseIndex1(I) = m_ReverseIndex4(I) * 4

            m_ReverseIndex2(I, 0) = m_ReverseIndex4(I) \ 16
            m_ReverseIndex2(I, 1) = (m_ReverseIndex4(I) And &HF) * 16

            m_ReverseIndex3(I, 0) = m_ReverseIndex4(I) \ 4
            m_ReverseIndex3(I, 1) = (m_ReverseIndex4(I) And &H3) * 64
        End If
    Next I
    

End Sub

Private Sub Class_Terminate()

  ' make sure sckMail is closed

    If Pop3sck.State <> sckClosed Then
        Pop3sck.CloseSocket
    End If

    ' release memory
    Set Pop3sck = Nothing

End Sub

Public Sub GetMail(strUsername As String, strPassword As String, strHost As String, Optional intPort As Integer)

    m_strPop3Host = strHost
    m_strUsername = strUsername
    m_strPassword = strPassword

    'Change current state of session
    m_State = POP3_Connect
    '
    'Reset current state of socket
    Pop3sck.CloseSocket
    '
    'Reset local port value to prevent "Address in use" error
    Pop3sck.LocalPort = 0
    '
    'POP3 server software is listening for client connection
    'requests on 110 port, therefore we need connect to host
    'on 110 port
    If intPort = 0 Then
        intPort = 110
    End If
    RaiseEvent Pop3Status("Connecting to Pop3 Server...")
    
    Pop3sck.Connect m_strPop3Host, intPort
    
    Call WaitUntilTrue(pbConnected, 30, True)
End Sub

Private Sub Pop3sck_OnConnect()
    pbConnected = True
    RaiseEvent Pop3Status("")
End Sub

'Retrieves all waiting E-Mails and send the raw E-Mail to the
'ParseMail function

Private Sub Pop3sck_OnDataArrival(ByVal lngBytesTotal As Long)

  Dim strData As String

    '  Static intMessages          As Integer
    '  Static intCurrentMessage    As Integer
    '  Static strBuffer            As String
    'Dim intSwap As Integer
    
   ' On Error GoTo error

    'Retrieve, received from server, data.
    Pop3sck.GetData strData

    If Left$(strData, 1) = "+" Or m_State = Pop3_retr Then
        'If first symbol of server response is "+"
        'server has accepted previous client command
        'and it is waiting for next actions.
        Select Case m_State
            'This should be tohe most realistic case
          Case Pop3_retr
            '
            'Accumulate message data in strBuffer static variable

            'Set initial condition
            If Len(strBuffer) = 0 Then DataPointer = 1
            'Test to see if new string will fit within current strBuffer
            If (DataPointer + Len(strData)) > Len(strBuffer) Then
                'If not, allocate more memory
                strBuffer = strBuffer & Space$(Len(strData) + BlockSize)
            End If

            'Assign the new data
            Mid$(strBuffer, DataPointer, Len(strData)) = strData
            'Move pointer to end of new data
            DataPointer = DataPointer + Len(strData)
            '
            'Until we have been found single dot symbol on a line.
            If InStr(1, strBuffer, vbLf & "." & vbCrLf) Then
                '
                'OK! We have received a message.
                '
                'Remove server response string
                strBuffer = Mid$(strBuffer, InStr(1, strBuffer, vbCrLf) + 2)
                '
                'Remove dot symbol that is at the end of a message
                strBuffer = Left$(strBuffer, Len(strBuffer) - 3)
                '
                RaiseEvent Pop3Status("Decode Mail..." & CStr(intCurrentMessage))

                ParseMail strBuffer, intCurrentMessage
                

        
                '
                'Clear buffer for next message
                strBuffer = ""
                '
                If intCurrentMessage = intMessages Then
                    '
                    'We have received all messages, and
                    'we need say QUIT
                    AttachmentCounter = 0
                    intCurrentMessage = 1
                    
                    If bolDelMail Then
                        m_State = Pop3_dele
                        RaiseEvent Pop3Status("All mails received!")
                        Pop3sck.SendData "DELE " & intCurrentMessage & vbCrLf
                    Else
                        m_State = POP3_QUIT
                        RaiseEvent Pop3Status("All mails received!")
                        Pop3sck.SendData "QUIT" & vbCrLf
                    End If
                    

                  Else
                    '
                    'We have messages to download
                    'Increase message counter
                    intCurrentMessage = intCurrentMessage + 1
                    '
                    'Change current state of session
                    m_State = Pop3_retr
                    '
                    'Send RETR command to download next message
                    RaiseEvent Pop3Status("Receive next mail...")
                    Pop3sck.SendData "RETR " & _
                                     CStr(intCurrentMessage) & vbCrLf

                End If
            End If

          Case POP3_Connect
            '
            'Reset message counter
            intMessages = 0
            intCurrentMessage = 0
            '
            'Change current state of session
            m_State = POP3_USER
            '
            'Send to server USER command to tell him
            'which mailbox we want check out
            RaiseEvent Pop3Status("Authenticate User...")
            Pop3sck.SendData "USER " & m_strUsername & vbCrLf

          Case POP3_USER
            '
            'Change current state of session
            m_State = POP3_PASS
            '
            'Send password with PASS command
            RaiseEvent Pop3Status("Send Password...")
            Pop3sck.SendData "PASS " & m_strPassword & vbCrLf

          Case POP3_PASS
            '
            'Change current state of session
            m_State = POP3_STAT
            '
            'Send STAT command to know how many
            'messages in the mailbox
            RaiseEvent Pop3Status("Get Number of E-Mails...")
            Pop3sck.SendData "STAT" & vbCrLf

          Case POP3_STAT
            '
            'Parse server response to get number
            'of messages in the mailbox
            intMessages = CInt(Mid$(strData, 5, _
                          InStr(5, strData, " ") - 5))

            If intMessages > 0 Then

                'Redim Buffer to download all Mails
                ReDim Mails(intMessages - 1)
                '
                'OK! We have one or more.
                'Change current state of session
                m_State = Pop3_retr
                '
                'Increase counter to know wich message
                'we will retrieving
                intCurrentMessage = intCurrentMessage + 1
                '
                'And send RETR command to download
                'first message
                Pop3sck.SendData "RETR 1" & vbCrLf

              Else
                '
                'We have not any message in the mailbox.
                'Send QUIT command and show to user a message
                'that she or he has not mail.
                m_State = POP3_QUIT
                Pop3sck.SendData "QUIT" & vbCrLf

                RaiseEvent Pop3Status("You have not mail!")
            End If
            
          Case Pop3_dele
            If intCurrentMessage = intMessages Then
                m_State = POP3_QUIT
                Pop3sck.SendData "QUIT" & vbCrLf
            Else
                m_State = Pop3_dele
                intCurrentMessage = intCurrentMessage + 1
                Pop3sck.SendData "DELE " & intCurrentMessage & vbCrLf
                
            End If
            
          Case POP3_QUIT
            AttachmentCounter = 0
            RaiseEvent Pop3Status("")
            RaiseEvent ReceivedSuccesful
            Pop3sck.CloseSocket

        End Select
      Else
error:
        'Hide Status
        RaiseEvent Pop3Status("")
        'Show Error
        RaiseEvent MimeFailed(strData)
        Pop3sck.CloseSocket
    End If

End Sub

Private Sub Pop3sck_OnError(ByVal intNumber As Integer, strDescription As String, ByVal lngScode As Long, ByVal strSource As String, ByVal strHelpFile As String, ByVal lngHelpContext As Long, fCancelDisplay As Boolean)

    RaiseEvent MimeFailed("Winsock Error: #" & intNumber & "Desc: " & strDescription)

End Sub

Public Sub ParseMail(strMessage As String, MailCounter As Integer)

  Dim intPosA         As Long
  Dim intPosB         As Long
  Dim intPos          As Long
  Dim intCount        As Long
  Dim intFrom         As Long
  Dim intTo           As Long
  Dim intTemp         As Long
  Dim EndBoundary     As Long
  'Dim Counter         As Long
  Dim Counter2        As Long
  Dim vHeaders        As Variant
  Dim strTemp         As String
  Dim BoundArray      As Variant
  Dim strHeader       As String
  Dim strHeaderName   As String
  Dim strHeaderValue  As String
  Dim TmpString       As String
  Dim Boundary        As String
  Dim BoundaryVal     As String
  Dim strFilename     As String
  Dim MimeHeaders()   As String


    intPosA = InStr(1, strMessage, vbCrLf & vbCrLf)

    'A little Error Check
    If Not intPosA > 0 Then
        Exit Sub
    End If

    'Only the Mail Headers
    m_strHeaders = Left$(strMessage, intPosA - 1)
    'E-Mail + Attachments
    m_strMessageBody = Right$(strMessage, Len(strMessage) - intPosA - 3)
    'Whole E-Mail (Header + Message + Attachments)
    m_strMessageText = strMessage

    'Hmm I try to unfold the Mail Header...
    m_strHeaders = Replace(m_strHeaders, vbCrLf + Chr$(9), " ")
    m_strHeaders = Replace(m_strHeaders, vbCrLf + Chr$(11), " ")
    m_strHeaders = Replace(m_strHeaders, vbCrLf + Chr$(32), " ")
    m_strHeaders = Replace(m_strHeaders, vbCrLf + Chr$(255), " ")

    'Parse Mail Header and save data
    vHeaders = Split(m_strHeaders, vbCrLf)

    intFrom = LBound(vHeaders)
    intTo = UBound(vHeaders)

    For intTemp = intFrom To intTo
        strHeader = vHeaders(intTemp)
        intPosA = InStr(1, strHeader, ":")
        If intPosA Then
            strHeaderName = LCase$(Left$(strHeader, intPosA - 1))
          Else
            strHeaderName = ""
        End If

        strHeaderValue = Trim$(Right$(strHeader, Len(strHeader) - intPosA))

        With Mails(MailCounter - 1)

            Select Case LCase$(strHeaderName)
              Case "from"
                .from = DecodeIso(strHeaderValue)
              Case "to"
                .To = DecodeIso(strHeaderValue)
              Case "subject"
                .Subject = DecodeIso(strHeaderValue)
              Case "date"
                .Date = DecodeIso(strHeaderValue)
            End Select
        End With
    Next  'VFIELD INTTEMP

    intFrom = 0
    intTo = 0

    Mails(MailCounter - 1).Size = Len(m_strMessageText)
    Mails(MailCounter - 1).Header = m_strHeaders


'+++++++++++++++++++++ All Headers Processed, now decode the Mail!++++++++++++++++++
    'Load the Mail line by line into an array
    strlines = Split(m_strMessageText, vbCrLf)

    'Free some Memory
    m_strMessageText = ""
    m_strHeaders = ""

    'Search for Attachments
    Boundary = "boundary="

    intPosA = findLine(0, Boundary, strlines(), True)

    'Check if the Mail have Mime Attachments
    If intPosA = -1 Then
        GoTo Plaintext
    End If

    'Get all boundary Strings
    Do Until intPosA = -1
        intPosA = findLine(intPosA, Boundary, strlines(), True)

        If intPosA <> -1 Then
            strTemp = GetInfo(intPosA, Boundary, strlines())
            BoundaryVal = BoundaryVal + " " + "--" + strTemp
            intPosA = intPosA + 1
        End If

    Loop

    'Convert to Array
    BoundArray = Split(Trim$(BoundaryVal), " ")

    intFrom = LBound(BoundArray)
    intTo = UBound(BoundArray)

    'Now we extract all Attachments!
    intTemp = findLine(0, Boundary, strlines())
    
    For Counter2 = intFrom To intTo

        BoundaryVal = BoundArray(Counter2)

        intPosA = intTemp

        'Search Last Boundary
        EndBoundary = RevfindLine(BoundaryVal + "--", strlines())
        
        If EndBoundary = -1 Then
            EndBoundary = RevfindEmptyLine(strlines())
        End If

        Do Until intPosA >= EndBoundary

            intPosA = findLine(intPosA, BoundaryVal, strlines())
            intPosB = findLine(intPosA + 1, BoundaryVal, strlines())
            
            If intPosB = -1 Then
                intPosB = RevfindEmptyLine(strlines())
            End If
            
            intPos = findLine(intPosA, "Content-Type:", strlines())

            'Prevent extracting several "Sub"Attachments
            If intPos <> -1 Then
                If InStr(LCase$(strlines(intPos)), "boundary=") > 0 Then
                    GoTo Skip
                End If
            End If

            'Extract Attachment

            'First copy Mail to temp Array
            ptSpan = strlines

            'Move temp Array to destination array
            MoveStringArray ptSpan, strLine, intPosA + 1, intPosB - 1

            intCount = 0

            'This Part should be worked out => please Mail me your suggestions

            'It's pure US Plaintext
            If intPos = -1 Then
                TmpString = DecodeAttachment(strLine)
                Mails(MailCounter - 1).Message = TmpString

                GoTo Skip
            End If

            If InStr(LCase$(strlines(intPos)), "text/html") > 0 Then

                TmpString = DecodeAttachment(strLine)
                Mails(MailCounter - 1).HTMLMessage = TmpString
                GoTo Skip
            End If

            If InStr(LCase$(strlines(intPos)), "text") > 0 Then
                TmpString = DecodeAttachment(strLine)
                Mails(MailCounter - 1).Message = TmpString
                GoTo Skip
            End If

            If InStr(LCase$(strlines(intPos)), "multipart") > 0 Then
                TmpString = DecodeAttachment(strLine)
                Mails(MailCounter - 1).Message = TmpString
                GoTo Skip
            End If

            'Search the Filename
            intPos = findEmptyLine(0, strLine)

            If intPos <> -1 Then
                MimeHeaders = UnfoldArray(0, intPos, strLine)
                intPos = findLine(0, "name=", MimeHeaders, True)
                strFilename = GetInfo(intPos, "name=", MimeHeaders)
              Else
                intPos = findLine(0, "name=", strLine(), True)
                strFilename = GetInfo(intPos, "name=", strLine())
            End If

            strFilename = DecodeIso(strFilename)

            If strFilename = "" Then
                strFilename = "unnamed"
            End If

            'Save Attachment
            AddAttachment MailCounter - 1, strLine, strFilename
            AttachmentCounter = AttachmentCounter + 1

Skip:
            intPosA = intPosB
        Loop

    Next Counter2
    
    AttachmentCounter = 0

Exit Sub

Plaintext:

    intPos = findLine(1, "Content-Type:", strlines())

    m_strMessageBody = DecodeAttachment(strlines())

    If intPos > 0 Then
        If InStr(LCase$(strlines(intPos)), "text/html") > 0 Then
            Mails(MailCounter - 1).HTMLMessage = m_strMessageBody
          Else

            Mails(MailCounter - 1).Message = m_strMessageBody
        End If

      Else
        'Save the E-Mail
        Mails(MailCounter - 1).Message = m_strMessageBody
    End If

    AttachmentCounter = 0

    Erase strlines

End Sub

Public Function DecodeAttachment(ByRef Encoded() As String) As String

  Dim tmpEncoding   As String
  Dim tmpAttachment As String
  Dim intPosA       As Long
  Dim intPosB       As Long
  'Dim Counter       As Long
  Dim tmplong       As Long
  Dim Attachment() As String

    On Error GoTo error

    Attachment = Encoded

    tmplong = UBound(Attachment)

    If Not tmplong > 0 Then
        DecodeAttachment = ""
        Exit Function
    End If

    '1. What kind of Attachment is it?

    'Get Encoding-Type
    intPosA = findLine(0, "Content-Transfer-Encoding:", Attachment())

    If intPosA <> -1 Then
        tmpEncoding = GetInfo(intPosA, "Content-Transfer-Encoding:", Attachment())
      Else
        intPosA = 0
    End If

    'After the empty line the attachment waits for us!
    intPosA = findEmptyLine(intPosA, Attachment()) + 1

    'Extract Attachment

    '2. Decode it

    Select Case True
      Case InStr(LCase$(tmpEncoding), "base64") > 0

        'Very Fast Array Copy Routine (about 10x)!
        MoveStringArray Attachment, ptSpan, intPosA, tmplong

        tmpAttachment = Join(ptSpan, "")
        'tmpEncoding = tmpAttachment
        
        '2x times  faster (65 ms 120 ms)
        tmpAttachment = Decode(tmpAttachment)
        
        
      Case InStr(LCase$(tmpEncoding), "x-uue") > 0

        tmplong = RevfindLine("end", Attachment) - 1

        If tmplong = -1 Then
            tmplong = UBound(Attachment)
        End If
        
        'Very Fast Array Copy Routine (about 10x)!
        MoveStringArray Attachment, ptSpan, intPosA, tmplong

        
        intPosB = 0

        tmpAttachment = Join(ptSpan, vbCrLf)

        tmpAttachment = Replace(tmpAttachment, Chr$(0), "")

        tmpAttachment = DecodeUUE(tmpAttachment)

      Case InStr(LCase$(tmpEncoding), "quoted-printable") > 0

        'Very Fast Array Copy Routine (about 10x)!
        MoveStringArray Attachment, ptSpan, intPosA, tmplong

        tmpAttachment = Join(ptSpan, "=_")

        tmpAttachment = Replace(tmpAttachment, Chr$(0), "")

        tmpAttachment = DecodeQP(tmpAttachment)

      Case Else

        'Very Fast Array Copy Routine (about 10x)!
        MoveStringArray Attachment, ptSpan, intPosA, tmplong

        tmpAttachment = Join(ptSpan, vbCrLf)

        tmpAttachment = Replace(tmpAttachment, Chr$(0), "")

    End Select

    DecodeAttachment = tmpAttachment
    Erase ptSpan


Exit Function

error:
    DecodeAttachment = ""

End Function

'Saves the attachment into an UDT
Private Sub AddAttachment(intMail As Integer, strLine() As String, strFilename As String)

  Dim intElements As Integer
  Dim intBlockSize As Integer
  Dim intCounter As Integer

    On Error GoTo error

    intCounter = Mails(intMail).AttachedFiles + 1

    intElements = UBound(Mails(intMail).Attachments())

    If intElements > 0 Then

        intBlockSize = 10

        If intCounter - 1 > intElements Then
            ReDim Preserve Mails(intMail).Attachments(intElements + intBlockSize)
        End If

      Else
        'Initiate the Mail UDT for the first time
error:
        intBlockSize = 10
        ReDim Mails(intMail).Attachments(intBlockSize - 1)
    End If

    intElements = UBound(Mails(intMail).Attachments())

    'Save Attachment
    Mails(intMail).Attachments(intCounter - 1).Data = strLine
    Mails(intMail).Attachments(intCounter - 1).Name = DecodeIso(strFilename)
    Mails(intMail).AttachedFiles = intCounter

End Sub

'Clear all Mails
Public Sub ClearMails()

    Erase Mails

End Sub

''**************************************************************************************
' Base64 Decode
'
' This is an optimized version of the common Base 64 encode/decode.
' This version eliminates the repeditive calls to chr$() and asc(),
' as well as the linear searches I've seen in some routines.
'
' This method does use a bit more memory in permanent lookup tables
' than most do.  However, this eliminates the need for using vb's
' rather slow method of bit shifting (multiplication and division).
' This appears not to make much difference in the IDE, but make
' a huge difference in the exe.
'   Encodeing Index = 834 vs. 64 bytes standard
'   Decoding Index  = 1536 vs. 64 to 256 standard
'
' This routine also adds the CrLf on the fly rather than making
' a temporary copy of the encoded string then adding the crlf
'
' Encoding/Decoding data from and to a file should be changed to
' use a fixed buffer to reduce the memory requirements of EncodeFile, etc.
'
' All of this results in a speed increase:
'   Encode:
'         100 reps on a string of 28311 bytes
'                               IDE      EXE
'   Base64                      2824     300 (220 w/no overflow & array bound checks)
'   Base64a (unknown author)  375500* 185300*
'   Base64b (Wil Johnson)       2814     512 (410 w/no overflow & array bound checks)
'     *Extrapolated (based on 1 rep, I didn't have time to wait 30 minutes for 100)
'     *Unknown code is from ftp:altecdata.com/base64.cls
'
'  Decode
'         100 reps on a string of 28311 bytes
'                              IDE    EXE
'   Base64                    3384     351 (271 w/no overflow & array bound checks)
'   Base64a (unknown author)
'   Base64b (Wil Johnson)     5969    1191 (981 w/no overflow & array bound checks)
'   *Failed
'   *Unknown code is from ftp:altecdata.com/base64.cls
'
'
' Author: Tim Arheit - tarheit@wcoil.com
' Version: 1.0
'
' This code is provided as-is.  You are free to use and modify it
' as you wish.  Please report bugs, fixes and enhancements to the
' author.
'
' History:
'    11/13/00 - Code release. It appears to work.
'
'   09/02/02 I clean the source code and remove the encoding routines

'Decode a string to a string.
Public Function Decode(sInput As String) As String

  Dim bTemp() As Byte

    'Convert to a byte array then convert.
    'This is faster the repetitive calls to asc() or chr$()
    bTemp = StrConv(sInput, vbFromUnicode)

    Decode = StrConv(DecodeArr(bTemp), vbUnicode)

End Function

Public Sub DecodeToFile(sInput As String, sOutputFile As String)

  Dim bTemp() As Byte
  Dim fh As Long

    bTemp = StrConv(sInput, vbFromUnicode)
    bTemp = DecodeArr(bTemp)

    fh = FreeFile(0)
    Open sOutputFile For Binary Access Write As fh
    Put fh, , bTemp
    Close fh

End Sub

Public Sub DecodeFile(sInputFile As String, sOutputFile As String)

  Dim bTemp() As Byte
  Dim fh As Long

    fh = FreeFile(0)
    Open sInputFile For Binary Access Read As fh
    ReDim bTemp(0 To LOF(fh) - 1)
    Get fh, , bTemp
    Close fh

    bTemp = DecodeArr(bTemp)
    Open sOutputFile For Binary Access Write As fh
    Put fh, , bTemp
    Close fh

End Sub


Private Function DecodeArr(bInput() As Byte) As Byte()

  Dim bOutput() As Byte
  Dim OutLength As Long
  Dim CurrentOut As Long

  Dim k As Long
  Dim l As Long
  Dim I As Long
  

  Dim b As Byte
  Dim c As Byte
  Dim d As Byte
  Dim e As Byte

    k = LBound(bInput)
    l = UBound(bInput)

    'Calculate the length of the input
    I = l - k + 1

    'Allocate the output

  Dim BytesDataIn As Long ':( Move line to top of current Function
  Dim BytesDataOut As Long ':( Move line to top of current Function
  Dim ExtraBytes As Integer ':( Move line to top of current Function

    If bInput(l) = 61 Then
        ExtraBytes = 1
        If bInput(l - 1) = 61 Then
            ExtraBytes = 2
        End If
    End If

    BytesDataIn = l + 1 'BytesDataIn of the string
    BytesDataOut = (BytesDataIn * 0.75) - ExtraBytes ' how many bytes will the decoded string have

    ReDim bOutput(BytesDataOut - 1)

    CurrentOut = 0

    For I = k To l
        Select Case bInput(I)
          Case CHAR_CR
            'Do nothing
          Case CHAR_LF
            'Do nothing
          Case Else
            If l - I >= 3 Then
                b = bInput(I)
                c = bInput(I + 1)
                d = bInput(I + 2)
                e = bInput(I + 3)

                If e <> CHAR_EQUAL Then
                    bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                    bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
                    bOutput(CurrentOut + 2) = m_ReverseIndex3(d, 1) Or m_ReverseIndex4(e)
                    CurrentOut = CurrentOut + 3
                    I = I + 3 ':( Modifies active For-Variable
                  ElseIf d <> CHAR_EQUAL Then 'NOT E...
                    bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                    bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
                    CurrentOut = CurrentOut + 2
                    I = I + 3 ':( Modifies active For-Variable
                  Else 'NOT D...
                    bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                    CurrentOut = CurrentOut + 1
                    I = I + 3 ':( Modifies active For-Variable
                End If

              Else 'NOT L...
                'Possible input code error, but may also be
                'an extra CrLf, so we will ignore it.
            End If
        End Select
    Next I

    'On properly formed input we should have to do this.
    If OutLength <> CurrentOut + 1 Then
        ReDim Preserve bOutput(0 To CurrentOut - 1)
    End If

    DecodeArr = bOutput

End Function




'Saves a String to a File
Public Sub SaveStr2File(strInput As String, strPathName As String)

  Dim iFreeFile As Integer

    '-----
    ' Reference to a free file
    '-----
    iFreeFile = FreeFile
    Open strPathName For Binary As iFreeFile
    '-----
    ' Save the total size of the array in a variable, this stops
    ' VB to calculate the size each time it comes into the loop,
    ' which of course, takes (much) more time then this sollution
    '-----

    Put iFreeFile, , strInput

    Close iFreeFile

End Sub

'==========
' StrToAry;
' Convert the string into a byte array
'==========
Public Sub StringToByteArray(ByVal strIn As String, ByRef pbArrayOutput() As Byte)

    pbArrayOutput = StrConv(strIn, vbFromUnicode)

End Sub

'==========
' AryToSr;
' Convert the byte array into a string
'==========
Public Sub ByteArrayToString(ByRef pbArrayInput() As Byte, ByRef strOut As String)

    strOut = StrConv(pbArrayInput, vbUnicode)

End Sub



Public Function StringArrayToString(pbIn() As String) As String

  Dim lSize     As Long
  Dim lNow      As Long
  Dim lTotal    As Long
  Dim lNowArray As Long
  Dim lNow2     As Long
  Dim lTotal2   As Long
  Dim tTemp     As String
  Dim bTemp()   As Byte

    '-----
    ' Calculate size of inputarray
    '-----
    lSize = 0
    lTotal = UBound(pbIn)
    For lNow = 0 To lTotal
        lSize = lSize + Len(pbIn(lNow))
    Next lNow

    '-----
    ' Create byte array which is big
    ' enough to hold all the bytes
    '-----
    ReDim bTemp(0 To lSize)

    '-----
    ' Convert the string array to a byte array
    '-----
    lNow = 0
    lNowArray = 0
    While lNow <> lSize
        tTemp = pbIn(lNowArray)
        lTotal2 = Len(tTemp)

        '-----
        ' Loop through the temp string
        ' and place the byte character
        ' in the correct position
        ' Mid$(...) is faster then Mid$(...)
        '-----
        For lNow2 = 0 To lTotal2 - 1
            bTemp(lNow + lNow2) = Asc(Mid$(tTemp, lNow2 + 1, 1))
        Next lNow2

        lNow = lNow + lTotal2
        lNowArray = lNowArray + 1
    Wend

    '-----
    ' Convert byte array to string
    '-----
    StringArrayToString = StrConv(bTemp(), vbUnicode)

End Function

'**************************************************************************************
'UUE decoding class
'
'Author: PSC
'
'Desc:
'
'This class have several routines that support the decoding UU
'encoded attachments
''**************************************************************************************

Public Function DecodeUUE(strUUCodeData As String) As String

  Dim vDataLine   As Variant
  Dim vDataLines  As Variant
  Dim strDataLine As String
  Dim intSymbols  As Integer
  Dim strTemp     As String
  Dim strUUDecode As String
  Dim I As Long

    On Error GoTo error

    'remove begin marker
    If Left$(strUUCodeData, 6) = "begin " Then
        strUUCodeData = Mid$(strUUCodeData, InStr(1, strUUCodeData, vbLf) + 1)
    End If
    '
    'remove end marker
    If Right$(strUUCodeData, 5) = "end" + vbCrLf Then
        strUUCodeData = Left$(strUUCodeData, Len(strUUCodeData) - 10)
    End If

    'Split encoded data to vDataLines array.
    'Now each array member contains a line of encoded data
    vDataLines = Split(strUUCodeData, vbCrLf)

    For Each vDataLine In vDataLines
        'Decode data line by line
        '
        strDataLine = CStr(vDataLine)

        If strDataLine = "" Then
            GoTo Skip
        End If

        'Get quantity of symbols in a line
        intSymbols = Asc(Left$(strDataLine, 1)) - 32
        'remove first symbol that just informs
        'about quantity of symbols
        strDataLine = Mid$(strDataLine, 2)
        'replace "`" with " "
        strDataLine = Replace(strDataLine, "`", " ")
        'Convert every 4-byte chunk to 3-byte chunk by
        For I = 1 To Len(strDataLine) Step 4
            '1 byte
            strTemp = strTemp + Chr$((Asc(Mid$(strDataLine, I, 1)) - 32) * 4 + _
                      (Asc(Mid$(strDataLine, I + 1, 1)) - 32) \ 16)
            '2 byte
            strTemp = strTemp + Chr$((Asc(Mid$(strDataLine, I + 1, 1)) Mod 16) * 16 + _
                      (Asc(Mid$(strDataLine, I + 2, 1)) - 32) \ 4)
            '3 byte
            strTemp = strTemp + Chr$((Asc(Mid$(strDataLine, I + 2, 1)) Mod 4) * 64 + _
                      Asc(Mid$(strDataLine, I + 3, 1)) - 32)
        Next I
        '
        strTemp = Left$(strTemp, intSymbols)
        'write decoded line to the file
        strUUDecode = strUUDecode + strTemp
        'clear buffer for next line
        strTemp = ""
Skip:
    Next vDataLine

    DecodeUUE = strUUDecode
error:

End Function

'**************************************************************************************
'Quoted printable decoding class
'
'Author: PSC
'
'Desc:
'
'This class have several routines that support the decoding of quoted printable
'encoded attachments
''**************************************************************************************

Public Function DecodeQP(ByRef StrToDecode As String) As String

  Dim sTemp As String
  Dim strChar As String
  Dim I As Integer

    sTemp = StrToDecode

    sTemp = Replace(sTemp, "==_", "")
    sTemp = Replace(sTemp, "=_", vbCrLf)

    For I = 255 To 16 Step -1
        strChar = UCase$(Hex$(I))
        If InStr(1, sTemp, "=" & strChar) <> 0 Then sTemp = Replace(sTemp, "=" & strChar, Chr$(I) + Chr$(0))
    Next I

    For I = 15 To 1 Step -1
        strChar = UCase$(Hex$(I))
        If InStr(1, sTemp, "=" & "0" & strChar) <> 0 Then
            sTemp = Replace(sTemp, "=" & "0" & strChar, Chr$(I) + Chr$(0))
        End If
    Next I

    sTemp = Replace(sTemp, Chr$(0), "")
    sTemp = Replace(sTemp, "=00", Chr$(0))
    sTemp = Replace(sTemp, Chr$(255) & Chr$(254), "=")

    DecodeQP = sTemp

End Function

'*************************************************************************************
'Function to decode ?iso? encoded Strings
'
'
'Author: David Bue Pedersen + Sebastian Fahrenkrog
'*************************************************************************************
Function DecodeIso(strEncoded As String)

  'Dim StrtoReplace As String
  'Dim StrReplacement As String
  Dim StringtoDecode As String
  Dim strLookup As String
  Dim b As Boolean

    On Error GoTo error

    StringtoDecode = strEncoded

    If IsNull(StringtoDecode) Then
        Exit Function
      ElseIf InStr(1, LCase$(StringtoDecode), "=?iso-") <= 0 Then
        DecodeIso = StringtoDecode
        Exit Function
    End If

  Dim IsoArray As Variant
  Dim UCounter As Integer
  Dim Counter As Integer
  Dim Pattern As String

    IsoArray = StringtoDecode
    IsoArray = Split(StringtoDecode, "?")

    UCounter = UBound(IsoArray)

    For Counter = 0 To UCounter
        strLookup = IsoArray(Counter)
        Select Case strLookup
          Case "="
          Case "= ="
          Case "=="
          Case "Q"
            b = False
          Case "B"
            b = True
          Case Else
            Pattern = "ISO-" & "[0-9]" & "[0-9]" & "[0-9]" & "[0-9]" & "-" & "*"
            If Not UCase(IsoArray(Counter)) Like Pattern Then
                If b Then 'Decode Base64
                    StringtoDecode = IsoArray(Counter)
                    StringtoDecode = Decode(StringtoDecode)
                    DecodeIso = DecodeIso + StringtoDecode
                  Else 'dann quoted printable
                    StringtoDecode = IsoArray(Counter)
                    StringtoDecode = DecodeQP(StringtoDecode)
                    DecodeIso = DecodeIso + StringtoDecode
                End If
            End If
        End Select
    Next Counter

Exit Function

error:
    'Return original String
    DecodeIso = strEncoded

End Function

Public Property Let DelMail(bolDeleteMail As Boolean)
    bolDelMail = bolDeleteMail
End Property

Public Property Get DelMail() As Boolean
    DelMail = bolDelMail
End Property

Private Function WaitUntilTrue(ByRef Flag As Boolean, ByVal SecondsToWait As Long, Optional ByVal bRaiseTimeOutError As Boolean = True) As Boolean

  Dim fStart              As Single
  Dim fTimetoQuit         As Single

    If SecondsToWait < 1 Then Exit Function

    fStart = Timer

    ' Deal with timer being reset at Midnight
    If fStart + SecondsToWait < 86400 Then
        fTimetoQuit = fStart + SecondsToWait
      Else
        fTimetoQuit = (fStart - 86400) + SecondsToWait
    End If

    Do Until Flag = True
        If Timer >= fTimetoQuit Then
            If bRaiseTimeOutError Then Timeout
            Exit Function
        End If
        If pbExitImmediately Then Exit Function
        DoEvents
        Sleep (10)                                  ' added to reduce CPU load during wait periods
    Loop

    WaitUntilTrue = Flag

End Function

Private Sub Timeout()

  ' time out occured
  
        'Hide Status
        RaiseEvent Pop3Status("")
        'Show Error
        RaiseEvent MimeFailed("Can't connect to the server!")
        Pop3sck.CloseSocket
        
        
End Sub

