VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Option Explicit

Private Const CHAR_EQUAL As Byte = 61
Private Const CHAR_CR As Byte = 13
Private Const CHAR_LF As Byte = 10

Private m_ReverseIndex1(0 To 255) As Byte
Private m_ReverseIndex2(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex3(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex4(0 To 255) As Byte


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


