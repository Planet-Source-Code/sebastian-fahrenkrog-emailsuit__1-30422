VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMail 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "New Message"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   10680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtTo 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   8775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attach file"
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   6360
      Width           =   5895
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "&Add..."
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox lstAttachments 
         Height          =   645
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox txtBcc 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   8775
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   1560
      Width           =   8775
   End
   Begin RichTextLib.RichTextBox rtfMail 
      Height          =   4215
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7435
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMail.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7080
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0082
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0194
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":02A6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":03B8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":04CA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":05DC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":06EE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0800
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0912
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0A24
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0B36
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0C48
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0D5A
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0E6C
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPreviousQuery 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   525
      Left            =   9840
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image imgPrevious 
      Height          =   345
      Left            =   10200
      Picture         =   "frmMail.frx":0F7E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   315
   End
   Begin VB.Label Label4 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Bcc:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu newMail 
         Caption         =   "New Mail"
      End
      Begin VB.Menu SendMail 
         Caption         =   "Send E-Mail"
      End
      Begin VB.Menu strich00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Import txt"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Message"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Printer Page Setup"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu cmdAttachment 
      Caption         =   "&Attachment"
      Begin VB.Menu cmdAttachfile 
         Caption         =   "Attach file"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu format 
      Caption         =   "Format"
      Begin VB.Menu CheckBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu CheckItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu CheckStrikeLine 
         Caption         =   "Strike Line"
      End
      Begin VB.Menu Line 
         Caption         =   "-"
      End
      Begin VB.Menu mHtmlMail 
         Caption         =   "Send Mail as HTML Mail"
      End
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Win32 Declarations for Print sub
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_USER = &H400
Const EM_CANUNDO = &HC6
Const EM_UNDO = &HC7

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long     ' First character of range (0 for start of doc)
    cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long       ' Actual DC to draw on
    hdcTarget As Long ' Target DC for determining text formatting
    rc As RECT        ' Region of the DC to draw to (in twips)
    rcPage As RECT    ' Region of the entire DC (page size) (in twips)
    chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private ComDialog As New cmDlg
' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Private WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

Private bolHtmlMail As Boolean

Private Sub CheckBold_Click()

    CheckBold.Checked = Not CheckBold.Checked
    rtfMail.SelBold = CheckBold.Checked
    
If CheckBold.Checked = True Then
    mHtmlMail.Checked = True
    bolHtmlMail = True
Else
    mHtmlMail.Checked = False
    bolHtmlMail = False
End If

If CheckBold.Checked Then
    tbToolBar.Buttons("Bold").Value = tbrPressed
Else
    tbToolBar.Buttons("Bold").Value = tbrUnpressed
End If
End Sub

Private Sub CheckItalic_Click()

    CheckItalic.Checked = Not CheckItalic.Checked
    rtfMail.SelItalic = CheckItalic.Checked
    
If CheckItalic.Checked = True Then
    mHtmlMail.Checked = True
    bolHtmlMail = True
Else
    mHtmlMail.Checked = False
    bolHtmlMail = False
End If

If CheckItalic.Checked Then
    tbToolBar.Buttons("Italic").Value = tbrPressed
Else
    tbToolBar.Buttons("Italic").Value = tbrUnpressed
End If

End Sub

Private Sub CheckStrikeLine_Click()

    CheckStrikeLine.Checked = Not CheckStrikeLine.Checked
    rtfMail.SelUnderline = CheckStrikeLine.Checked
    
If CheckStrikeLine.Checked = True Then
    mHtmlMail.Checked = True
    bolHtmlMail = True
Else
    mHtmlMail.Checked = False
    bolHtmlMail = False
End If

If CheckStrikeLine.Checked Then
    tbToolBar.Buttons("Underline").Value = tbrPressed
Else
    tbToolBar.Buttons("Underline").Value = tbrUnpressed
End If


End Sub

Private Sub cmdAddFile_Click()

    On Error GoTo error

    With ComDialog

        .ShowOpen
        

        If Err = 0 Then

            If Trim(.FileName) <> "" Then

                lstAttachments.AddItem .FileName
              Else
error:
                Exit Sub
            End If
        End If

    End With

End Sub

Private Sub cmdAttachfile_Click()

    Call cmdAddFile_Click

End Sub

Private Sub cmdRemove_Click()

    On Error Resume Next

      lstAttachments.RemoveItem lstAttachments.ListIndex

End Sub

Private Sub FilePageSetup_Click()

End Sub

Private Sub Form_Activate()
Load_LastMail
End Sub

Private Sub Form_Load()

  'Initiate vbSendMail.cls

    Set poSendMail = New clsSendMail

End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' *****************************************************************************
  ' Unload the component before quiting.
  ' *****************************************************************************

    Set poSendMail = Nothing
    Set ComDialog = Nothing
End Sub

Private Sub imgPrevious_Click()
PhoneBook.Show
End Sub

Private Sub lblPreviousQuery_Click()
PhoneBook.Show
End Sub

Private Sub lstAttachments_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim Counter As Integer

    For Counter = 1 To Data.Files.Count
        If (GetAttr(Data.Files.Item(Counter)) And vbDirectory) = 0 Then lstAttachments.AddItem Data.Files.Item(Counter)
    Next Counter

End Sub

Private Sub mHtmlMail_Click()
    mHtmlMail.Checked = Not mHtmlMail.Checked
    bolHtmlMail = Not bolHtmlMail
End Sub

Private Sub mnuFilePrint_Click()

    PrintRTF rtfMail, 720, 720, 720, 720

End Sub

Private Sub mnuFileSave_Click()

  Dim strTemp As String

    On Error GoTo error

    With ComDialog
        On Error GoTo error

        .FileName = "Message.txt"
        .ShowSave

        If Err = 0 Then
            SaveStr2File strTemp, .FileName
        End If

    End With

Exit Sub

error:
    MsgBox "Sorry, can't save Message!"

End Sub

Private Sub newMail_Click()

  Dim c As Control

    'Clear all fields
    For Each c In Me.Controls
        If TypeOf c Is TextBox Then
            c.Text = ""
        End If
    Next c

    rtfMail.TextRTF = ""

    lstAttachments.Clear

End Sub

Private Sub SendMail_Click()

  Dim I As Integer
  Dim ulimit As Integer
  Dim m_strAttachedFiles As String
  Dim strTemp As String
  Dim c As Control

    On Error GoTo error

    'Error Handler
    If Me.txtTo = "" Then
        MsgBox "Please enter an E-Mail Address!"
        Exit Sub
    End If

    'Check up textboxes frmmain
    For Each c In frmOptions.Controls
        If TypeOf c Is TextBox Or TypeOf c Is ComboBox Then
            If Len(c.Text) = 0 Then
                MsgBox "Please check your Account Settings!"
                frmOptions.Show
                Exit Sub
            End If
        End If
    Next c

    'Read all Attachments
    ulimit = lstAttachments.ListCount

    Select Case ulimit

      Case Is > 1
        For I = 0 To ulimit - 1
            
            m_strAttachedFiles = lstAttachments.List(I) + ";" + m_strAttachedFiles
        Next I
            'Cut the ; from the rest
            If Right$(m_strAttachedFiles, 1) = ";" Then
                m_strAttachedFiles = Left$(m_strAttachedFiles, Len(m_strAttachedFiles) - 1)
            End If
      Case 1
            I = 0
            m_strAttachedFiles = lstAttachments.List(I)

    End Select

    Me.Hide
    frmStatus.Show
    
    'Convert the mail from rtf to html
    
    If bolHtmlMail Then
        strTemp = rtfMail.TextRTF
        strTemp = rtf2html.rtf2html(strTemp, "+H")
    Else
        strTemp = rtfMail.Text
    End If
    
     Save_LastMail
   
    
    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = validate_none 'VALIDATE_HOST_DNS     ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = frmOptions.txtServer            ' Required the fist time, optional thereafter
        .from = frmOptions.txtfromaddress           ' Required the fist time, optional thereafter
        .FromDisplayName = frmOptions.txtfromname   ' Optional, saved after first use
        .Recipient = Me.txtTo                       ' Required, separate multiple entries with delimiter character
        .Subject = Me.txtSubject                    ' Optional
        .Message = strTemp                  ' Optional
        .Attachment = Trim(m_strAttachedFiles)      ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bolHtmlMail                             ' Optional, default = FALSE, send mail as html or plain text
        .UseAuthentication = frmOptions.ckLogin.Value             ' Optional, default = FALSE
        .UsePopAuthentication = frmOptions.ckPopLogin.Value      ' Optional, default = FALSE
        .Username = frmOptions.txtUsername          ' Optional, default = Null String
        .Password = frmOptions.txtPassword                     ' Optional, default = Null String, value is NOT saved
        .POP3Host = frmOptions.txtPop3Server

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        .send                                       ' Required

    End With
    
   
    Unload frmStatus

Exit Sub

error:
    MsgBox "Sorry an error occurred while sending the mail!"

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)

  

    On Error Resume Next
      Select Case Button.Key
        Case "New"
          newMail_Click
        Case "Open"
          mnuFileOpen_Click
        Case "Save"
          mnuFileSave_Click
        Case "Print"
          PrintRTF rtfMail, 720, 720, 720, 720
        Case "Cut"
          mnuEditCut_Click
        Case "Copy"
          mnuEditCopy_Click
        Case "Paste"
          mnuEditPaste_Click
        Case "Bold"
          CheckBold_Click
        Case "Italic"
          CheckItalic_Click
        Case "Underline"
            
          CheckStrikeLine_Click
        Case "Align Left"
          rtfMail.SelAlignment = rtfLeft
          rtfMail.SetFocus
          bolHtmlMail = False
          Me.mHtmlMail.Checked = False
        Case "Center"
          rtfMail.SelAlignment = rtfCenter
          rtfMail.SetFocus
          bolHtmlMail = True
          Me.mHtmlMail.Checked = True
        Case "Align Right"
          rtfMail.SelAlignment = rtfRight
          rtfMail.SetFocus
          bolHtmlMail = True
          Me.mHtmlMail.Checked = True
      End Select

End Sub

Private Sub mnuViewOptions_Click()

    frmOptions.Show vbModal, Me

End Sub









Private Sub mnuEditPaste_Click()

    rtfMail.SelText = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()

    If rtfMail.SelLength > 0 Then
        Clipboard.SetText rtfMail.SelText
    End If

End Sub

Private Sub mnuEditCut_Click()

    If rtfMail.SelLength > 0 Then
        Clipboard.Clear
        Clipboard.SetText rtfMail.SelText
        rtfMail.SelText = ""
    End If

End Sub



Private Sub mnuFileExit_Click()

  'unload the form

    Unload Me

End Sub

Private Sub mnuFilePageSetup_Click()

    On Error Resume Next
      With ComDialog
          .DialogTitle = "Page Setup"
          .CancelError = True
          .ShowPrinter
      End With

End Sub



Private Sub mnuFileOpen_Click()

  Dim sFile As String

    With ComDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Import Message (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
        rtfMail.LoadFile sFile
    End With

End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

  ' vbSendMail 'Progress Event'

    With frmMain
        .lstStatus.AddItem lPercentCompete
        .lstStatus.ListIndex = .lstStatus.ListCount - 1
        .lstStatus.ListIndex = -1
    End With

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

  ' vbSendMail 'SendFailed Event

    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    frmStatus.Hide

End Sub

Private Sub poSendMail_SendSuccesful()

  ' vbSendMail 'SendSuccesful Event'

    frmStatus.Hide
    Unload frmMail

End Sub

Private Sub poSendMail_Status(Status As String)

  ' vbSendMail 'Status Event'

    With frmMain
        .lstStatus.AddItem Status
        .lstStatus.ListIndex = .lstStatus.ListCount - 1
        .lstStatus.ListIndex = -1
    End With

    frmStatus.Status = Status

End Sub

Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)

  '** Description:
  '** Print the active document

    On Error GoTo PrintError
  Dim LeftOffset As Long, TopOffset As Long
  Dim LeftMargin As Long, TopMargin As Long
  Dim RightMargin As Long, BottomMargin As Long
  Dim fr As FormatRange
  Dim rcDrawTo As RECT
  Dim rcPage As RECT
  Dim TextLength As Long
  Dim NextCharPosition As Long
  Dim r As Long

    ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips

    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight

    ' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin

    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text

    ' Get length of text in RTF
    TextLength = Len(RTF.Text)

    ' Loop printing each page until done
    Do
        ' Print the page by sending EM_FORMATRANGE message
        NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do  'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page
        Printer.NewPage                  ' Move on to next page
        Printer.Print Space(1) ' Re-initialize hDC
        fr.hdc = Printer.hdc
        fr.hdcTarget = Printer.hdc
    Loop

    ' Commit the print job
    Printer.EndDoc

    ' Allow the RTF to free up memory
    r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
PrintError:

End Sub

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

Private Sub Save_LastMail()
Dim MailNumber As Integer

If Not CheckExistence(txtTo, CStr(txtTo)) Then
    MailNumber = txtTo.ListCount
    If MailNumber > 10 Then MailNumber = 9
    SaveIni "Last Addresses", CStr(MailNumber), txtTo.Text
End If
End Sub

Private Sub Load_LastMail()
Dim Counter As Integer
Dim strTemp As String

'Load Last 10 Adresses
For Counter = 9 To 0 Step -1
    strTemp = LoadIni("Last Addresses", CStr(Counter))
    If strTemp <> "" Then
        txtTo.AddItem strTemp
    End If
Next

End Sub

':) Ulli's VB Code Formatter V2.12.7 (19.06.2002 23:12:58) 43 + 526 = 569 Lines
