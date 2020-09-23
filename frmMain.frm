VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pop3 Popper"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TV1 
      Height          =   2535
      Left            =   1680
      TabIndex        =   12
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4471
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   6720
      Width           =   10560
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00997367&
      Height          =   6975
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   6915
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   0
      Width           =   1455
      Begin VB.Image ImgRecuiter 
         Height          =   720
         Left            =   480
         Picture         =   "frmMain.frx":030A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   600
      End
      Begin VB.Image imgQuery 
         Height          =   345
         Left            =   480
         Picture         =   "frmMain.frx":0614
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label lblRecruiters 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get E-Mails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   720
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblQuery 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Create E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   765
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Image imgPrevious 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":068A
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblPreviousQuery 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   765
         Left            =   0
         TabIndex        =   8
         Top             =   4440
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2175
      Left            =   1680
      TabIndex        =   2
      Top             =   3360
      Width           =   9735
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save As..."
         Height          =   375
         Left            =   8400
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin ComctlLib.ListView lvAttachments 
         Height          =   2175
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6588
      TabWidthStyle   =   1
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Message"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Attachments"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser HtmlMail 
      Height          =   2535
      Left            =   2280
      TabIndex        =   6
      Top             =   3240
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   4471
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame4 
      Caption         =   "Messages"
      Height          =   2895
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin ComctlLib.ListView lvMessages 
         Height          =   2535
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Subject"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Attachments"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Menu m_Messages 
      Caption         =   "&Messages"
      Begin VB.Menu cmdCheckMailbox 
         Caption         =   "Check Mailbox"
      End
      Begin VB.Menu cmdnewMail 
         Caption         =   "Create new E-Mail"
      End
      Begin VB.Menu m_SaveMessage 
         Caption         =   "Save E-Mail Text"
      End
      Begin VB.Menu cmdDelselMessage 
         Caption         =   "Delete selected Message"
      End
      Begin VB.Menu cmdReplyMessage 
         Caption         =   "Reply selected Message"
      End
      Begin VB.Menu Strich 
         Caption         =   "-"
      End
      Begin VB.Menu m_Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mAccount 
      Caption         =   "&Account"
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      Begin VB.Menu m_MailHeader 
         Caption         =   "Show Rfc822 Header"
      End
   End
   Begin VB.Menu m_language 
      Caption         =   "Language"
      Begin VB.Menu mEnglish 
         Caption         =   "English"
      End
      Begin VB.Menu mGerman 
         Caption         =   "German"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intMailSelected As Integer
Private ComDialog As New cmDlg
Private Conn As New ADODB.Connection


'Declare Events for the vbMime Class
Private WithEvents Mime As vbMime
Attribute Mime.VB_VarHelpID = -1

Sub OpenConn() 'Connection string :-)

    Conn.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & App.Path & "\data.mdb"

End Sub

Sub CompactDatabase() 'DBP We compact the MDB. The MDB Dosent shrink as records is delteted. So... We have to do everything ourselves

  Dim JRO As JRO.JetEngine

    On Error GoTo error

    Set JRO = New JRO.JetEngine

    JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\tempbase.mdb" & ";Jet OLEDB:Engine Type=5"
    Kill App.Path & "\data.mdb"
    Name App.Path & "\tempbase.mdb" As App.Path & "\data.mdb"
    Set JRO = Nothing

Exit Sub

error:
    MsgBox "The programm could not open the E-Mail database." & vbCrLf & _
           "Please close all programms and try again!"
    End

End Sub

Public Sub cmdCheckMailbox_Click()

  Dim c As Control
  Dim Pop3Server As String, Pop3Username As String, Pop3Password As String

    'Check up textboxes frmmain
    For Each c In frmOptions.Controls
        If TypeOf c Is TextBox Then
            If Len(c.Text) = 0 Then
                MsgBox "Please check your Account Settings!"
                frmOptions.Show
                Exit Sub
            End If
        End If
    Next c

    For Each c In Controls
        If TypeOf c Is Image Then
            c.Enabled = False
        End If

        If TypeOf c Is Label Then
            c.Enabled = False
        End If

    Next c

    cmdCheckMailbox.Enabled = False

    With frmOptions
        'Set property if the mails received should be deleted or not
        Mime.DelMail = .chkDelMails.Value
        'Go and get it tiger! GRRR!
        Mime.GetMail .txtUsername, .txtPassword, .txtPop3Server
    End With

    'Query Database and retreive the Account Info then Get All E-Mails!
    'Set rsAccount = Conn.Execute("Select * from accounts")

    'Do Until rsAccount.EOF
    '    Pop3Server = rsAccount("pop3server")
    '    Pop3Username = rsAccount("username")
    '    Pop3Password = rsAccount("password")

    '    Mime.GetMail Pop3Username, Pop3Password, Pop3Server
    '    rsAccount.MoveNext
    'Loop

End Sub

'Display all E-Mail Data
Public Sub ShowMail()

  Dim lvItem As ListItem
  Dim rsMail As New ADODB.Recordset

    On Error Resume Next

      Me.lvAttachments.ListItems.Clear
      Me.lvMessages.ListItems.Clear

      'Query the Database and get all Mail Infos
      Set rsMail = Conn.Execute("Select * from mails")

      Do Until rsMail.EOF
          Set lvItem = lvMessages.ListItems.Add
          lvItem.Text = rsMail("From")
          lvItem.SubItems(1) = rsMail("Subject")
          lvItem.SubItems(2) = rsMail("Date")
          lvItem.SubItems(3) = rsMail("Size")
          lvItem.Tag = rsMail("id")
          rsMail.MoveNext
      Loop

End Sub

'Convert an String to HTML File
Public Sub TextToHTML(strInputMessage As String, strOutputFile As String, strTitle As String, strBgcolor As String, strTextcolor As String)

  Dim Newline As String

    Newline = Chr$(13) + Chr$(10)

    Open strOutputFile For Output As #2
    If strTitle = "" Then
        strTitle = "No Document Title"
    End If
    If strBgcolor = "" Then
        strBgcolor = "white"
    End If
    If strTextcolor = "" Then
        strTextcolor = "black"
    End If

    ' Replaces common symbols
    strInputMessage = Replace$(strInputMessage, "&", "&amp;")
    strInputMessage = Replace$(strInputMessage, "<", "&lt;")
    strInputMessage = Replace$(strInputMessage, ">", "&gt;")
    strInputMessage = Replace$(strInputMessage, Chr$(34), "&quot;")
    strInputMessage = Replace$(strInputMessage, "ß", "&szlig;")

    Print #2, "<HTML>" + Newline
    Print #2, "<HEAD>" + Newline
    Print #2, "<TITLE>" + strTitle + "</TITLE>" + Newline
    Print #2, "</HEAD>" + Newline
    Print #2, "<BODY bgcolor=" + strBgcolor + " text=" + strTextcolor + ">" + Newline
    Print #2, Replace(strInputMessage, vbCrLf, "<BR>")
    Print #2, Newline
    Print #2, "</BODY>" + Newline
    Print #2, "</HTML>"
    Close #2

End Sub

Private Sub cmdDelselMessage_Click()

    DeleteMessage (intMailSelected)

End Sub

Private Sub cmdnewMail_Click()

    frmMail.Show

End Sub

Private Sub cmdReplyMessage_Click()

  Dim Message As String
  Dim strfrom As String
  Dim rsMail As New ADODB.Recordset
  Dim strQuery As String

    On Error GoTo error

    'Query the Database and get all Mail Infos
    strQuery = "Select * from mails where id=" _
               & lvMessages.ListItems(intMailSelected).Tag

    Set rsMail = Conn.Execute(strQuery)

    Message = "You wrote on the last E-Mail:" & vbCrLf & vbCrLf & rsMail("Message")
    strfrom = rsMail("From")
    frmMail.rtfMail = Message
    frmMail.txtTo = strfrom
    frmMail.Show

error:

End Sub

Private Sub cmdSave_Click()

  Dim strFilename As String
  Dim strAttachment As String
  Dim rsAttachments As New ADODB.Recordset
  Dim strQuery As String

    On Error GoTo error

    strQuery = "Select * from attachments where id=" _
               & lvAttachments.ListItems(intMailSelected).Tag

    Set rsAttachments = Conn.Execute(strQuery)

    strFilename = rsAttachments("filename")

    With ComDialog

        .FileName = Replace(strFilename, vbCrLf, "")
        .ShowSave

        If Err = 0 Then

            strAttachment = rsAttachments("filedata")

            If Trim(strAttachment) <> "" Then
                Mime.SaveStr2File strAttachment, .FileName
              Else
error:
                MsgBox "This Attachment could not be decoded. Sorry!"
            End If
        End If

    End With

End Sub

Private Sub Form_Activate()

  Dim c As Control
  Dim strTemp As String

    On Error Resume Next

      strTemp = LoadIni("Properties", "Language")
      If InStr(strTemp, "English") > 0 Then
          mEnglish.Checked = True
          mEnglish_Click
        Else
          mGerman.Checked = True
          mGerman_Click
      End If

      For Each c In frmOptions.Controls
          If TypeOf c Is TextBox Then
              c.Text = LoadIni("Account", c.Name)
          End If

          If TypeOf c Is CheckBox Then
              c.Value = LoadIni("Account", c.Name)
          End If
      Next c

      ShowMail

End Sub

'This Routine must only be executed once!
Private Sub Form_Initialize()

    CompactDatabase
    OpenConn
    'Start Trapping Right-Mouse clicks in WebBrowser Control:
    gLngMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseHookProc, App.hInstance, GetCurrentThreadId)
    'Preload Mail Form
    Load frmMail

    TV1.Nodes.Add , , "main", "Messages" 'Create Main Parent
    TV1.Nodes.Add "main", tvwChild, "Inbox1", "Inbox" 'Child Node to the Main Parent or ROOT
    TV1.Nodes.Add "main", tvwChild, "Outbox1", "Outbox"
    TV1.Nodes.Add "main", tvwChild, "Sent1", "Sent"
    TV1.Nodes.Add "main", tvwChild, "Trash1", "Trash"

    TV1.Nodes.Item(1).Expanded = True 'expands the 1st or ROOT node

PerfInit
End Sub

Private Sub Form_Load()

  'Load Dialogs

    Load frmStatus
    Load frmOptions

    'Initiate Mail Decoding Class
    Set Mime = New vbMime

    With TabStrip1
        HtmlMail.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        Frame5.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
    End With

    HtmlMail.ZOrder 0

    'Initialize Browser with empty site
    Open App.Path & "\Temp.html" For Output As #1 'open a file for output
    Print #1, ""
    Close #1 'closes that file

    HtmlMail.Navigate App.Path & "\Temp.html"

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  'this is the place to control the buttons

    If Button = 2 Then 'if they right click, 1=left, 2=right
        'Me.PopupMenu mnuPopUp 'show popup menu
        Me.PopupMenu m_Messages
      Else 'else if they clicked the left button
        DoEvents

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Mime = Nothing
    Set ComDialog = Nothing
    End

    'Cancel the trapping of the code
    UnhookWindowsHookEx gLngMouseHook

End Sub

Private Sub HtmlMail_NewWindow2(ppDisp As Object, Cancel As Boolean)

    Cancel = True

End Sub

Private Sub imgPrevious_Click()

    PhoneBook.Show

End Sub

Private Sub imgQuery_Click()

    frmMail.Show

End Sub

Private Sub ImgRecuiter_Click()

    Me.cmdCheckMailbox_Click

End Sub

Private Sub lblPreviousQuery_Click()

    PhoneBook.Show

End Sub

Private Sub lblQuery_Click()

    frmMail.Show

End Sub

Private Sub lblRecruiters_Click()

    Me.cmdCheckMailbox_Click

End Sub

Private Sub lvAttachments_ItemClick(ByVal Item As ComctlLib.ListItem)

    intMailSelected = Item.index

End Sub

Private Sub lvMessages_ItemClick(ByVal Item As ComctlLib.ListItem)

    Show_Message (Item.index)

End Sub

Private Sub lvMessages_KeyDown(KeyCode As Integer, Shift As Integer)

  'If someone press Del

    If KeyCode = 46 Then
        DeleteMessage (intMailSelected)
    End If

End Sub

Private Sub DeleteMessage(MailtoDelete As Integer)

  Dim rsMail As New ADODB.Recordset
  Dim rsAttachments As New ADODB.Recordset
  Dim SearchTag As String

    On Error GoTo error

    rsMail.Open "mails", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
    rsAttachments.Open "attachments", Conn, adOpenKeyset, adLockOptimistic, adCmdTable

    SearchTag = lvMessages.ListItems(MailtoDelete).Tag

    Do Until rsMail.EOF

        If SearchTag = rsMail("id") Then
            rsMail.Delete
            Exit Do
        End If
        rsMail.MoveNext
    Loop

    Do Until rsAttachments.EOF

        If SearchTag = rsAttachments("email") Then
            rsAttachments.Delete
        End If
        rsAttachments.MoveNext
    Loop

    'Remove Mail from Listview Box
    lvMessages.ListItems.Remove (MailtoDelete)
    lvAttachments.ListItems.Clear
    HtmlMail.Navigate "about:blank"

Exit Sub

error:
    MsgBox "Can not delete mail!"

End Sub

Private Sub lvMessages_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  'this is the place to control the buttons

    If Button = 2 Then 'if they right click, 1=left, 2=right
        'Me.PopupMenu mnuPopUp 'show popup menu
        Me.PopupMenu m_Messages
      Else 'else if they clicked the left button
        DoEvents

    End If

End Sub

Private Sub m_Exit_Click()

    Unload frmStatus
    Unload frmOptions
    Unload Me
    'End

End Sub

Private Sub m_MailHeader_Click()

    On Error GoTo error

    Me.m_MailHeader.Checked = Not Me.m_MailHeader.Checked
    Show_Message intMailSelected

error:

End Sub

Private Sub m_SaveMessage_Click()

  Dim strTemp As String
  Dim HTMLMessage As String
  Dim Message As String
  Dim rsMail As New ADODB.Recordset
  Dim strQuery As String
  Dim strfrom As String, strTo As String, strSubject As String

    'Query the Database and get all Mail Infos
    strQuery = "Select * from mails where id=" _
               & lvMessages.ListItems(intMailSelected).Tag

    Set rsMail = Conn.Execute(strQuery)

    Message = rsMail("Message")
    HTMLMessage = rsMail("HTMLMessage")
    strfrom = rsMail("From")
    strTo = rsMail("To")
    strSubject = rsMail("Subject")

    On Error GoTo error

    If HTMLMessage <> "" Then
        strTemp = "From: " + strfrom + vbCrLf + "To:" + strTo + vbCrLf + "Subject: " _
                  + strSubject + vbCrLf + vbCrLf + HTMLMessage
      Else
        strTemp = "From: " + strfrom + vbCrLf + "To:" + strTo + vbCrLf + "Subject: " _
                  + strSubject + vbCrLf + vbCrLf + Message
    End If

    With ComDialog
        On Error GoTo error

        .FileName = "Mail.txt"
        .ShowSave

        If Err = 0 Then
            Mime.SaveStr2File strTemp, .FileName
        End If

    End With

Exit Sub

error:
    MsgBox "Sorry, can't save Message!"

End Sub

Private Sub mAccount_Click()

    frmOptions.Show

End Sub

Private Sub mEnglish_Click()

  ' Lets load the entire language pack. It doesn't apply the language pack in the form.

    mEnglish.Checked = True
    mGerman.Checked = False

    cLanguage.LoadLanguagePack App.Path & "\english.lpk"
    ' Now it applies the language pack in the form
    cLanguage.SetLanguageInForm Me
    cLanguage.SetLanguageInForm frmOptions
    cLanguage.SetLanguageInForm frmMail

    SaveIni "Properties", "Language", "English"

End Sub

Private Sub mGerman_Click()

  ' Lets load the entire language pack. It doesn't apply the language pack in the form.

    mEnglish.Checked = False
    mGerman.Checked = True
    cLanguage.LoadLanguagePack App.Path & "\german.lpk"
    ' Now it applies the language pack in the form
    cLanguage.SetLanguageInForm Me
    cLanguage.SetLanguageInForm frmOptions
    cLanguage.SetLanguageInForm frmMail

    SaveIni "Properties", "Language", "German"

End Sub

Private Sub Mime_MimeFailed(Explanation As String)

  Dim c As Control

    MsgBox Explanation

    For Each c In Controls
        If TypeOf c Is Image Then
            c.Enabled = True
        End If

        If TypeOf c Is Label Then
            c.Enabled = True
        End If

    Next c

    cmdCheckMailbox.Enabled = True

End Sub

Private Sub Mime_Pop3Status(Status As String)

    With frmStatus
        If Status <> "" Then

            .Show
            .Status = Status
            lstStatus.AddItem Status
          Else
            .Hide
            .Status = ""
            frmMain.Show
        End If
    End With

End Sub

Private Sub Mime_ReceivedSuccesful()

  'All the Mails are ready to show

  Dim b As Long
  Dim Counter As Integer
  Dim intFrom As Integer
  Dim intTo As Integer
  Dim intAttachFrom As Integer
  Dim intAttachTo As Integer
  Dim strAttachment As String
  Dim rsMail As New ADODB.Recordset
  Dim rsAttachments As New ADODB.Recordset
  Dim c As Control



    For Each c In Controls
        If TypeOf c Is Image Then
            c.Enabled = True
        End If

        If TypeOf c Is Label Then
            c.Enabled = True
        End If

    Next c

    cmdCheckMailbox.Enabled = True

    'If there is no mail the Array is not existent ->error
    On Error GoTo No_Mail

    intFrom = LBound(Mails)
    intTo = UBound(Mails)

    On Error GoTo error

    rsMail.Open "mails", Conn, adOpenKeyset, adLockOptimistic, adCmdTable
    rsAttachments.Open "attachments", Conn, adOpenKeyset, adLockOptimistic, adCmdTable

    'Save every received Mail in the Database
    For Counter = intFrom To intTo

        With Mails(Counter)

            rsMail.AddNew

            If frmOptions.txtUsername <> "" Then
                rsMail("owner") = frmOptions.txtUsername
            End If

            rsMail("to") = .To
            rsMail("from") = .from
            rsMail("date") = .Date
            rsMail("size") = .Size
            rsMail("subject") = .Subject
            rsMail("message") = .Message
            rsMail("HTMLMessage") = .HTMLMessage
            rsMail("Header") = .Header
            rsMail.Update

            If .AttachedFiles > 0 Then
                intAttachFrom = LBound(.Attachments)
                intAttachTo = UBound(.Attachments)

                For b = intAttachFrom To intAttachTo
                    If .Attachments(b).Name = "" Then GoTo Skip:

                    rsAttachments.AddNew
                    rsAttachments("email") = rsMail("id")
                    rsAttachments("filename") = .Attachments(b).Name
                     
                    strAttachment = Mime.DecodeAttachment(.Attachments(b).Data)
                    
                    rsAttachments("filedata").AppendChunk (strAttachment)
                    rsAttachments.Update

Skip:
                Next b

            End If

        End With

    Next Counter

    Mime.ClearMails

    'NOW SHOW THE USER THE RECEIVED E-MAILS
    ShowMail

No_Mail:

Exit Sub

error:
    MsgBox "An error occured during the saving of the E-Mails!"
    ShowMail

End Sub

Private Sub Show_Message(intMailNumber As Integer, Optional ShowHeader As Boolean)

  Dim lvItem As ListItem
  Dim strTxtBody As String
  Dim strTemp As String
  Dim strQuery As String
  Dim HTMLMessage As String
  Dim Message As String
  Dim Header As String
  Dim rsMail As New ADODB.Recordset
  Dim rsAttachments As New ADODB.Recordset

    On Error Resume Next

      'Query the Database and get all Mail Infos
      strQuery = "Select * from mails where id=" _
                 & lvMessages.ListItems(intMailNumber).Tag

      Set rsMail = Conn.Execute(strQuery)

      strQuery = "Select * from attachments where email=" _
                 & lvMessages.ListItems(intMailNumber).Tag

      Set rsAttachments = Conn.Execute(strQuery)

      Message = rsMail("Message")
      HTMLMessage = rsMail("HTMLMessage")
      Header = rsMail("Header")

      ShowHeader = Me.m_MailHeader.Checked

      'Adds the Message to the Internet Explorer Box
      If HTMLMessage = "" Then

          If ShowHeader Then
              strTxtBody = Header + vbCrLf + vbCrLf + Message
            Else
              strTxtBody = Message
          End If

          Call TextToHTML(strTxtBody, App.Path & "\Temp.html", "", "", "")

          HtmlMail.Navigate App.Path & "\Temp.html"

        Else
          If ShowHeader Then
              strTxtBody = Header + vbCrLf + vbCrLf + HTMLMessage
            Else
              strTxtBody = HTMLMessage
          End If

          Open App.Path & "\Temp.html" For Output As #1 'open a file for output
          Print #1, strTxtBody 'print text (HTML Code) into it
          Close #1 'closes that file

          HtmlMail.Navigate App.Path & "\Temp.html"

      End If

      'Adds the Attachmentlist to the Listview
      lvAttachments.ListItems.Clear

      Do Until rsAttachments.EOF
          strTemp = rsAttachments("filename")

          If strTemp <> "" Then
              Set lvItem = lvAttachments.ListItems.Add
              lvItem.Text = strTemp
              lvItem.SubItems(1) = "?"
              lvItem.Tag = rsAttachments("id")
          End If

          rsAttachments.MoveNext
      Loop

      intMailSelected = intMailNumber

End Sub

Private Sub Show_Sent()

  
End Sub

Private Sub TabStrip1_Click()

    If TabStrip1.SelectedItem.index = 1 Then
        HtmlMail.ZOrder 0
      Else
        Frame5.ZOrder 0
    End If

End Sub

Private Sub ShowSent()

  Dim intCounter As Integer
  Dim intFrom As Integer
  Dim intTo As Integer
  Dim lvItem As ListItem
  Dim rsMail As New ADODB.Recordset

    On Error Resume Next

      Me.lvAttachments.ListItems.Clear
      Me.lvMessages.ListItems.Clear

      'Query the Database and get all Mail Infos
      Set rsMail = Conn.Execute("Select * from Sent")

      Do Until rsMail.EOF
          Set lvItem = lvMessages.ListItems.Add
          lvItem.Text = rsMail("From")
          lvItem.SubItems(1) = rsMail("Subject")
          lvItem.SubItems(2) = rsMail("Date")
          lvItem.SubItems(3) = rsMail("Size")
          lvItem.Tag = rsMail("id")
          rsMail.MoveNext
      Loop

End Sub

Private Sub TV1_NodeClick(ByVal Node As MSComCtlLib.Node)

    Select Case Node.Key

      Case "Inbox1"
        ShowMail
      Case "Outbox1"
        'ShowQueue
      Case "Sent1"
        'ShowSent
      Case "Trash1"
        'ShowTrash
    End Select

End Sub

':) Ulli's VB Code Formatter V2.12.7 (28.06.2002 15:27:58) 7 + 882 = 889 Lines
