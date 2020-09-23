VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PhoneBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adress Register"
   ClientHeight    =   5355
   ClientLeft      =   2415
   ClientTop       =   2055
   ClientWidth     =   6705
   ClipControls    =   0   'False
   Icon            =   "PhoneBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEdit 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   4
      Left            =   4200
      TabIndex        =   20
      ToolTipText     =   "Press to add a new post to the database"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "AddNew"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   3
      Left            =   5400
      TabIndex        =   17
      ToolTipText     =   "Press to add a new post to the database"
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   2
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Press to delete the current Post"
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   1
      Left            =   5400
      TabIndex        =   15
      ToolTipText     =   "Press to enable AddNew"
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Index           =   0
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Press to update the current post"
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Frame frmEditMode 
      Caption         =   "Editmode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
      Begin VB.OptionButton optEditMode 
         Caption         =   "Editable"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optEditMode 
         Caption         =   "Readable"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmSearch 
      Caption         =   "Search post"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4200
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
      Begin VB.OptionButton optSearch 
         Caption         =   "Cellular"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Adress"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Country"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "City"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Telephone"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Company"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "FirstName"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "LastName"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Go"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "You can use % as wildcard"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmSelPers 
      Caption         =   "Select post to view"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstSelpers 
         Height          =   1230
         ItemData        =   "PhoneBook.frx":0442
         Left            =   120
         List            =   "PhoneBook.frx":0444
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   21
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "PhoneBook.frx":0446
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPers(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPers(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPers(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPers(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPers(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPers(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblPers(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPers(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPers(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPers(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPers(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPers(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdWebEmail(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdWebEmail(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdMove(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdMove(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdMove(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdMove(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtPers(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtPers(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtPers(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtPers(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtPers(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPers(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPers(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtPers(8)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPers(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtPers(10)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Info"
      TabPicture(1)   =   "PhoneBook.frx":0462
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPers(12)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Photo"
      TabPicture(2)   =   "PhoneBook.frx":047E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPers(11)"
      Tab(2).Control(1)=   "cmdPhotopath"
      Tab(2).Control(2)=   "Image1"
      Tab(2).ControlCount=   3
      Begin VB.TextBox txtPers 
         Height          =   4695
         Index           =   12
         Left            =   -74880
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton cmdPhotopath 
         Caption         =   "Browse"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -71880
         TabIndex        =   40
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   11
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         Top             =   4320
         Width           =   2655
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   10
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         Top             =   4200
         Width           =   3735
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   8
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   36
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   7
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   6
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   4
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   1
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "I<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Move to the first post"
         Top             =   4560
         Width           =   635
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   720
         TabIndex        =   27
         ToolTipText     =   "Move to the next post"
         Top             =   4560
         Width           =   635
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   1320
         TabIndex        =   26
         ToolTipText     =   "Move to the previous post"
         Top             =   4560
         Width           =   635
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   ">I"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   1920
         TabIndex        =   25
         ToolTipText     =   "Move to the last post"
         Top             =   4560
         Width           =   635
      End
      Begin VB.CommandButton cmdWebEmail 
         Height          =   540
         Index           =   1
         Left            =   3360
         Picture         =   "PhoneBook.frx":049A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Go to the person in this post webpage"
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdWebEmail 
         Height          =   540
         Index           =   0
         Left            =   2640
         Picture         =   "PhoneBook.frx":08DC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Send a mail to the person in this post"
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtPers 
         Height          =   285
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3630
         Left            =   -74760
         Top             =   480
         Width           =   3585
      End
      Begin VB.Label lblPers 
         Caption         =   "Webpage adress"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label lblPers 
         Caption         =   "Email"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Cellular"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   50
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Telephone"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Country"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   48
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "City"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "PostalCode"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   46
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Adress"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Firstname"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   44
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Lastname"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblPers 
         Caption         =   "Company"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenAdressRegister 
         Caption         =   "&Open Adress Register"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCreateAdressRegister 
         Caption         =   "&Create Adress Register"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Adress Register"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRestoreBackup 
         Caption         =   "&Restore Backup"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuMailDeveloper 
         Caption         =   "&Mail Developer"
      End
      Begin VB.Menu mnuWebDeveloper 
         Caption         =   "&Developers Webpage"
      End
   End
End
Attribute VB_Name = "PhoneBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objRs As ADODB.Recordset     'The recordset object
Private conString As String           'The string to use in objRs.ActiveConnection (what database to open)
Private bolEdit As Boolean            'Tells what kind of locktype to use in recordset
Private WhereString As String         'What to get in the recordset (used in the search function)
Private WhereVal As String            'What column to use in the wherestring
Private bolSearch As Boolean          'Tells if you are searching or not (to be used if the db is empty)
Private AdressRegisterPath As String  'Tells the path to the choosen Adressregister

Private CD1 As New cmDlg
Private CD12 As New cmDlg
Private CDCreateOpen2 As New cmDlg

'***Open Database***'
Private Sub OpenDatabase()


    mnuBackup.Enabled = True
    optEditMode(0).Enabled = True
    optEditMode(1).Enabled = True
    cmdSearch.Enabled = True
    cmdMove(0).Enabled = True
    cmdMove(1).Enabled = True
    cmdMove(2).Enabled = True
    cmdMove(3).Enabled = True

    conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AdressRegisterPath & _
                ";Persist Security Info=False"
    Set objRs = New ADODB.Recordset
    OpenRs

End Sub

'***Show the person in the current record***'
Private Sub showCurrentRec()

  Dim I As Integer

    With objRs 'Fill the textboxes with the record
        For I = 1 To .Fields.Count - 1
            txtPers(I - 1).Text = .Fields(I) & ""
        Next I
    End With
    On Error GoTo errHandler 'In case the photopath is wrong
    Image1.Picture = LoadPicture(txtPers(11).Text) 'Set the picture = the photopath

errHandler:
    If Err.Number = 53 Then 'Wrong photopath
        MsgBox "The Picture of this person" & vbCrLf & _
               "Seems to not exist or the path is wrong !"
    End If

End Sub



'***Move within the recordset***'
Private Sub cmdMove_Click(index As Integer)

    On Error GoTo error

    Select Case index
      Case 0 'move to the first record
        objRs.MoveFirst
      Case 1 'move to next record
        objRs.MoveNext
      Case 2 'move to previous record
        objRs.MovePrevious
      Case 3 'move to the last record
        objRs.MoveLast
    End Select
    If objRs.BOF Then objRs.MoveFirst 'if it is the beginning of the file move to the first record
    If objRs.EOF Then objRs.MoveLast 'if it is the end of the file move to the last record
    showCurrentRec
error:

End Sub

'***Get the recordset***'
Private Sub OpenRs()

    On Error GoTo errHandler
    With objRs
        If .State = adStateOpen Then .Close 'if it is open close it

        .ActiveConnection = conString 'to which database to connect to
        .CursorLocation = adUseClient   'Use the cursor on the client
        .CursorType = adOpenKeyset 'Moveable recordset in any direction
        Select Case bolEdit
          Case False 'Readmode
            .LockType = adLockReadOnly 'Read only recordset
          Case True 'Editmode
            .LockType = adLockOptimistic 'Editable recordset
        End Select
        .Source = "select * from tblPhonebook " & WhereString & " order by lastname" 'What to get
        .Open
    End With

    listPers
    objRs.MoveFirst
    showCurrentRec
errHandler:
    If Err.Number = 3021 Then 'if the recordset holds no records (empty database or nothing found in the search)
        If bolSearch = False Then 'Empty database
            NoPostInDb
          Else 'Nothing found in the search
            MsgBox "No records found"
            WhereString = ""
            txtSearch.Text = ""
            cmdEdit(4).Enabled = False
            cmdEdit(4).Caption = ""
            OpenRs
        End If

      ElseIf Err.Number = -2147467259 Then 'if the database is missing
        mnuRestoreBackup_Click
      ElseIf Err.Number <> 0 Then 'in any other error tell what have happen
        MsgBox Err.Number & " " & Err.Description
    End If

End Sub

'***Routine for adding a new post in an empty database
Private Sub NoPostInDb()

  Dim I As Integer

    If MsgBox("You have no posts in your Adress Register!" & vbCrLf & _
       "Do you want to add a new post ?", vbYesNo, "Add a new post") = vbYes Then
        bolEdit = True
        cmdPhotopath.Enabled = True
        For I = 0 To 12
            txtPers(I).Locked = False
        Next I
        For I = 0 To 3 'enable/disable editbuttons
            cmdEdit(I).Enabled = bolEdit
        Next I
        If bolEdit = True Then cmdEdit(3).Enabled = False
        cmdEdit_Click (1)
        MsgBox "Add a new post in your Adress Register" & vbCrLf & _
               "Press AddNew when you are done", , "Add a new post"
      Else
        Exit Sub
    End If

    With objRs
        If .State = adStateOpen Then .Close 'if it is open close it

        .ActiveConnection = conString 'what database to connect to
        .CursorLocation = adUseClient 'Use the clients cursor
        .CursorType = adOpenKeyset 'Moveable recordset in any direction
        .LockType = adLockOptimistic 'Editable recordset
        .Source = "select * from tblPhonebook order by lastname" 'What to get
        .Open
    End With

End Sub

'***List lastname, firstname in the listbox***'
Private Sub listPers()

    lstSelpers.Clear 'empty it first, no duplicates

    With objRs
        .MoveFirst
        While Not .EOF
            lstSelpers.AddItem .Fields(1) & " " & .Fields(2)
            .MoveNext
        Wend
    End With

End Sub

'***Browse to the photopath to store in db***'
Private Sub cmdPhotopath_Click()

    CD1.InitDir = App.Path 'where it should begin to look
    CD1.ShowOpen 'Open the dialog
    txtPers(11).Text = CD1.FileName 'Set the pathname
    Image1.Picture = LoadPicture(CD1.FileName) 'set the picture, to see if it is correct

End Sub

'*** Send mail to person or goto the webpage***'
Private Sub cmdWebEmail_Click(index As Integer)

    frmMail.txtTo = txtPers(9)
    frmMail.Show
    Unload Me

End Sub

Private Sub Form_Load()

    Set objRs = Nothing

    CDCreateOpen2.InitDir = App.Path
    CDCreateOpen2.DialogTitle = "Open Adress Register"
    CDCreateOpen2.FileName = App.Path + "\adressbook.adr"
    AdressRegisterPath = CDCreateOpen2.FileName
    OpenDatabase

    optEditMode_Click (1)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set CD1 = Nothing
    Set CD12 = Nothing
    Set CDCreateOpen2 = Nothing

End Sub

'***On click move to the selected record and show it***'
Private Sub lstSelPers_Click()

    objRs.MoveFirst
    objRs.Move (lstSelpers.ListIndex)
    showCurrentRec

End Sub

'***Make a backup of the Adress register***'
Private Sub mnuBackup_Click()

  Dim strTemp As String
  Dim I As Integer

    On Error GoTo errHandler
    Set objRs = Nothing
    CD12.DialogTitle = "Where do you want to put your backup ?"

    For I = 1 To Len(AdressRegisterPath) - 1
        If Mid(AdressRegisterPath, I, 1) = "\" Then
            strTemp = Mid(AdressRegisterPath, 1, I)
        End If
    Next I
    CD12.FileName = Mid(AdressRegisterPath, Len(strTemp) + 1)
    CD12.ShowSave

    If CD12.FileName <> "" Then FileCopy AdressRegisterPath, CD12.FileName
    CD12.FileName = ""
    OpenDatabase
errHandler:
    Set objRs = New ADODB.Recordset

End Sub

'***Create a new adress register***'
Private Sub mnuCreateAdressRegister_Click()

    Set objRs = Nothing

    CDCreateOpen2.InitDir = App.Path
    CDCreateOpen2.DialogTitle = "Create Adress Register as"
    CDCreateOpen2.ShowSave
    If CDCreateOpen2.FileName <> "" Then
        FileCopy App.Path & "\TEMPLATE.bak", CDCreateOpen2.FileName
        AdressRegisterPath = CDCreateOpen2.FileName
        OpenDatabase
    End If

End Sub

'***Select a adress register to open***'
Private Sub mnuOpenAdressRegister_Click()

    Set objRs = Nothing

    CDCreateOpen2.InitDir = App.Path
    CDCreateOpen2.DialogTitle = "Open Adress Register"
    CDCreateOpen2.ShowOpen
    AdressRegisterPath = CDCreateOpen2.FileName
    OpenDatabase

End Sub

'***Restore the AdressRegister***'
Private Sub mnuRestoreBackup_Click()

  Dim strTemp As String
  Dim I As Integer

    On Error GoTo errHandler
    Set objRs = Nothing
    CD12.DialogTitle = "Select Adress Register to restore"
    CD12.ShowOpen
    If CD12.FileName <> "" Then
        AdressRegisterPath = CD12.FileName

        For I = 1 To Len(AdressRegisterPath) - 1
            If Mid(AdressRegisterPath, I, 1) = "\" Then
                strTemp = Mid(AdressRegisterPath, 1, I)
            End If
        Next I
        strTemp = "\" & Mid(AdressRegisterPath, Len(strTemp) + 1)
        FileCopy CD12.FileName, App.Path & strTemp
    End If
    OpenDatabase

errHandler:
    Set objRs = New ADODB.Recordset

End Sub

'***Exit***'
Private Sub mnuExit_Click()

    Unload Me

End Sub

'***Set what kind of recordset to get***'
Private Sub optEditMode_Click(index As Integer)

  Dim I As Integer

    Select Case index
      Case 0 'Readable recordset
        bolEdit = False
        cmdPhotopath.Enabled = False
        For I = 0 To 12
            txtPers(I).Locked = True
        Next I
      Case 1 'Editable recordset
        bolEdit = True
        cmdPhotopath.Enabled = True
        For I = 0 To 12
            txtPers(I).Locked = False
        Next I
    End Select
    For I = 0 To 3 'enable/disable editbuttons
        cmdEdit(I).Enabled = bolEdit
    Next I
    If bolEdit = True Then cmdEdit(3).Enabled = False
    WhereString = ""
    OpenRs

End Sub

'***Set what column to use in the where criteria, also work as search***'
Private Sub optSearch_Click(index As Integer)

    WhereVal = optSearch(index).Caption

End Sub

'***Create part of the string to use in the recordset source***'
Private Sub cmdSearch_Click()

    If WhereVal = "" Then WhereVal = "LastName"
    bolSearch = True
    WhereString = " Where " & WhereVal & " Like '" & txtSearch.Text & "'"
    cmdEdit(4).Enabled = True
    cmdEdit(4).Caption = "Get all posts"
    OpenRs
    bolSearch = False

End Sub

'***Update, Delete, AddNew record and clear textboxes***'
Private Sub cmdEdit_Click(index As Integer)

  Dim I As Integer
  Dim bookMark As Variant

    Select Case index
      Case 0 'Edit and update current record
        If txtPers(0).Text = "" Then
            MsgBox "you must enter a value in Lastname !"
            txtPers(0).SetFocus
          ElseIf txtPers(1).Text = "" Then
            MsgBox "you must enter a value in Firstname !"
            txtPers(1).SetFocus
          Else
            bookMark = objRs.bookMark 'Set bookMark to the current record
            For I = 0 To 12
                If txtPers(I) = "" Then 'Dont store an empty string
                    objRs.Fields(I + 1) = Null
                  Else
                    objRs.Fields(I + 1) = Trim(txtPers(I).Text)
                End If
            Next I
            objRs.Update
            listPers
            objRs.bookMark = bookMark
            showCurrentRec
        End If
      Case 1 'Clear the texboxes and enable AddNew
        cmdEdit(3).Enabled = True
        cmdEdit(0).Enabled = False
        cmdEdit(2).Enabled = False
        cmdEdit(4).Enabled = True
        cmdEdit(4).Caption = "Disable AddNew"
        cmdPhotopath.Enabled = True
        For I = 0 To 12
            txtPers(I).Text = ""
        Next I
      Case 2 'Delete current record
        If MsgBox("Do you want to delete the Post" & vbCrLf & _
           objRs.Fields(1) & " " & objRs.Fields(2) & " ?", vbOKCancel) = vbOK Then
            objRs.Delete adAffectCurrent
            objRs.Requery 'refresh the recordset
            If objRs.RecordCount = 0 Then 'If it was the only record
                For I = 0 To 12
                    txtPers(I).Text = ""
                Next I
                lstSelpers.Clear
                NoPostInDb 'Routine for making a new record in an empty database
              Else
                listPers
                objRs.MoveLast
                showCurrentRec
            End If
        End If
      Case 3 'Addnew, Add a new record to DB
        If txtPers(0).Text = "" Then
            MsgBox "you must enter a value in Lastname !"
            txtPers(0).SetFocus
          ElseIf txtPers(1).Text = "" Then
            MsgBox "you must enter a value in Firstname !"
            txtPers(1).SetFocus
          Else
            objRs.AddNew
            For I = 0 To 12
                If txtPers(I) = "" Then 'Dont store empty strings
                    objRs.Fields(I + 1) = Null
                  Else
                    objRs.Fields(I + 1) = Trim(txtPers(I).Text)
                End If
            Next I
            objRs.Update
            objRs.Requery 'Refresh the recordset
            listPers
            objRs.MoveLast
            showCurrentRec
            cmdEdit(3).Enabled = False 'disable the Addnew cmdbutton
            cmdEdit(0).Enabled = True
            cmdEdit(2).Enabled = True
        End If

      Case 4 'Get Records back after search
        WhereString = ""
        txtSearch.Text = ""
        OpenRs
        If bolEdit = True Then
            cmdEdit(3).Enabled = False
            cmdEdit(0).Enabled = True
            cmdEdit(2).Enabled = True
        End If
        cmdEdit(4).Enabled = False
        cmdEdit(4).Caption = ""
    End Select

End Sub

':) Ulli's VB Code Formatter V2.12.7 (26.06.2002 19:52:39) 12 + 451 = 463 Lines
