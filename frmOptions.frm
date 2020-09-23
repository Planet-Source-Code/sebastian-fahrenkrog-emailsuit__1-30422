VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Options"
   Begin VB.CheckBox chkDelMails 
      Caption         =   "Delete received mails"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   5640
      Width           =   5415
   End
   Begin VB.CheckBox ckLogin 
      Caption         =   "Use SMTP Login"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      ToolTipText     =   "Use Login Authorization When Connecting to a Host"
      Top             =   5040
      Width           =   5235
   End
   Begin VB.CheckBox ckPopLogin 
      Caption         =   "Use POP Login before sending an E-Mail"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      ToolTipText     =   "Use Login Authorization When Connecting to a Host"
      Top             =   5340
      Width           =   5415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3000
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   6120
      Width           =   3015
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   9
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   8
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   6
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   0
      Left            =   210
      ScaleHeight     =   4572.581
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame Frame1 
         Caption         =   "From Information"
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   5535
         Begin VB.TextBox txtfromaddress 
            Height          =   285
            Left            =   1560
            TabIndex        =   22
            Top             =   840
            Width           =   3735
         End
         Begin VB.TextBox txtfromname 
            Height          =   285
            Left            =   1560
            TabIndex        =   21
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label2 
            Caption         =   "From Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "From Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtPop3Server 
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtServer 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lbl_popserver 
         Caption         =   "POP3 Server:"
         Height          =   255
         Left            =   660
         TabIndex        =   14
         Top             =   1140
         Width           =   975
      End
      Begin VB.Image img_popserver 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":0000
         Top             =   960
         Width           =   480
      End
      Begin VB.Image img_popusername 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":030A
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image img_poppassword 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":0FD4
         Top             =   2460
         Width           =   480
      End
      Begin VB.Label lbl_popusername 
         Caption         =   "Username:"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label lbl_poppassword 
         Caption         =   "Password:"
         Height          =   195
         Left            =   660
         TabIndex        =   12
         Top             =   2580
         Width           =   735
      End
      Begin VB.Image img_server 
         Height          =   480
         Left            =   120
         Picture         =   "frmOptions.frx":12DE
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbl_server 
         Caption         =   "SMTP Server:"
         Height          =   255
         Left            =   660
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   5925
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10451
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mail Transport"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

    Me.Hide

End Sub

Private Sub cmdOK_Click()

  Dim c As Control

    For Each c In Controls
        If TypeOf c Is TextBox Then
            SaveIni "Account", c.Name, c.Text
        End If
        If TypeOf c Is CheckBox Then
            SaveIni "Account", c.Name, c.Value
        End If
    Next c

    Me.Hide

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim i As Integer

    i = tbsOptions.SelectedItem.index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
          Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
      ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
          Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If

End Sub

Private Sub tbsOptions_Click()

  Dim i As Integer

    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
          Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next i

End Sub



':) Ulli's VB Code Formatter V2.12.7 (19.06.2002 23:13:04) 2 + 81 = 83 Lines
