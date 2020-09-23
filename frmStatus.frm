VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Status"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Status 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Option Explicit
Private WithEvents poSendMail As clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

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

':) Ulli's VB Code Formatter V2.12.7 (19.06.2002 23:13:00) 7 + 50 = 57 Lines
