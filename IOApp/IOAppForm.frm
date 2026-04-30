VERSION 5.00
Begin VB.Form frmIOApp
   Caption         =   "I/O App"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart
      Caption         =   "Start"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdStop
      Caption         =   "Stop"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox lstLog
      Height          =   3585
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6720
   End
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   240
   End
   Begin VB.Label lblStatusValue
      Caption         =   "Stopped"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2000
   End
   Begin VB.Label lblStatusLabel
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   800
   End
End
Attribute VB_Name = "frmIOApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Phase 1: UI scaffolding only. No HTTP, no Profibus yet.
' Start enables a heartbeat timer that logs "tick" to the listbox.
' Stop disables the timer. Lays the groundwork for Phases 2 (HTTP)
' and 3 (Profibus reads/writes).
'
' We use a VB.ListBox for the log instead of a multi-line VB.TextBox
' because multi-line TextBoxes require a companion .frx file in VB6.

Private Const MAX_LOG_LINES As Long = 1000

Private Sub Form_Load()
    AppendLog "I/O App started"
    lblStatusValue.Caption = "Stopped"
End Sub

Private Sub cmdStart_Click()
    AppendLog "Start clicked"
    Timer1.Enabled = True
    lblStatusValue.Caption = "Running"
End Sub

Private Sub cmdStop_Click()
    AppendLog "Stop clicked"
    Timer1.Enabled = False
    lblStatusValue.Caption = "Stopped"
End Sub

Private Sub Timer1_Timer()
    ' Phase 1 heartbeat. Phase 2 replaces this with HTTP traffic.
    AppendLog "tick"
End Sub

Private Sub AppendLog(msg As String)
    ' Avoid the name "Log" because VB6's built-in Log() is the natural logarithm.
    Dim line As String
    line = Format$(Now, "hh:nn:ss") & "  " & msg
    lstLog.AddItem line
    ' Scroll to bottom so the newest line is visible.
    If lstLog.ListCount > 0 Then
        lstLog.TopIndex = lstLog.ListCount - 1
    End If
    ' Cap memory: drop the oldest line once we exceed MAX_LOG_LINES.
    If lstLog.ListCount > MAX_LOG_LINES Then
        lstLog.RemoveItem 0
    End If
End Sub
