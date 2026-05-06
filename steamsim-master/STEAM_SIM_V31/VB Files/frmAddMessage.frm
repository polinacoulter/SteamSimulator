VERSION 5.00
Begin VB.Form frmAddMessage 
   Caption         =   "Add Message To Log"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1050
   ScaleWidth      =   5100
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Message"
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Height          =   465
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   4995
   End
End
Attribute VB_Name = "frmAddMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
            Call AddToLog(txtMessage.Text & " " & Time)
End Sub

Private Sub Form_Load()
    Me.Width = ScreenSettings(idAddMessage).Width
    Me.Height = ScreenSettings(idAddMessage).Height
End Sub

Private Sub Form_Resize()
    If Me.Height < 1500 Then
        Me.Height = 1500
    Else
        txtMessage.Width = Me.ScaleWidth - txtMessage.Left - 50
        txtMessage.Height = Me.ScaleHeight - cmdAdd.Height - 50 - txtMessage.Top
        cmdAdd.Top = Me.ScaleHeight - cmdAdd.Height - 5
        cmdAdd.Left = Me.ScaleWidth / 2 - cmdAdd.Width / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ScreenSettings(idAddMessage).Width = Me.Width
    ScreenSettings(idAddMessage).Height = Me.Height
End Sub
