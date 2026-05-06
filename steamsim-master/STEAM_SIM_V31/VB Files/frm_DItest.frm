VERSION 5.00
Begin VB.Form frm_DItest 
   Caption         =   "Digital Input Test"
   ClientHeight    =   8460
   ClientLeft      =   2415
   ClientTop       =   1785
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8460
   ScaleWidth      =   6690
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   5055
      Begin VB.CommandButton ExitCommand 
         Caption         =   "OK"
         Height          =   495
         Left            =   3720
         TabIndex        =   2
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   600
         Top             =   5160
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1200
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frm_DItest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExitCommand_Click()
Unload Me
End Sub

Private Sub Form_Load()
For i% = 136 To 759
    
        DIold(i%) = DI(i%)
  
Next
End Sub

Private Sub Timer1_Timer()
For i% = 136 To 759
    If DIold(i%) <> DI(i%) Then
        Text1.Text = "DI(" & i% & ") set " & DI(i%)
        DIold(i%) = DI(i%)
    End If
Next
End Sub


