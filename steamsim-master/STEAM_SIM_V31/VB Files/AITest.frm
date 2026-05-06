VERSION 5.00
Begin VB.Form frm_AITest 
   Caption         =   "Analog Input Test"
   ClientHeight    =   8265
   ClientLeft      =   2685
   ClientTop       =   2250
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8265
   ScaleWidth      =   6345
   Begin VB.CommandButton ExitCommand 
      Caption         =   "OK"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6015
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   3720
         Top             =   5520
      End
      Begin VB.ListBox List1 
         Height          =   5520
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frm_AITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExitCommand_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
List1.Clear
For i% = 8 To 31
    List1.AddItem "AI(" & i% & ") = " & AI(i%)
    'List1.ListIndex = i%
    'List1.Text(List1.ListIndex) = "AI(" & i% & ") = " & AI(i%)
Next
End Sub


