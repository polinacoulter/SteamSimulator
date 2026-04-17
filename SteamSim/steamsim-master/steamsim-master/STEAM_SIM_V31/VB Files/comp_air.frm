VERSION 5.00
Begin VB.Form comp_air 
   Caption         =   "Compressed Air"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6075
   Begin VB.CommandButton Command5 
      Caption         =   "CA08 valve"
      Height          =   315
      Left            =   2220
      TabIndex        =   9
      Top             =   1440
      Width           =   2835
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   2820
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PCDI(724) =  MSTARTSW(1)"
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   2700
      Width           =   2835
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1800
      Width           =   5835
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PCDI(19) = MELO#1 Start/Stop"
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   2160
      Width           =   2835
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5280
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   420
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PCDI(33) = Comp#1 ON/OFF"
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PCDI(161) = CA08 valve"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   900
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "comp_air"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   
  Call setPCDI(161)
   
End Sub

Private Sub Command2_Click()

  Call setPCDI(33)
   
End Sub

Private Sub Command3_Click()

Call setPCDI(19)

End Sub

Private Sub Command4_Click()

Call setPCDI(724)   'MSTARTSW(1)

End Sub

Private Sub Command5_Click()

'  Valve CA08 is valve 60
 ' Call setPCDI(161)
  Call setPCDI(DI_VIO(60))
  
End Sub

Private Sub Timer1_Timer()


Text1.Text = PCA(1)
'Text2.Text = VB_DO(2108) & " " & VB_DO(2109) & " " & VB_DO(749)

' 28/1/98
'Text3.Text = "MELO#1 p/p " & VB_DO(657) & " " & VB_DO(11) & "      MSTARTLP(1) =" & VB_DO(1129) & "      MDA_1_ACB_closedLP = " & VB_DO(1093)
Text3.Text = "MSTARTLP(1) =  " & VB_lamps(8) & "       MDA_1_ACB_closedLP = " & VB_lamps(9)

Text5.Text = "PML10 = " & Format(PML(4), "#0.0")


End Sub
