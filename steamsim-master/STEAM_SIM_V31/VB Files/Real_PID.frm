VERSION 5.00
Begin VB.Form Real_PID 
   Caption         =   "Real PID Selection"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Caption         =   "Real PID Control Selection"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   5880
         Top             =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Instructions:  Click the System required to be controlled by the Real (Hardware) PID Controller"
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   180
         TabIndex        =   11
         Top             =   2220
         Width           =   5415
      End
      Begin VB.Shape PID_lamp 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1620
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1020
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   720
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   4020
         Shape           =   3  'Circle
         Top             =   1620
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   7
         Left            =   4020
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   4020
         Shape           =   3  'Circle
         Top             =   1020
         Width           =   195
      End
      Begin VB.Shape PID_lamp 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   4020
         Shape           =   3  'Circle
         Top             =   720
         Width           =   195
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "NONE Selected"
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   10
         Top             =   1620
         Width           =   1935
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "SW Temperature"
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   9
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "LTCWTemperature"
         Height          =   255
         Index           =   6
         Left            =   4020
         TabIndex        =   8
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "HTCWTemperature"
         Height          =   255
         Index           =   5
         Left            =   4020
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "NONE Selected"
         Height          =   255
         Index           =   4
         Left            =   900
         TabIndex        =   6
         Top             =   1620
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "FO Viscosity"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "Scavenge AirTemperature"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   2355
      End
      Begin VB.Label PID_label 
         Alignment       =   2  'Center
         Caption         =   "MELO Temperature"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Real PID"
         Height          =   195
         Left            =   2940
         TabIndex        =   2
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "123.4"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2940
         TabIndex        =   1
         Top             =   840
         Width           =   675
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   2880
         Shape           =   4  'Rounded Rectangle
         Top             =   780
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   2760
         Top             =   660
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Real_PID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)


End Sub

Private Sub Form_Load()


Me.Height = 3225
Me.Width = 7080


End Sub

Private Sub PID_label_Click(Index As Integer)
Dim i As Integer

If In_a_group Then

' Select REAL PID
'  Call setPCDI(RMAIO(1,nRMA)     nRMA=1 for MELUB RMA
   Call setPCDI(RMAIO(1, Index))
   
 ' De-Select the "other" REAL PIDs
    For i = 1 To 8
      If Index <> i Then
          Call setPCDI(RMAIO(1, i) + 1)
      End If
    Next i
Else
               response = MsgBox("OFF LINE. Can not select the REAL PID", vbOK, "Information")
End If



End Sub

Private Sub Timer1_Timer()
Dim i As Integer

For i = 1 To 8
    ' If rma_status(nRMA) = 4 Then
      If rma_status(i) = 4 Then
         PID_lamp(i).FillColor = &HFFFFFF      ' WHITE
      Else
         PID_lamp(i).FillColor = &HC0C0C0      ' GREY
      End If
Next i

End Sub
