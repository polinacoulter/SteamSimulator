VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form TEST_SCOPE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Synchro-Scope and Synchro-Lamp Test"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   10470
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4455
      Left            =   60
      Top             =   60
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7858
      Begin VB.CommandButton Command1 
         Caption         =   "LOCK RELAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PARK ALL RELAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "PARK SCOPE RELAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "TOP LAMP RELAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "LEFT LAMP RELAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "RIGHT LAMP RELAY"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   3120
         Width           =   2175
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1335
         Left            =   120
         Max             =   100
         TabIndex        =   3
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   120
         Top             =   120
      End
      Begin VB.ListBox Scope_list 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3660
         Left            =   4080
         TabIndex        =   1
         Top             =   0
         Width           =   6135
      End
      Begin VB.CommandButton Command7 
         Caption         =   "SCOPE ON/OFF"
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Frequency Difference"
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Inverted LOGIC  >>>>>"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Inverted LOGIC  >>>>>"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1280
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Synchro-Scope Phase Difference ( 0 to 255 )"
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Label Phase_diff 
         Caption         =   "Label5"
         Height          =   375
         Left            =   7560
         TabIndex        =   10
         Top             =   3840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "TEST_SCOPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Scope_controls(2) = Not Scope_controls(2)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Command2_Click()

Scope_controls(3) = Not Scope_controls(3)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Command3_Click()

Scope_controls(4) = Not Scope_controls(4)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Command4_Click()

Scope_controls(5) = Not Scope_controls(5)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Command5_Click()

Scope_controls(6) = Not Scope_controls(6)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Command6_Click()

Scope_controls(7) = Not Scope_controls(7)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Command7_Click()

'VB_DO(2311) = Not VB_DO(2311)
'Call set_DO(VB_DO(0))

Scope_controls(1) = Not Scope_controls(1)
Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Form_Load()


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

RunIO = True

Scope_controls(1) = True     ' Put the Synchro-scope ON
Scope_controls(2) = True     ' Put the LOCK RELAY ON     (inverted logic)
Scope_controls(3) = True     ' Put the PARKALL RELAY ON  (inverted logic)
Scope_controls(4) = False    ' Put the PARKSCOPE RELAY OFF

Scope_controls(5) = False    ' Do NOT fail the   TOP  LAMP
Scope_controls(6) = False    ' Do NOT fail the  LEFT  LAMP
Scope_controls(7) = False    ' Do NOT fail the RIGHT  LAMP

Call TEST_SSCOPE(Fdiff, Scope_controls(1))

End Sub

Private Sub Timer1_Timer()
Dim scope_strings(7) As String
Dim i As Integer

scope_strings(1) = "SCOPE ON/OFF           = "
scope_strings(2) = "LOCK       RELAY       = "
scope_strings(3) = "PARK ALL   RELAY       = "
scope_strings(4) = "PARK SCOPE RELAY       = "
scope_strings(5) = "TOP   LAMP RELAY       = "    ' (A)
scope_strings(6) = "LEFT  LAMP RELAY       = "    ' (B)
scope_strings(7) = "RIGHT LAMP RELAY       = "    ' (C)

Text1.text = "Frequency difference = " & Format(Fdiff, "##0.00")

Scope_list.clear

For i = 1 To 7
   Scope_list.AddItem " "
  If Scope_controls(i) Then
   Scope_list.AddItem scope_strings(i) & "ON   (Energised)"
  Else
   Scope_list.AddItem scope_strings(i) & "OFF  (De-energised)"
  End If
Next i

' 24/11/99
Phase_diff.Caption = "AI(47) = " & Format(A_INPUT(47), "##0")


End Sub

Private Sub VScroll1_Change()

Fdiff = (VScroll1.Value - 50) / 50#
Call TEST_SSCOPE(-Fdiff, Scope_controls(1))

End Sub
