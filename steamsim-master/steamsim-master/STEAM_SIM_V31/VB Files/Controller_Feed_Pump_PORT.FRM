VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{237BBFE9-4F2C-11D3-AA49-00C04FA979C0}#10.0#0"; "EdgeGauge.ocx"
Begin VB.Form Controller_FEED_PUMP_PORT 
   Caption         =   "Feed Pump PORT"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4725
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4395
      Left            =   120
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7752
      Begin VB.CommandButton Command4 
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   300
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Settings"
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   2720
         TabIndex        =   13
         Top             =   3750
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   80
         TabIndex        =   12
         Top             =   3750
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   480
         Top             =   2400
      End
      Begin EdgeGauge.EdgeGaugectl EdgeGaugectl1 
         Height          =   2625
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   4630
         Value           =   127
         GaugeScaleText  =   "PSI"
      End
      Begin EdgeGauge.EdgeGaugectl EdgeGaugectl2 
         Height          =   2625
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   4630
         GaugeScaleText  =   "PSI"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Balance"
         Height          =   255
         Left            =   -180
         TabIndex        =   16
         Top             =   2100
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Manual"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   3480
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4275
         Left            =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   600
         X2              =   720
         Y1              =   1080
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   480
         X2              =   600
         Y1              =   1200
         Y2              =   1080
      End
      Begin VB.Label INC_SP_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   600
         X2              =   720
         Y1              =   1680
         Y2              =   1560
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   480
         X2              =   600
         Y1              =   1560
         Y2              =   1680
      End
      Begin VB.Label DEC_SP_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Set Point"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   540
         Width           =   735
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   1800
         X2              =   1920
         Y1              =   3240
         Y2              =   3120
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   1680
         X2              =   1800
         Y1              =   3120
         Y2              =   3240
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   2400
         X2              =   2520
         Y1              =   3120
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   2280
         X2              =   2400
         Y1              =   3240
         Y2              =   3120
      End
      Begin VB.Label DEC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label INC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label VLV_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VLV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label SP_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label AUTO_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label MAN_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAN  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   495
      End
   End
End
Attribute VB_Name = "Controller_FEED_PUMP_PORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INC_SP_Mousedown As Boolean
Dim DEC_SP_Mousedown As Boolean
Dim INC_MAN_Mousedown As Boolean
Dim DEC_MAN_Mousedown As Boolean
' 29/11/99
Dim SP_manual As Integer
Dim SP_automatic As Integer


Private Sub AUTO_label_Click()

  Call setPCDI(132)

End Sub

Private Sub Command1_Click()

Controller_Drum_level_PORT.Show


End Sub

Private Sub Command2_Click()

Controller_Temp_PORT.Show


End Sub

Private Sub Command3_Click()


' 12/12/99
'c   === FEED PUMP   ( PID LOOPNUM 12 and 13 ) ===
'
'    equivalence(CONTROLLER_VARS(191),Td_feed_pump(1))
'    equivalence(CONTROLLER_VARS(193),Ti_feed_pump(1))
'    equivalence(CONTROLLER_VARS(195),PBand_feed_pump(1))
'    equivalence(CONTROLLER_VARS(197),C_feed_pump(1))
'
'    real*4              Td_feed_pump(2),Ti_feed_pump(2),PBand_feed_pump(2),C_feed_pump(2)

n_Ti = 194
n_Pband = 196

nLOOPNUMBER = 13
Call set_CONTROLLER_PID_constants(0, 0#, nLOOPNUMBER)

PID_constants.Show

End Sub

Private Sub Command4_Click()

' 9/1/00
Controller_feed_pump_Balance.Show

End Sub

Private Sub DEC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


' 29/11/99  MAN = 1 , SP = 3
'If VB_steam_integers(44) = 1 Then
If (VB_steam_integers(44) = 1 Or VB_steam_integers(44) = 3) Then
    DEC_MAN_Mousedown = True
    DEC_label.BorderStyle = 0
End If

End Sub

Private Sub DEC_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 DEC_MAN_Mousedown = False
 DEC_label.BorderStyle = 1

End Sub

Private Sub DEC_SP_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 29/11/99  AUTO = 2  , VLV = 4
'If VB_steam_integers(44) = 2 Then
If (VB_steam_integers(44) = 2 Or VB_steam_integers(44) = 4) Then
    DEC_SP_Mousedown = True
    DEC_SP_label.BorderStyle = 0
End If


End Sub

Private Sub DEC_SP_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


 DEC_SP_Mousedown = False
 DEC_SP_label.BorderStyle = 1


End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Edge_gauge_string(6) As String

On Error Resume Next
Call JYAScreenInit(Me, id_feed_pump_stbd)


EdgeGaugectl1.MajorTicks = 6
EdgeGaugectl1.MinorTicks = 4
Edge_gauge_string(0) = "100"
Edge_gauge_string(1) = "80"
Edge_gauge_string(2) = "60"
Edge_gauge_string(3) = "40"
Edge_gauge_string(4) = "20"
Edge_gauge_string(5) = " 0"
EdgeGaugectl1.GaugeText Edge_gauge_string
EdgeGaugectl1.GaugeScaleText = ""

EdgeGaugectl2.MajorTicks = 6
EdgeGaugectl2.MinorTicks = 4
EdgeGaugectl2.GaugeText Edge_gauge_string
EdgeGaugectl2.GaugeScaleText = ""

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub


Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_feed_pump_stbd)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub INC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


'    equivalence(STEAM_integers(43),CONTROLLER_FEED_PUMP_STBD)
'    integer*2            CONTROLLER_FEED_PUMP_STBD
'    equivalence(STEAM_integers(44),CONTROLLER_FEED_PUMP_PORT)
'    integer*2            CONTROLLER_FEED_PUMP_PORT



' 29/11/99  MAN = 1 , SP = 3
'If VB_steam_integers(44) = 1 Then
If (VB_steam_integers(44) = 1 Or VB_steam_integers(44) = 3) Then
  INC_MAN_Mousedown = True
  INC_label.BorderStyle = 0
End If

End Sub

Private Sub INC_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 INC_MAN_Mousedown = False
 INC_label.BorderStyle = 1

End Sub

Private Sub INC_SP_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
' 29/11/99  AUTO = 2  , VLV = 4
'If VB_steam_integers(44) = 2 Then
If (VB_steam_integers(44) = 2 Or VB_steam_integers(44) = 4) Then
  INC_SP_Mousedown = True
  INC_SP_label.BorderStyle = 0
End If

 
End Sub

Private Sub INC_SP_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 INC_SP_Mousedown = False
 INC_SP_label.BorderStyle = 1

End Sub

Private Sub MAN_label_Click()

   Call setPCDI(131)

End Sub

Private Sub SP_label_Click()

'  Call setpcdi( )

End Sub

Private Sub Timer1_Timer()


'    equivalence(STEAM_integers(43),CONTROLLER_FEED_PUMP_STBD)
'    integer*2            CONTROLLER_FEED_PUMP_STBD
'    equivalence(STEAM_integers(44),CONTROLLER_FEED_PUMP_PORT)
'    integer*2            CONTROLLER_FEED_PUMP_PORT


'c   FEED PUMP STBD
'          PC_REALS(445) = dials(26)
'          PC_REALS(446) = dials(27)
'
'c   FEED PUMP PORT
'          PC_REALS(447) = dials(78)
'          PC_REALS(448) = dials(79)

'c   === FEED PUMP ===
'
'    equivalence(CONTROLLER_VARS(71),SP_auto_feed_pump(1))
'    equivalence(CONTROLLER_VARS(73),SP_man_feed_pump(1))
'    equivalence(CONTROLLER_VARS(75),mV_feed_pump(1))
'    equivalence(CONTROLLER_VARS(77),PID_out_feed_pump(1))
'    equivalence(CONTROLLER_VARS(79),VV_pos_feed_pump(1))
'
'    real*4   SP_auto_feed_pump(2),SP_man_feed_pump(2),mV_feed_pump(2),PID_out_feed_pump(2),VV_pos_feed_pump(2)


If VB_steam_integers(44) = 1 Then

   ' MANUAL
     MAN_label.BackColor = &HFF00&      ' BRIGHT GREEN
    AUTO_label.BackColor = &H8000&      ' DARK GREEN
      SP_label.BackColor = &H8000&      ' DARK GREEN
     VLV_label.BackColor = &H8000&      ' DARK GREEN
     
     ' 29/11/99
    'If INC_MAN_Mousedown Then
    '      EdgeGaugectl1.Value = EdgeGaugectl1.Value + 1
    '
    '      Call setCONTROLLERvars(74, EdgeGaugectl1.Value)     'SP_man_feed_pump(2)
    '
    'ElseIf DEC_MAN_Mousedown Then
    '      EdgeGaugectl1.Value = EdgeGaugectl1.Value - 1
    '
    '       Call setCONTROLLERvars(74, EdgeGaugectl1.Value)     'SP_man_feed_pump(2)
    'Else
    '      EdgeGaugectl1.Value = Format(LSS_reals(447), "##0")   ' LEFT Gauge
    'End If

     
ElseIf VB_steam_integers(44) = 2 Then

   ' AUTO
     MAN_label.BackColor = &H8000&      ' DARK GREEN
    AUTO_label.BackColor = &HFF00&      ' BRIGHT GREEN
      SP_label.BackColor = &H8000&      ' DARK GREEN
     VLV_label.BackColor = &H8000&      ' DARK GREEN
     
     ' 29/11/99
     '     EdgeGaugectl1.Value = Format(LSS_reals(447), "##0")   ' LEFT Gauge
     
ElseIf VB_steam_integers(44) = 3 Then

   ' SP
     MAN_label.BackColor = &H8000&      ' DARK GREEN
    AUTO_label.BackColor = &H8000&      ' DARK GREEN
      SP_label.BackColor = &HFF00&      ' BRIGHT GREEN
     VLV_label.BackColor = &H8000&       ' DARK GREEN
     
     ' 29/11/99
    'If INC_SP_Mousedown Then
    '      EdgeGaugectl1.Value = EdgeGaugectl1.Value + 1
    '
    '       Call setCONTROLLERvars(72, EdgeGaugectl1.Value)   ' SP_auto_feed_pump(2)
    '
    'ElseIf DEC_SP_Mousedown Then
    '      EdgeGaugectl1.Value = EdgeGaugectl1.Value - 1
    '
    '      Call setCONTROLLERvars(72, EdgeGaugectl1.Value)   ' SP_auto_feed_pump(2)
    '
    'Else
    '      EdgeGaugectl1.Value = Format(LSS_reals(447), "##0")   ' LEFT Gauge
    'End If
   
ElseIf VB_steam_integers(44) = 4 Then

   ' VALVE
     MAN_label.BackColor = &H8000&      ' DARK GREEN
    AUTO_label.BackColor = &H8000&      ' DARK GREEN
      SP_label.BackColor = &H8000&      ' DARK GREEN
     VLV_label.BackColor = &HFF00&      ' BRIGHT GREEN
     
      ' 29/11/99
      '    EdgeGaugectl1.Value = Format(LSS_reals(447), "##0")  ' LEFT Gauge
     
End If

' 29/11/99
Call set_controller(44, 74, 72, 548, 546, 447)

'        PC_REALS(545) = SP_auto_feed_pump(1) * 255#
'        PC_REALS(546) = SP_auto_feed_pump(2) * 255#
'        PC_REALS(547) = SP_man_feed_pump(1) * 255#
'        PC_REALS(548) = SP_man_feed_pump(2) * 255#


'c   FEED PUMP STBD
'          PC_REALS(445) = dials(26)
'          PC_REALS(446) = dials(27)
'
'c   FEED PUMP PORT
'          PC_REALS(447) = dials(78)
'          PC_REALS(448) = dials(79)



'   Measured Value on the RIGHT HAND GAUGE
          EdgeGaugectl2.Value = LSS_reals(448)   ' RIGHT Gauge
    


End Sub

Private Sub VLV_label_Click()

'   Call setpcdi( )

End Sub
Sub set_controller(nINTEGERS As Integer, c_man As Integer, c_auto As Integer, _
R_man As Integer, R_auto As Integer, Left_gauge As Integer)

' 28/11/99          1 = MAN   ,   3 = SP
If VB_steam_integers(nINTEGERS) = 1 Or VB_steam_integers(nINTEGERS) = 3 Then

    If INC_MAN_Mousedown Then
    
            SP_manual = SP_manual + 1
            Call setCONTROLLERvars(c_man, SP_manual)
                     
    ElseIf DEC_MAN_Mousedown Then
            
            SP_manual = SP_manual - 1
            Call setCONTROLLERvars(c_man, SP_manual)
   Else
            SP_manual = LSS_reals(R_man)
    End If

End If
' 28/11/99          2 = AUTO   ,   4 = VLV
If VB_steam_integers(nINTEGERS) = 2 Or VB_steam_integers(nINTEGERS) = 4 Then

     If INC_SP_Mousedown Then
     
            SP_automatic = SP_automatic + 1
            Call setCONTROLLERvars(c_auto, SP_automatic)
         
    ElseIf DEC_SP_Mousedown Then
    
            SP_automatic = SP_automatic - 1
          Call setCONTROLLERvars(c_auto, SP_automatic)
    Else
            SP_automatic = LSS_reals(R_auto)
    End If
End If

          EdgeGaugectl1.Value = Format(LSS_reals(Left_gauge), "##0")   ' LEFT Gauge


End Sub

