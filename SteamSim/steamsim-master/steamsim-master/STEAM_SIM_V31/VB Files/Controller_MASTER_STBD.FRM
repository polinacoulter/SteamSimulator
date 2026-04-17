VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{237BBFE9-4F2C-11D3-AA49-00C04FA979C0}#10.0#0"; "EdgeGauge.ocx"
Begin VB.Form Controller_MASTER_STBD 
   Caption         =   "MASTER Controller STBD"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3795
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4275
      Left            =   120
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   7541
      Begin VB.CommandButton Controller_command 
         Height          =   195
         Left            =   540
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Low Fire Setting"
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   3780
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Settings"
         Height          =   315
         Left            =   1740
         TabIndex        =   14
         Top             =   3780
         Width           =   915
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
         Height          =   4215
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
         Top             =   1920
         Visible         =   0   'False
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
Attribute VB_Name = "Controller_MASTER_STBD"
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

  ' === 11/11/02 ===
  'Call set_PCDI_LSS(73)
   Call setPCDI(73)

End Sub

Private Sub Command1_Click()

Controller_FUEL_AIR_STBD.Show

End Sub

Private Sub Command2_Click()

Controller_PLANT_MASTER.Show

End Sub

Private Sub Command3_Click()

' 9/10/99  Select the  MASTER  STBD PID

' See set_CONTROLLER_PID_constants.F90
n_Ti = 155
n_Pband = 157

nLOOPNUMBER = 9
Call set_CONTROLLER_PID_constants(0, 0#, nLOOPNUMBER)

PID_constants.Show


End Sub

Private Sub Command4_Click()


' 7/4/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 254

Alarm_set_point.Show

End Sub


Private Sub Controller_command_Click()

' === 25/09/02 ===
Controller_Stbd.ZOrder

End Sub

Private Sub DEC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 29/11/99  MAN = 1 , SP = 3
'If VB_steam_integers(41) = 1 Then
If (VB_steam_integers(41) = 1 Or VB_steam_integers(41) = 3) Then
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
'If VB_steam_integers(41) = 2 Then
If (VB_steam_integers(41) = 2 Or VB_steam_integers(41) = 4) Then
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
Call JYAScreenInit(Me, id_master_stbd)


' === 25/09/02  UNDER DEVELOPMENT... ===
'If I_AM_HOST Then
'    Controller_command.Visible = True
'Else
'    Controller_command.Visible = False
'End If



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
    Call JYAScreenResize(Me, id_master_stbd)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub INC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


'    equivalence(STEAM_integers(41),CONTROLLER_MASTER_STBD)
'    integer*2            CONTROLLER_MASTER_STBD
'    equivalence(STEAM_integers(42),CONTROLLER_MASTER_PORT)
'    integer*2            CONTROLLER_MASTER_PORT

' 29/11/99  MAN = 1 , SP = 3
'If VB_steam_integers(41) = 1 Then
If (VB_steam_integers(41) = 1 Or VB_steam_integers(41) = 3) Then
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
'If VB_steam_integers(41) = 2 Then
If (VB_steam_integers(41) = 2 Or VB_steam_integers(41) = 4) Then
  INC_SP_Mousedown = True
  INC_SP_label.BorderStyle = 0
End If

 
End Sub

Private Sub INC_SP_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 INC_SP_Mousedown = False
 INC_SP_label.BorderStyle = 1

End Sub

Private Sub MAN_label_Click()

  ' === 11/11/02 ===
  ' Call set_PCDI_LSS(72)
    Call setPCDI(72)

End Sub

Private Sub SP_label_Click()

'  Call set_PCDI_LSS( )

End Sub

Private Sub Timer1_Timer()


'    equivalence(CONTROLLER_VARS(41),SP_auto_MASTER(1))
'    equivalence(CONTROLLER_VARS(43),SP_man_MASTER(1))
'    equivalence(CONTROLLER_VARS(45),mV_MASTER(1))
'    equivalence(CONTROLLER_VARS(47),PID_out_MASTER(1))
'    equivalence(CONTROLLER_VARS(49),VV_pos_MASTER(1))
'
'    real*4   SP_auto_MASTER(2),SP_man_MASTER(2),mV_MASTER(2),PID_out_MASTER(2),VV_pos_MASTER(2)


'c   MASTER Controller STBD
'          PC_REALS(286) = dials(37)
'          PC_REALS(287) = dials(38)
'
'
'c   MASTER Controller PORT
'          PC_REALS(288) = dials(86)
'          PC_REALS(289) = dials(87)


'    equivalence(STEAM_integers(41),CONTROLLER_MASTER_STBD)
'    integer*2            CONTROLLER_MASTER_STBD
'    equivalence(STEAM_integers(42),CONTROLLER_MASTER_PORT)
'    integer*2            CONTROLLER_MASTER_PORT

If VB_steam_integers(41) = 1 Then

   ' MANUAL
     MAN_label.BackColor = &HFF00&      ' BRIGHT GREEN
    AUTO_label.BackColor = &H8000&      ' DARK GREEN
      SP_label.BackColor = &H8000&      ' DARK GREEN
     VLV_label.BackColor = &H8000&      ' DARK GREEN
     
     
    ' 29/11/99
   ' If INC_MAN_Mousedown Then
   '       EdgeGaugectl1.Value = EdgeGaugectl1.Value + 1
   '
   '       Call setCONTROLLERvars(43, EdgeGaugectl1.Value)     'SP_man_MASTER(1)
   '
   ' ElseIf DEC_MAN_Mousedown Then
   '       EdgeGaugectl1.Value = EdgeGaugectl1.Value - 1
   '
   '        Call setCONTROLLERvars(43, EdgeGaugectl1.Value)     'SP_man_MASTER(1)
   ' Else
   '       EdgeGaugectl1.Value = Format(LSS_reals(286), "##0")   ' LEFT Gauge
   ' End If

     
ElseIf VB_steam_integers(41) = 2 Then

   ' AUTO
     MAN_label.BackColor = &H8000&      ' DARK GREEN
    AUTO_label.BackColor = &HFF00&      ' BRIGHT GREEN
      SP_label.BackColor = &H8000&      ' DARK GREEN
     VLV_label.BackColor = &H8000&      ' DARK GREEN
     
     ' 29/11/99
       '   EdgeGaugectl1.Value = Format(LSS_reals(286), "##0")   ' LEFT Gauge
     
ElseIf VB_steam_integers(41) = 3 Then

   ' SP
     MAN_label.BackColor = &H8000&      ' DARK GREEN
    AUTO_label.BackColor = &H8000&      ' DARK GREEN
      SP_label.BackColor = &HFF00&      ' BRIGHT GREEN
     VLV_label.BackColor = &H8000&       ' DARK GREEN
     
     ' 29/11/99
   ' If INC_SP_Mousedown Then
   '       EdgeGaugectl1.Value = EdgeGaugectl1.Value + 1
   '
   '        Call setCONTROLLERvars(41, EdgeGaugectl1.Value)   ' SP_auto_MASTER(1)
   '
   ' ElseIf DEC_SP_Mousedown Then
   '       EdgeGaugectl1.Value = EdgeGaugectl1.Value - 1
   '
    '      Call setCONTROLLERvars(41, EdgeGaugectl1.Value)   ' SP_auto_MASTER(1)
    '
    'Else
   '       EdgeGaugectl1.Value = Format(LSS_reals(286), "##0")   ' LEFT Gauge
   ' End If
   
ElseIf VB_steam_integers(41) = 4 Then

   ' VALVE
     MAN_label.BackColor = &H8000&      ' DARK GREEN
    AUTO_label.BackColor = &H8000&      ' DARK GREEN
      SP_label.BackColor = &H8000&      ' DARK GREEN
     VLV_label.BackColor = &HFF00&      ' BRIGHT GREEN
     
      ' 29/11/99
       '   EdgeGaugectl1.Value = Format(LSS_reals(286), "##0")  ' LEFT Gauge
     
End If

' 29/11/99
Call set_controller(41, 43, 41, 537, 535, 286)

'        PC_REALS(535) = SP_auto_MASTER(1) * 255#
'        PC_REALS(536) = SP_auto_MASTER(2) * 255#
'        PC_REALS(537) = SP_man_MASTER(1) * 255#
'        PC_REALS(538) = SP_man_MASTER(2) * 255#


'c   MASTER Controller STBD
'          PC_REALS(286) = dials(37)
'          PC_REALS(287) = dials(38)
'
'
'c   MASTER Controller PORT
'          PC_REALS(288) = dials(86)
'          PC_REALS(289) = dials(87)



'   Measured Value on the RIGHT HAND GAUGE
          EdgeGaugectl2.Value = LSS_reals(287)   ' RIGHT Gauge
    


End Sub

Private Sub VLV_label_Click()

'   Call set_PCDI_LSS( )

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

