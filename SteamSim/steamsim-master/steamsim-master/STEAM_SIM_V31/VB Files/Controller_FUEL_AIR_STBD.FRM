VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{17FE8C5E-5183-11D3-AA4D-00C04FA979C0}#7.0#0"; "HorizontalGauge.ocx"
Object = "{237BBFE9-4F2C-11D3-AA49-00C04FA979C0}#10.0#0"; "EdgeGauge.ocx"
Begin VB.Form Controller_FUEL_AIR_STBD 
   Caption         =   "Fuel/Air Controller STBD"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3750
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4095
      Left            =   120
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7223
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   2800
         TabIndex        =   14
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   80
         TabIndex        =   13
         Top             =   3720
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   480
         Top             =   2400
      End
      Begin EdgeGauge.EdgeGaugectl EdgeGaugectl1 
         Height          =   2625
         Left            =   1320
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   4630
         GaugeScaleText  =   "PSI"
      End
      Begin HorizontalGauge.HGauge H_gauge 
         Height          =   555
         Left            =   1320
         TabIndex        =   12
         ToolTipText     =   "Balance Gauge"
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         MinorTicks      =   1
      End
      Begin VB.Label Balance_lbl 
         Alignment       =   2  'Center
         Caption         =   "100%"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Balance SP"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   3600
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3855
         Left            =   0
         Top             =   120
         Width           =   3135
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
         X1              =   240
         X2              =   360
         Y1              =   3240
         Y2              =   3120
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   240
         X2              =   360
         Y1              =   3240
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   960
         X2              =   1080
         Y1              =   3120
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   960
         X2              =   1080
         Y1              =   3360
         Y2              =   3240
      End
      Begin VB.Label DEC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
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
         Height          =   495
         Left            =   720
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
         Left            =   960
         TabIndex        =   5
         Top             =   1440
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
         Left            =   960
         TabIndex        =   4
         Top             =   960
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
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
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
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "Controller_FUEL_AIR_STBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INC_SP_Mousedown As Boolean
Dim DEC_SP_Mousedown As Boolean
Dim INC_MAN_Mousedown As Boolean
Dim DEC_MAN_Mousedown As Boolean


Private Sub AUTO_label_Click()

 ' Call set_PCDI_LSS( )

End Sub

Private Sub Command1_Click()

Controller_AIR_STBD.Show


End Sub

Private Sub Command2_Click()

Controller_MASTER_STBD.Show

End Sub

Private Sub DEC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DEC_MAN_Mousedown = True
    DEC_label.BorderStyle = 0

End Sub

Private Sub DEC_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 DEC_MAN_Mousedown = False
 DEC_label.BorderStyle = 1

End Sub

Private Sub DEC_SP_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  ' SP
    DEC_SP_Mousedown = True
    DEC_SP_label.BorderStyle = 0


End Sub

Private Sub DEC_SP_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


 DEC_SP_Mousedown = False
 DEC_SP_label.BorderStyle = 1


End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Edge_gauge_string(6) As String

On Error Resume Next
Call JYAScreenInit(Me, id_fuel_air_stbd)


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


'   Horizontal BALANCE Gauge
    H_gauge.GaugeText 1, "2/1"
    H_gauge.GaugeText 2, ""
    H_gauge.GaugeText 3, "1/1"
    H_gauge.GaugeText 4, ""
'   H_gauge.GaugeText 5, "8"
    H_gauge.GaugeText 5, "1/2"
'   H_gauge.GaugeText 6, "1/2"
    
    H_gauge.Value = 127  '  MIDDLE  ( Scaled 0 to 255 )

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub


Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_fuel_air_stbd)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub INC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


  ' BALANCE
  INC_MAN_Mousedown = True
  INC_label.BorderStyle = 0


End Sub

Private Sub INC_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 INC_MAN_Mousedown = False
 INC_label.BorderStyle = 1

End Sub

Private Sub INC_SP_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' SP
  INC_SP_Mousedown = True
  INC_SP_label.BorderStyle = 0

 
End Sub

Private Sub INC_SP_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 INC_SP_Mousedown = False
 INC_SP_label.BorderStyle = 1

End Sub

Private Sub MAN_label_Click()

'   Call set_PCDI_LSS(  )

End Sub

Private Sub SP_label_Click()

'  Call set_PCDI_LSS( )

End Sub

Private Sub Timer1_Timer()
Dim R_num As Single


'c   === FUEL/AIR RATIO ===
'
'    equivalence(CONTROLLER_VARS(51),SP_auto_fuel_air(1))
'    equivalence(CONTROLLER_VARS(53),Balance_fuel_air(1))
'    equivalence(CONTROLLER_VARS(55),mV_fuel_air(1))
'    equivalence(CONTROLLER_VARS(57),PID_out_fuel_air(1))
'    equivalence(CONTROLLER_VARS(59),VV_pos_fuel_air(1))
'
'    real*4   SP_auto_fuel_air(2),Balance_fuel_air(2),mV_fuel_air(2),PID_out_fuel_air(2),VV_pos_fuel_air(2)

'c   FUEL/AIR CONTROLLER STBD
'          PC_REALS(293) = dials(34)
'          PC_REALS(294) = dials(35)
'          PC_REALS(295) = dials(36)
'
'c   FUEL/AIR CONTROLLER PORT
'          PC_REALS(442) = dials(88)
'          PC_REALS(443) = dials(89)
'          PC_REALS(444) = dials(90)


' BALANCE CONTROL
     
If INC_MAN_Mousedown Then
    
      If (H_gauge.Value < 255) Then H_gauge.Value = H_gauge.Value + 1
      'H_gauge.Value = H_gauge.Value + 1
                
      Call setCONTROLLERvars(53, H_gauge.Value)     ' Balance_fuel_air(1)
                  
ElseIf DEC_MAN_Mousedown Then
    
      If (H_gauge.Value > 0) Then H_gauge.Value = H_gauge.Value - 1
      'H_gauge.Value = H_gauge.Value - 1
          
      Call setCONTROLLERvars(53, H_gauge.Value)       ' Balance_fuel_air(1)
Else
      H_gauge.Value = Format(LSS_reals(295), "##0")     ' BOTTOM Gauge
End If

' === 18/12/01 ===
Balance_lbl.Caption = Format(LSS_reals(295) * 100# / 255#, "##0.0") & "%"
'R_num = (1# - LSS_reals(295) / 255#)
'Balance_lbl.Caption = Format((R_num * 1.5 + 0.5), "##0.00")
     

' AUTO CONTROL
       
If INC_SP_Mousedown Then
    
      EdgeGaugectl1.Value = EdgeGaugectl1.Value + 1
          
      Call setCONTROLLERvars(51, EdgeGaugectl1.Value)   ' SP_auto_fuel_air(1)
         
ElseIf DEC_SP_Mousedown Then

      EdgeGaugectl1.Value = EdgeGaugectl1.Value - 1
          
      Call setCONTROLLERvars(51, EdgeGaugectl1.Value)   ' SP_auto_fuel_air(1)
          
Else
      EdgeGaugectl1.Value = Format(LSS_reals(293), "##0")   ' LEFT Gauge
End If
   


'c   FUEL/AIR CONTROLLER STBD
'          PC_REALS(293) = dials(34)
'          PC_REALS(294) = dials(35)
'          PC_REALS(295) = dials(36)
'
'c   FUEL/AIR CONTROLLER PORT
'          PC_REALS(442) = dials(88)
'          PC_REALS(443) = dials(89)
'          PC_REALS(444) = dials(90)



'   Measured Value on the RIGHT HAND GAUGE
          EdgeGaugectl2.Value = LSS_reals(294)   ' RIGHT Gauge
    


End Sub

Private Sub VLV_label_Click()

'   Call set_PCDI_LSS( )

End Sub
