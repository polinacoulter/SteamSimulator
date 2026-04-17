VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{17FE8C5E-5183-11D3-AA4D-00C04FA979C0}#7.0#0"; "HorizontalGauge.ocx"
Object = "{237BBFE9-4F2C-11D3-AA49-00C04FA979C0}#10.0#0"; "EdgeGauge.ocx"
Begin VB.Form Controller_PLANT_MASTER 
   Caption         =   "PLANT MASTER Controller"
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
      Height          =   4275
      Left            =   120
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7541
      Begin VB.CommandButton Command4 
         Caption         =   "Status"
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   3840
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Settings"
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   3840
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   2820
         TabIndex        =   14
         Top             =   3740
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   80
         TabIndex        =   13
         Top             =   3740
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   420
         Top             =   2280
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
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Balance_lbl 
         Alignment       =   2  'Center
         Caption         =   "100%"
         Height          =   255
         Left            =   1860
         TabIndex        =   16
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Balance"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4095
         Left            =   20
         Top             =   120
         Width           =   3135
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         X1              =   600
         X2              =   720
         Y1              =   1080
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
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
         Width           =   495
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   600
         X2              =   720
         Y1              =   1680
         Y2              =   1560
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
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
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Set Point"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1920
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
Attribute VB_Name = "Controller_PLANT_MASTER"
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

Controller_MASTER_STBD.Show


End Sub

Private Sub Command2_Click()

Controller_MASTER_PORT.Show

End Sub

Private Sub Command3_Click()


' 8/10/99  Select the PLANT MASTER PID

n_Ti = 172
n_Pband = 173

nLOOPNUMBER = 11
Call set_CONTROLLER_PID_constants(0, 0#, nLOOPNUMBER)

PID_constants.Show


End Sub

Private Sub Command4_Click()


' === 24/01/02 ===
Controller_Status.ZOrder


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
Call JYAScreenInit(Me, id_plant_master)


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
    H_gauge.GaugeText 1, "Stbd"
    H_gauge.GaugeText 2, ""
    H_gauge.GaugeText 3, "Both"
    H_gauge.GaugeText 4, ""
'   H_gauge.GaugeText 5, "8"
    H_gauge.GaugeText 5, "Port"
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
    Call JYAScreenResize(Me, id_plant_master)
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


'c   === PLANT MASTER ===
'
'    equivalence(CONTROLLER_VARS(61),SP_auto_plant_master)
'    equivalence(CONTROLLER_VARS(62),Balance_plant_master)
'    equivalence(CONTROLLER_VARS(63),mV_plant_master)
'    equivalence(CONTROLLER_VARS(64),PID_out_plant_master)
'    equivalence(CONTROLLER_VARS(65),VV_pos_plant_master)
'
'    real*4   SP_auto_plant_master,Balance_plant_master,mV_plant_master,PID_out_plant_master,VV_pos_plant_master


'c   PLANT MASTER Controller
'   PC_REALS(290) = dials(39)
'   PC_REALS(291) = dials(40)
'   PC_REALS(292) = dials(41)


' BALANCE CONTROL
     
If INC_MAN_Mousedown Then
    
      If (H_gauge.Value < 255) Then H_gauge.Value = H_gauge.Value + 1
     ' H_gauge.Value = H_gauge.Value + 1
                
      Call setCONTROLLERvars(62, H_gauge.Value)     ' Balance_plant_master
                  
ElseIf DEC_MAN_Mousedown Then
    
      If (H_gauge.Value > 0) Then H_gauge.Value = H_gauge.Value - 1
     'H_gauge.Value = H_gauge.Value - 1
          
      Call setCONTROLLERvars(62, H_gauge.Value)       ' Balance_plant_master
Else
      H_gauge.Value = Format(LSS_reals(292), "##0")     ' BOTTOM Gauge
End If

    
' === 112/12/01 4.120 ===
Balance_lbl.Caption = Format(LSS_reals(292) * 100# / 255#, "##0.0") & "%"


' AUTO CONTROL
       
If INC_SP_Mousedown Then
    
      EdgeGaugectl1.Value = EdgeGaugectl1.Value + 1
          
      Call setCONTROLLERvars(61, EdgeGaugectl1.Value)   ' SP_auto_plant_master
         
ElseIf DEC_SP_Mousedown Then

      EdgeGaugectl1.Value = EdgeGaugectl1.Value - 1
          
      Call setCONTROLLERvars(61, EdgeGaugectl1.Value)   ' SP_auto_plant_master
          
Else
      EdgeGaugectl1.Value = Format(LSS_reals(290), "##0")   ' LEFT Gauge
End If
   


'c   PLANT MASTER Controller
'   PC_REALS(290) = dials(39)
'   PC_REALS(291) = dials(40)
'   PC_REALS(292) = dials(41)



'   Measured Value on the RIGHT HAND GAUGE
          EdgeGaugectl2.Value = LSS_reals(291)   ' RIGHT Gauge
    
' === 18/12/01 TESTING ===
'          PC_REALS(831) = PID_out_MASTER(1)
'          PC_REALS(832) = mV_fuel(1)
'          PC_REALS(833) = mV_air(1)
'          PC_REALS(834) = PID_out_air(1)
'          PC_REALS(835) = PID_out_fuel(1)

List1.clear
List1.AddItem "SP_auto_master(1) " & Format(LSS_reals(535) / 255#, "#0.000")
List1.AddItem "PID_out_master(1) " & Format(LSS_reals(831), "#0.000")
List1.AddItem "mV_fuel(1)        " & Format(LSS_reals(832), "#0.000")
List1.AddItem "mV_air(1)         " & Format(LSS_reals(833), "#0.000")
List1.AddItem "PID_out_air(1)    " & Format(LSS_reals(834), "#0.000")
List1.AddItem "PID_out_fuel(1)   " & Format(LSS_reals(835), "#0.000")



End Sub

Private Sub VLV_label_Click()

'   Call set_PCDI_LSS( )

End Sub
