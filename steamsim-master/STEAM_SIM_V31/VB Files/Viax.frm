VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{17FE8C5E-5183-11D3-AA4D-00C04FA979C0}#7.0#0"; "HorizontalGauge.ocx"
Begin VB.Form VIAX 
   Caption         =   "VIAX"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9465
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   5955
      Left            =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10504
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4260
         Top             =   4020
      End
      Begin HorizontalGauge.HGauge Vibration_gauge 
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   900
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1085
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HorizontalGauge.HGauge Vibration_gauge 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   1
         Top             =   3180
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1085
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HorizontalGauge.HGauge Axial_gauge 
         Height          =   615
         Index           =   0
         Left            =   5400
         TabIndex        =   12
         Top             =   900
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1085
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HorizontalGauge.HGauge Axial_gauge 
         Height          =   615
         Index           =   1
         Left            =   5400
         TabIndex        =   13
         Top             =   3180
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1085
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin VB.Label LAMP_test 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " LAMP TEST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   4260
         TabIndex        =   33
         Top             =   4860
         Width           =   525
         WordWrap        =   -1  'True
      End
      Begin VB.Shape LAMP_TEST_shape 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   4140
         Top             =   4740
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TEST SELECT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5700
         TabIndex        =   31
         Top             =   5400
         Width           =   2295
      End
      Begin VB.Label FWD_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FWD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6900
         TabIndex        =   30
         Top             =   5040
         Width           =   375
      End
      Begin VB.Shape FWD_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   6840
         Top             =   4980
         Width           =   465
      End
      Begin VB.Label AFT_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AFT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         Top             =   5040
         Width           =   345
      End
      Begin VB.Shape AFT_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   6300
         Top             =   4980
         Width           =   465
      End
      Begin VB.Label TEST_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   7440
         TabIndex        =   28
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label BYPASS_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BYPASS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7380
         TabIndex        =   27
         Top             =   4440
         Width           =   645
      End
      Begin VB.Label TEST_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7440
         TabIndex        =   26
         Top             =   4140
         Width           =   525
      End
      Begin VB.Label ALARM_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALARM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   25
         Top             =   4440
         Width           =   585
      End
      Begin VB.Label TRIP_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5700
         TabIndex        =   24
         Top             =   4140
         Width           =   525
      End
      Begin VB.Label FAILURE_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PROBE FAILURE"
         Height          =   495
         Index           =   3
         Left            =   6480
         TabIndex        =   23
         Top             =   4200
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label BYPASS_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BYPASS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   7380
         TabIndex        =   22
         Top             =   2220
         Width           =   645
      End
      Begin VB.Label ALARM_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALARM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   21
         Top             =   2220
         Width           =   585
      End
      Begin VB.Label TRIP_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5700
         TabIndex        =   20
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label FAILURE_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PROBE FAILURE"
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   6480
         TabIndex        =   19
         Top             =   1980
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label BYPASS_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BYPASS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   18
         Top             =   4440
         Width           =   645
      End
      Begin VB.Label TEST_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   17
         Top             =   4140
         Width           =   525
      End
      Begin VB.Label ALARM_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALARM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   4440
         Width           =   585
      End
      Begin VB.Label TRIP_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   15
         Top             =   4140
         Width           =   525
      End
      Begin VB.Label FAILURE_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ACCEL FAILURE"
         Height          =   495
         Index           =   1
         Left            =   1920
         TabIndex        =   14
         Top             =   4200
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VIBRATION    G "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AXIAL POSITION    mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LP TURBINE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   3300
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HP TURBINE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4020
         TabIndex        =   8
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label FAILURE_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ACCEL FAILURE"
         Height          =   495
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   1980
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TURBINE MONITOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2940
         TabIndex        =   6
         Top             =   180
         Width           =   3135
      End
      Begin VB.Label TRIP_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   5
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label ALARM_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ALARM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   2220
         Width           =   585
      End
      Begin VB.Label TEST_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   3
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label BYPASS_lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BYPASS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2820
         TabIndex        =   2
         Top             =   2220
         Width           =   645
      End
      Begin VB.Shape TEST_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   0
         Left            =   2760
         Top             =   1860
         Width           =   705
      End
      Begin VB.Shape BYPASS_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   0
         Left            =   2760
         Top             =   2160
         Width           =   705
      End
      Begin VB.Shape FAILURE_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Index           =   0
         Left            =   1860
         Top             =   1860
         Width           =   765
      End
      Begin VB.Shape TRIP_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   0
         Left            =   1020
         Top             =   1860
         Width           =   705
      End
      Begin VB.Shape ALARM_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   0
         Left            =   1020
         Top             =   2160
         Width           =   705
      End
      Begin VB.Shape FAILURE_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Index           =   1
         Left            =   1860
         Top             =   4080
         Width           =   765
      End
      Begin VB.Shape TEST_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   1
         Left            =   2760
         Top             =   4080
         Width           =   705
      End
      Begin VB.Shape BYPASS_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   1
         Left            =   2760
         Top             =   4380
         Width           =   705
      End
      Begin VB.Shape TRIP_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   1
         Left            =   1020
         Top             =   4080
         Width           =   705
      End
      Begin VB.Shape ALARM_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   1
         Left            =   1020
         Top             =   4380
         Width           =   705
      End
      Begin VB.Shape TRIP_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   2
         Left            =   5580
         Top             =   1860
         Width           =   705
      End
      Begin VB.Shape ALARM_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   2
         Left            =   5580
         Top             =   2160
         Width           =   705
      End
      Begin VB.Shape FAILURE_shape 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   645
         Index           =   2
         Left            =   6420
         Top             =   1860
         Width           =   765
      End
      Begin VB.Shape TEST_shape 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   2
         Left            =   7320
         Top             =   1860
         Width           =   705
      End
      Begin VB.Shape BYPASS_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   2
         Left            =   7320
         Top             =   2160
         Width           =   705
      End
      Begin VB.Shape TRIP_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   3
         Left            =   5580
         Top             =   4080
         Width           =   705
      End
      Begin VB.Shape ALARM_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   3
         Left            =   5580
         Top             =   4380
         Width           =   705
      End
      Begin VB.Shape FAILURE_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Index           =   3
         Left            =   6420
         Top             =   4080
         Width           =   765
      End
      Begin VB.Shape TEST_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   3
         Left            =   7320
         Top             =   4080
         Width           =   705
      End
      Begin VB.Shape BYPASS_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   345
         Index           =   3
         Left            =   7320
         Top             =   4380
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   5805
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   8895
      End
   End
End
Attribute VB_Name = "VIAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VIAX_lamp_test As Boolean

Private Sub AFT_lbl_Click()


'c   === 27/11/00 ===
'    equivalence(STEAM_logicals(184),VIAX_select_AFT)
'    logical*1           VIAX_select_AFT


    status_flag = True
    
    ' === 22/11/02 ===
  ' Call set_STEAM_logicals(184, status_flag)
    Call setSTEAMlogicals(184, status_flag)




End Sub

Private Sub BYPASS_lbl_Click(Index As Integer)

' 27/11/00
If VB_steam_logicals(166 + Index) Then

    status_flag = False
    
    ' === 22/11/02 ===
    'Call set_STEAM_logicals(166 + Index, status_flag)
     Call setSTEAMlogicals(166 + Index, status_flag)
    
Else
    status_flag = True
    
    ' === 22/11/02 ===
    'Call set_STEAM_logicals(166 + Index, status_flag)
     Call setSTEAMlogicals(166 + Index, status_flag)
    
End If


End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idVIAX)

For i = 0 To 1

    Vibration_gauge(i).GaugeText 1, "0.0"
    Vibration_gauge(i).GaugeText 2, "0.2"
    Vibration_gauge(i).GaugeText 3, "0.4"
    Vibration_gauge(i).GaugeText 4, "0.6"
    Vibration_gauge(i).GaugeText 5, "0.8"
    Vibration_gauge(i).GaugeText 6, "1.0"
    
    Vibration_gauge(i).Value = 0    ' ZERO
    
   ' Vibration_gauge(i).SetScale "G"
   
   
    Axial_gauge(i).GaugeText 1, "0.0"
    Axial_gauge(i).GaugeText 2, "0.2"
    Axial_gauge(i).GaugeText 3, "0.4"
    Axial_gauge(i).GaugeText 4, "0.6"
    Axial_gauge(i).GaugeText 5, "0.8"
    Axial_gauge(i).GaugeText 6, "1.0"
    
   Axial_gauge(i).Value = 0    ' ZERO
    
    'Axial_gauge(i).SetScale "mm"
   
Next i

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)



End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idVIAX)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub FWD_lbl_Click()

'c   === 27/11/00 ===
'    equivalence(STEAM_logicals(184),VIAX_select_AFT)
'    logical*1           VIAX_select_AFT

    status_flag = False
    
    ' === 22/11/02 ===
    'Call set_STEAM_logicals(184, status_flag)
     Call setSTEAMlogicals(184, status_flag)


End Sub

Private Sub LAMP_test_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

LAMP_TEST_shape.BorderStyle = 2 ' DASH

 VIAX_lamp_test = True

End Sub

Private Sub LAMP_test_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

LAMP_TEST_shape.BorderStyle = 1 ' SOLID
 VIAX_lamp_test = False
 

End Sub

Private Sub TEST_lbl_Click(Index As Integer)



' 27/11/00
If VB_steam_logicals(162 + Index) Then

    status_flag = False
    
   ' === 22/11/02 ===
    'Call set_STEAM_logicals(162 + Index, status_flag)
     Call setSTEAMlogicals(162 + Index, status_flag)
    
Else
    status_flag = True
    
   ' === 22/11/02 ===
   'Call set_STEAM_logicals(162 + Index, status_flag)
    Call setSTEAMlogicals(162 + Index, status_flag)
    
End If



End Sub

Private Sub Timer1_Timer()
Dim i As Integer

' === 20/11/02 ===
' IF this form is minimised then it crashes with run-time error 330
' on the Vibration_gauge(0).Value = Gauge_Value line
' Changing Gauge_Value to Integer did not solve the problem
Dim Gauge_Value As Single
'Dim Gauge_Value As Integer

' FIX FOR NOW .... remove the MIN button
'
Dim FAIL_it As Byte
Dim AFT_selected As Byte
Dim FWD_selected As Byte

' 28/11/00
'         PC_REALS(810) = VIAX_reading(1)
'         PC_REALS(811) = VIAX_reading(2)
'         PC_REALS(812) = VIAX_reading(3)
'         PC_REALS(813) = VIAX_reading(4)


' == HP Vibration ==
      Gauge_Value = LSS_reals(810) * 255#
      If Gauge_Value > 255# Then
        Vibration_gauge(0).Value = 255
      ElseIf Gauge_Value < 0# Then
        Vibration_gauge(0).Value = 0
      Else
        Vibration_gauge(0).Value = Gauge_Value
      End If
      
  Call RED_box(VB_steam_logicals(170), FAILURE_lbl(0), FAILURE_shape(0))

      
' == LP Vibration ==
      Gauge_Value = LSS_reals(811) * 255#
      If Gauge_Value > 255# Then
        Vibration_gauge(1).Value = 255
      ElseIf Gauge_Value < 0# Then
        Vibration_gauge(1).Value = 0
      Else
        Vibration_gauge(1).Value = Gauge_Value
      End If
      
  Call RED_box(VB_steam_logicals(171), FAILURE_lbl(1), FAILURE_shape(1))



' == Axial Displacement  ===

' HP Axial Gauge
Gauge_Value = LSS_reals(812) * 255
If Gauge_Value > 255# Then
    Axial_gauge(0).Value = 255
ElseIf Gauge_Value < 0# Then
    Axial_gauge(0).Value = 0#
Else
    Axial_gauge(0).Value = Gauge_Value
End If


' LP Axial Gauge
Gauge_Value = LSS_reals(813) * 255
If Gauge_Value > 255# Then
        Axial_gauge(1).Value = 255
ElseIf Gauge_Value < 0# Then
        Axial_gauge(1).Value = 0#
Else
        Axial_gauge(1).Value = Gauge_Value
End If


'   === 28/11/00 ===
'    equivalence(STEAM_logicals(184),VIAX_select_AFT)
'    logical*1           VIAX_select_AFT

' AFT/FWD Selection

If VB_steam_logicals(184) Then

    
      AFT_selected = True
      FWD_selected = False
    
      ' HP AFT
      Call RED_box(VB_steam_logicals(172), FAILURE_lbl(2), FAILURE_shape(2))

     ' LP AFT
      Call RED_box(VB_steam_logicals(173), FAILURE_lbl(3), FAILURE_shape(3))
    
     'AFT_shape.FillColor = &HFF&            ' BRIGHT RED
     '   AFT_lbl.ForeColor = &HE0E0E0         ' BRIGHT GREY
     ' FWD_shape.FillColor = &H80&            ' DARK RED
     '   FWD_lbl.ForeColor = &HC0C0C0        ' DARKGREY
                 
Else

' == Axial Displacement FWD ===
      AFT_selected = False
      FWD_selected = True
  
      ' HP FWD
      Call RED_box(VB_steam_logicals(174), FAILURE_lbl(2), FAILURE_shape(2))
      
      ' LP FWD
      Call RED_box(VB_steam_logicals(175), FAILURE_lbl(3), FAILURE_shape(3))
      
      
      'AFT_shape.FillColor = &H80&            ' DARK RED
      '  AFT_lbl.ForeColor = &HC0C0C0        ' DARKGREY
      'FWD_shape.FillColor = &HFF&            ' BRIGHT RED
       ' FWD_lbl.ForeColor = &HE0E0E0         ' BRIGHT GREY
        
End If

Call RED_box(AFT_selected, AFT_lbl, AFT_shape)
Call RED_box(FWD_selected, FWD_lbl, FWD_shape)


' 27/11/00
' TEST Boxes
For i = 0 To 3

  Call RED_box(VB_steam_logicals(162 + i), TEST_lbl(i), TEST_shape(i))
  
  Call RED_box(VB_steam_logicals(166 + i), BYPASS_lbl(i), BYPASS_shape(i))

Next i



' 28/11/00
' ALARM and TRIP  Boxes

'c   === 27/11/00 ===
'    equivalence(STEAM_logicals(176),VIAX_trip(1))
'    logical*1           VIAX_trip(4)
'
'c   === 27/11/00 ===
'    equivalence(STEAM_logicals(180),VIAX_alarm(1))
'    logical*1           VIAX_alarm(4)

For i = 0 To 3

  Call RED_box(VB_steam_logicals(180 + i), ALARM_lbl(i), ALARM_shape(i))
  
  Call RED_box(VB_steam_logicals(176 + i), TRIP_lbl(i), TRIP_shape(i))

Next i




End Sub
Sub RED_box(BOX_status As Byte, label_0 As Object, shape_0 As Object)


If BOX_status Or VIAX_lamp_test Then
      shape_0.FillColor = &HFF&            ' BRIGHT RED
      label_0.ForeColor = &HE0E0E0         ' BRIGHT GREY
Else
      shape_0.FillColor = &H80&            ' DARK RED
      label_0.ForeColor = &HC0C0C0         ' DARKGREY
End If


End Sub


