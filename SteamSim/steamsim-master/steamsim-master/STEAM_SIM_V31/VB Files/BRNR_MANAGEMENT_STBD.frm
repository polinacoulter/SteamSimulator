VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Object = "{17FE8C5E-5183-11D3-AA4D-00C04FA979C0}#7.0#0"; "HorizontalGauge.ocx"
Object = "{237BBFE9-4F2C-11D3-AA49-00C04FA979C0}#10.0#0"; "EdgeGauge.ocx"
Begin VB.Form BRNR_MANAGEMENT_STBD 
   Caption         =   "BURNER MANAGEMENT STBD"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12810
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8655
      Left            =   60
      Top             =   60
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   15266
      Begin VB.CommandButton Command3 
         Caption         =   "F"
         Height          =   315
         Left            =   1320
         TabIndex        =   60
         Top             =   6960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "T"
         Height          =   315
         Left            =   1320
         TabIndex        =   59
         Top             =   6660
         Visible         =   0   'False
         Width           =   615
      End
      Begin EdgeGauge.EdgeGaugectl Smoke_gauge 
         Height          =   3105
         Left            =   3780
         TabIndex        =   53
         Top             =   4980
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   5477
         MinorTicks      =   1
         MajorTicks      =   4
         Value           =   0
         GaugeScaleText  =   "Smoke"
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   420
         Top             =   60
      End
      Begin HorizontalGauge.HGauge Scanner_gauge1 
         Height          =   555
         Index           =   0
         Left            =   8700
         TabIndex        =   49
         ToolTipText     =   "Flame Detector"
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HorizontalGauge.HGauge Scanner_gauge1 
         Height          =   555
         Index           =   1
         Left            =   10500
         TabIndex        =   50
         ToolTipText     =   "Flame Detector"
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HorizontalGauge.HGauge Scanner_gauge2 
         Height          =   555
         Index           =   0
         Left            =   8760
         TabIndex        =   51
         ToolTipText     =   "Flame Detector"
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HorizontalGauge.HGauge Scanner_gauge2 
         Height          =   555
         Index           =   1
         Left            =   10500
         TabIndex        =   52
         ToolTipText     =   "Flame Detector"
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   17
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin EdgeGauge.EdgeGaugectl O2_gauge 
         Height          =   3105
         Left            =   2340
         TabIndex        =   56
         Top             =   4980
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   5477
         MajorTicks      =   4
         Value           =   0
         GaugeScaleText  =   "Oxygen"
      End
      Begin VB.Label XS_air 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "XS_air"
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   62
         Top             =   7680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label XS_air 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "XS_air"
         Height          =   255
         Index           =   2
         Left            =   6900
         TabIndex        =   61
         Top             =   7680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label O2_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "O2_lbl"
         Height          =   255
         Left            =   2580
         TabIndex        =   58
         Top             =   8100
         Width           =   1035
      End
      Begin VB.Label Smoke_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Smoke_lbl"
         Height          =   255
         Left            =   3840
         TabIndex        =   57
         Top             =   8100
         Width           =   1035
      End
      Begin VB.Label Message_label 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Push for 5 seconds"
         Height          =   435
         Left            =   600
         TabIndex        =   54
         Top             =   2580
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Shape B_trip 
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   6
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   4560
         Width           =   255
      End
      Begin VB.Shape B_trip 
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   5
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   4020
         Width           =   255
      End
      Begin VB.Shape B_trip 
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   4
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape B_trip 
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   3
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label Burner2_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     "
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   10560
         TabIndex        =   48
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Burner2_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      OFF"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   0
         Left            =   9720
         TabIndex        =   47
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Burner1_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "          ON"
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   1
         Left            =   10500
         TabIndex        =   46
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label Burner1_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  "
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   0
         Left            =   9660
         TabIndex        =   45
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label FO_recirc_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  "
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   300
         TabIndex        =   44
         Top             =   7320
         Width           =   735
      End
      Begin VB.Label FO_recirc_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "       OPEN"
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   1020
         TabIndex        =   43
         Top             =   7320
         Width           =   735
      End
      Begin VB.Label Boiler_mode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   3
         Left            =   1380
         TabIndex        =   42
         Top             =   5460
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Label Boiler_mode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   840
         TabIndex        =   41
         Top             =   5460
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Label Boiler_mode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   1
         Left            =   300
         TabIndex        =   40
         Top             =   5460
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Logic_alarm 
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   4
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   7680
         Width           =   255
      End
      Begin VB.Shape Logic_alarm 
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   3
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   7140
         Width           =   255
      End
      Begin VB.Shape Logic_alarm 
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   2
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   6600
         Width           =   255
      End
      Begin VB.Shape Logic_alarm 
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   1
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   6000
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   9
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   6600
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   8
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   6120
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   7
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   5640
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   6
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   4800
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   5
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   4260
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   4
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   3660
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   3
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   2
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   255
      End
      Begin VB.Shape Permissive 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   1
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   1740
         Width           =   255
      End
      Begin VB.Shape B_trip 
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   2
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2340
         Width           =   255
      End
      Begin VB.Shape B_trip 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   1
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label TRIP_label 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "TRIP"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   780
         TabIndex        =   39
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label Start_label 
         BackColor       =   &H0000FF00&
         Caption         =   "START"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   780
         TabIndex        =   38
         ToolTipText     =   "Start button must be pushed for 5 seconds"
         Top             =   3300
         Width           =   555
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   8640
         X2              =   8640
         Y1              =   540
         Y2              =   8520
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2100
         X2              =   2100
         Y1              =   540
         Y2              =   8520
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   540
         Y2              =   8520
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "UNSUCCESSFUL BNR SHUTDOWN"
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   4680
         Width           =   2835
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "ALL BURNER OIL VALVES CLOSED"
         Height          =   255
         Left            =   2520
         TabIndex        =   36
         Top             =   4080
         Width           =   2835
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "LOSS OF ALL FLAME"
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   3540
         Width           =   2655
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "OIL PRESSURE LOW"
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         ToolTipText     =   "Fuel Oil pressure < 80 psi"
         Top             =   3060
         Width           =   2655
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "DRUM LEVEL LOW LOW"
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         ToolTipText     =   "Low Low trip at -8 in water pressure (active only when boiler water has beeen heated )"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "AIR FLOW LOW"
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   1860
         Width           =   2415
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "OPERATOR TRIP"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "PERMISSION TO LIGHT BURNERS"
         Height          =   255
         Left            =   5760
         TabIndex        =   30
         ToolTipText     =   "Burner#1 lights automatically"
         Top             =   6720
         Width           =   2835
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN F.O. TRIP V/V OPEN"
         Height          =   255
         Left            =   5760
         TabIndex        =   29
         Top             =   6240
         Width           =   2415
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   " WATER LEVEL SATISFACTORY"
         Height          =   255
         Left            =   5700
         TabIndex        =   28
         ToolTipText     =   "-8.5 to +5.0 in drum water level is satisfactory"
         Top             =   5700
         Width           =   2715
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "PURGE COMPLETE"
         Height          =   255
         Left            =   5760
         TabIndex        =   27
         Top             =   4860
         Width           =   2415
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "PURGE IN PROGRESS"
         Height          =   255
         Left            =   5760
         TabIndex        =   26
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "PURGE AIR FLOW SATISFACTORY"
         Height          =   255
         Left            =   5760
         TabIndex        =   25
         ToolTipText     =   "2.3 in water pressure at the uptakes is satisfactory"
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "ALL BURNER AIR REGISTERS OPEN"
         Height          =   255
         Left            =   5760
         TabIndex        =   24
         Top             =   2880
         Width           =   2955
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "ALL BURNER OIL V/VS CLOSED"
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   2340
         Width           =   2775
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "F.O. RECIRC V/V CLOSED"
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN F.O. TRIP V/V CLOSED"
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   1260
         Width           =   2775
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "FUEL OIL TRIP VALVE CLOSED"
         Height          =   375
         Left            =   540
         TabIndex        =   20
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "IGNITER PROBE PROBLEM"
         Height          =   255
         Left            =   9120
         TabIndex        =   19
         Top             =   7740
         Width           =   3015
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "FLAME DETECTOR CHECK PROBLEM"
         Height          =   255
         Left            =   9120
         TabIndex        =   18
         Top             =   7200
         Width           =   3015
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "BURNER TRIPPED"
         Height          =   255
         Left            =   9120
         TabIndex        =   17
         Top             =   6660
         Width           =   3015
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "COMBUSTION CONTROL IN LOW FIRE"
         Height          =   255
         Left            =   9120
         TabIndex        =   16
         Top             =   6060
         Width           =   3075
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "BURNER MANAGEMENT IN MANUAL"
         Height          =   255
         Left            =   9120
         TabIndex        =   15
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIC ALARMS"
         Height          =   315
         Left            =   9120
         TabIndex        =   14
         Top             =   4980
         Width           =   2235
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PURGE PERMISSIVES"
         Height          =   315
         Left            =   6000
         TabIndex        =   13
         Top             =   660
         Width           =   2235
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CAUSE OF BOILER TRIP"
         Height          =   315
         Left            =   2580
         TabIndex        =   12
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "SCANNER R"
         Height          =   255
         Left            =   10860
         TabIndex        =   11
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "SCANNER R"
         Height          =   255
         Left            =   10800
         TabIndex        =   10
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SCANNER L"
         Height          =   255
         Left            =   9000
         TabIndex        =   9
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SCANNER L"
         Height          =   255
         Left            =   9060
         TabIndex        =   8
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BURNER No 1"
         Height          =   255
         Left            =   9540
         TabIndex        =   7
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BURNER No 2"
         Height          =   255
         Left            =   9480
         TabIndex        =   6
         Top             =   2340
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "MANUAL   BLR   CONSOLE                 FRONT"
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   4980
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FUEL OIL RECIRC."
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   8040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BOILER FIRING MODE"
         Height          =   315
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Switch to CONSOLE for automatic Purge/Boiler light up"
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BOILER FIRING"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Start button must be pushed for 5 seconds"
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EMERGENCY TRIP"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Shape Permissive 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   0
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   255
      End
      Begin VB.Shape B_trip 
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   0
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   1260
         Width           =   255
      End
      Begin VB.Shape Logic_alarm 
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   0
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   5400
         Width           =   255
      End
      Begin VB.Shape MOTV_Closed 
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   240
         Shape           =   3  'Circle
         Top             =   780
         Width           =   255
      End
      Begin VB.Shape Emergency_Trip_shape 
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   795
         Left            =   720
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   675
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   8640
         X2              =   12120
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   8640
         X2              =   12120
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   0
         X2              =   2100
         Y1              =   6780
         Y2              =   6780
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   0
         X2              =   2100
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   0
         X2              =   2100
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BURNER MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   120
         Width           =   12195
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   555
         Left            =   0
         Top             =   0
         Width           =   12315
      End
      Begin VB.Shape Boiler_start_shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   795
         Left            =   720
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   8535
         Left            =   0
         Top             =   0
         Width           =   12315
      End
   End
End
Attribute VB_Name = "BRNR_MANAGEMENT_STBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Down_counter As Integer
Dim Down_flag As Boolean




Private Sub Label43_Click()

End Sub

Private Sub Boiler_mode_Click(Index As Integer)


Select Case Index
'    Case 1
'        Call setPCDI(4001)
'    Case 2
'        Call setPCDI(4002)
'    Case 3
'        Call setPCDI(4003)
    Case 1
        Call setPCDI(84)
    Case 2
        Call setPCDI(85)
    Case 3
        Call setPCDI(86)

End Select

End Sub

Private Sub Burner1_label_Click(Index As Integer)



' V OLD way....
''              76,77           Light_Burner1(1-2)
''              78,79           Light_Burner2(1-2)
'If index = 0 Then
'    status_flag = False
'Else
'    status_flag = True
'End If
'Call set_STEAM_logicals(76, status_flag)
'


' 17/2/00  This should NOT open/close the BOTV....
'Call valve_pcdi_LSS(210)

' 17/2/00 Do this instead
If Index = 1 Then
  Call setPCDI(83)
Else
  Call setPCDI(82)
End If

'    Burner_start_SW(1, 1) = 83
'    Burner_start_SW(1, 2) = 81
'
'    Burner_start_SW(2, 1) = 164
'    Burner_start_SW(2, 2) = 162
'
'    Burner_stop_SW(1, 1) = 82
'    Burner_stop_SW(1, 2) = 80
'
'    Burner_stop_SW(2, 1) = 163
'    Burner_stop_SW(2, 2) = 161




End Sub

Private Sub Burner2_label_Click(Index As Integer)




'' 2/9/99
''              76,77           Light_Burner1(1-2)
''              78,79           Light_Burner2(1-2)
'If index = 0 Then
'    status_flag = False
'Else
'    status_flag = True
'End If
'Call set_STEAM_logicals(78, status_flag)
'
'' 24/9/99

' 17/2/00
'Call valve_pcdi_LSS(209)

' 17/2/00 Do this instead
If Index = 1 Then
  Call setPCDI(81)
Else
  Call setPCDI(80)
End If

'    Burner_start_SW(1, 1) = 83
'    Burner_start_SW(1, 2) = 81
'
'    Burner_start_SW(2, 1) = 164
'    Burner_start_SW(2, 2) = 162
'
'    Burner_stop_SW(1, 1) = 82
'    Burner_stop_SW(1, 2) = 80
'
'    Burner_stop_SW(2, 1) = 163
'    Burner_stop_SW(2, 2) = 161



End Sub

Private Sub Command2_Click()


' OPEN the "FORV" valve

' === 26/09/01 ===
' OPEN the "FORV" valve
status_flag = True
Call setDI(DI_ON_VVIO_LSS(215), status_flag)

End Sub

Private Sub Command3_Click()

' OPEN the "FORV" valve

' === 26/09/01 ===
' OPEN the "FORV" valve
status_flag = False
Call setDI(DI_ON_VVIO_LSS(215), status_flag)

End Sub

Private Sub FO_recirc_label_Click(Index As Integer)

' === 13/11/01 ===
Dim close_the_MOTV As Boolean

' 24/9/99
'Call valve_pcdi_LSS(214)


' === 14/03/01 ===
'Call valve_pcdi_LSS(215)
If Index = 1 Then
        ' OPEN the "FORV" valve
        Call setPCDI(DI_ON_VVIO_LSS(215))
        ' OPEN the "MOTV" valve
        Call setPCDI(DI_ON_VVIO_LSS(217))
Else
        ' CLOSE the "FORV" valve
        Call setPCDI(DI_OFF_VVIO_LSS(215))
    
    ' === 13/11/01 ==========
    
        close_the_MOTV = True
        
    ' If Burner 1 flame then....
    If VB_steam_logicals(1) Then
        close_the_MOTV = False
    End If
    ' If Burner 2 flame then....
    If VB_steam_logicals(3) Then
        close_the_MOTV = False
    End If
    If close_the_MOTV Then
        ' CLOSE the "MOTV" valve
        Call setPCDI(DI_OFF_VVIO_LSS(217))
    End If
    
    ' === 13/11/01 end =======
    
End If

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Edge_gauge_string(6) As String

On Error Resume Next
Call JYAScreenInit(Me, id_BRNR_MANAGEMENT_STBD)


'   6/9/99 Doesn't work when an .exe file... ( Why ? )
Start_label.ToolTipText = "Start button must be pushed for 5 seconds"


For i = 0 To 1

    Scanner_gauge1(i).GaugeText 1, "0"
    Scanner_gauge1(i).GaugeText 2, "2"
    Scanner_gauge1(i).GaugeText 3, "4"
    Scanner_gauge1(i).GaugeText 4, "6"
    Scanner_gauge1(i).GaugeText 5, "8"
    Scanner_gauge1(i).GaugeText 6, "10"
    
   ' Scanner_gauge1(i).Value = 127  '  MIDDLE  ( Scaled 0 to 255 )
    Scanner_gauge1(i).Value = 0    ' ZERO
    
   ' WORKS OK...
   ' Scanner_gauge(i).SetScale "X" & Chr(13) & "%" & Chr(13)
Next i

For i = 0 To 1

    Scanner_gauge2(i).GaugeText 1, "0"
    Scanner_gauge2(i).GaugeText 2, "2"
    Scanner_gauge2(i).GaugeText 3, "4"
    Scanner_gauge2(i).GaugeText 4, "6"
    Scanner_gauge2(i).GaugeText 5, "8"
    Scanner_gauge2(i).GaugeText 6, "10"
    
    Scanner_gauge2(i).Value = 0    ' ZERO
    
Next i


' TEST
O2_gauge.MajorTicks = 6
O2_gauge.MinorTicks = 2
Edge_gauge_string(0) = "25"
Edge_gauge_string(1) = "20"
Edge_gauge_string(2) = "15"
Edge_gauge_string(3) = "10"
Edge_gauge_string(4) = "5"
Edge_gauge_string(5) = " 0"
O2_gauge.GaugeText Edge_gauge_string  ' Re-size smudge at present
O2_gauge.GaugeScaleText = "Oxygen"

Smoke_gauge.MajorTicks = 6
Smoke_gauge.MinorTicks = 2
Edge_gauge_string(0) = "100"
Edge_gauge_string(1) = "80"
Edge_gauge_string(2) = "60"
Edge_gauge_string(3) = "40"
Edge_gauge_string(4) = "20"
Edge_gauge_string(5) = " 0"
Smoke_gauge.GaugeText Edge_gauge_string  ' Re-size smudge at present
Smoke_gauge.GaugeScaleText = "Smoke"


' === 12/10/01 ===
' Gauge1 does NOT exist so REMOVE the following code....
' This is a bit weird but can make an .exe which instals and runs OK
' on our d: test drive on PC08 even WITH the following code
' However S Green is having a "file not found error" on burner management
' We need to find out what this problem is but
' REMOVE the following code anyway because Gauge1 does NOT exist....
'
'Gauge1.NumLabels = 6    ' Same as majorticks
'Gauge1.NumMarkers = 2   ' MinirTicks
'Gauge1.ScaleText = Chr(176) & "C"
'Gauge1.GaugeLabel (1), "0"
'Gauge1.GaugeLabel (2), "20"
'Gauge1.GaugeLabel (3), "40"
'Gauge1.GaugeLabel (4), "60"
'Gauge1.GaugeLabel (5), "80"
'Gauge1.GaugeLabel (6), "100"
'Gauge1.Value = 127              ' MID Scale TEST

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_BRNR_MANAGEMENT_STBD)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Start_label_Click()

'  OLD
'        B_start_DI(1) = 4007
'        B_start_DI(2) = 4008

' NOT NEEDED:
'        Call setPCDI(4007)

End Sub

Private Sub Start_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 27/8/99
Boiler_start_shape.BorderWidth = 1
Boiler_start_shape.BorderStyle = 3  ' DOT
Down_flag = True

' 6/9/99
Message_label.Visible = True


End Sub

Private Sub Start_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 27/8/99
Boiler_start_shape.BorderWidth = 2
Boiler_start_shape.BorderStyle = 1  ' SOLID
Down_flag = False

' 6/9/99
Message_label.Visible = False

End Sub

Private Sub Timer1_Timer()
Dim O2_val As Integer
Dim smoke_val As Integer

If Down_flag Then
    If (Down_counter < 10) Then
        Down_counter = Down_counter + 1
    Else
    
       ' Call setPCDI(4007)
         Call setPCDI(90)   ' START
        
        Down_flag = False
    End If
Else
    Down_counter = 0
End If

' === 06/01/03 ===
'           PC_REALS(865) = XS_air(1)
XS_air(1).Caption = Format(LSS_reals(865), "##0.00")
XS_air(2).Caption = Format(LSS_reals(866), "##0.00")



'        PC_REALS(119) = Flame_scanner1_left(1)
'        PC_REALS(120) = Flame_scanner1_right(1)
'
'        PC_REALS(121) = Flame_scanner2_left(1)
'        PC_REALS(122) = Flame_scanner2_right(1)
'
'        PC_REALS(123) = Flame_scanner1_left(2)
'        PC_REALS(124) = Flame_scanner1_right(2)
'
'        PC_REALS(125) = Flame_scanner2_left(2)
'        PC_REALS(126) = Flame_scanner2_right(2)


'c   === 4/9/00 ===
'    equivalence(STEAM_logicals(130),Burner_Management_Power(1))
'    logical*1           Burner_Management_Power(2)

' 4/9/00
If VB_steam_logicals(130) Then
  Scanner_gauge1(0).Value = LSS_reals(119) * 25.5
  Scanner_gauge1(1).Value = LSS_reals(120) * 25.5

  Scanner_gauge2(0).Value = LSS_reals(121) * 25.5
  Scanner_gauge2(1).Value = LSS_reals(122) * 25.5
Else
  Scanner_gauge1(0).Value = 0
  Scanner_gauge1(1).Value = 0
  Scanner_gauge2(0).Value = 0
  Scanner_gauge2(1).Value = 0
End If


'   === 26/7/00 ===
'      PC_REALS(137) = O2_percent_flue_gas(1) * 100#
'      PC_REALS(138) = O2_percent_flue_gas(2) * 100#
'      PC_REALS(139) = Smoke_flue_gas(1) * 100#
'      PC_REALS(140) = Smoke_flue_gas(2) * 100#


' Oxygen Gauge is 25 % FSD so...
If LSS_reals(137) > 25# Then
  O2_val = 255
ElseIf LSS_reals(137) < 0# Then
  O2_val = 0
Else
  O2_val = LSS_reals(137) * 4# * 2.55
End If
O2_gauge.Value = O2_val
O2_lbl.Caption = Format(LSS_reals(137), "##0.0") & " %"

If LSS_reals(139) > 100# Then
  smoke_val = 255
ElseIf LSS_reals(139) < 0 Then
  smoke_val = 0
Else
  smoke_val = LSS_reals(139) * 2.55
End If
Smoke_gauge.Value = smoke_val
Smoke_lbl.Caption = Format(LSS_reals(139), "##0.0") & " %"



' Steam_Logicals
'               19-21       Boiler Firing Mode STBD ( Burner_Management.F90 )
'               22-24       Boiler Firing Mode PORT ( Burner_Management.F90 )
'               25-34       Permissives STBD   ( Burner_Management.F90 )
'               35-44       Permissives PORT   ( Burner_Management.F90 )
'               45-51           B_trip_DO(1,1-7) ( Burner_Management.F90 )
'               52-58           B_trip_DO(2,1-7) ( Burner_Management.F90 )
'
'               59-60           MOTV_closed_DO(2) ( Burner_Management.F90 )
'               61-65           B_logic_alarm(1,1-5) ( Burner_Management.F90 )
'               66-70           B_logic_alarm(2,1-5) ( Burner_Management.F90 )
'               72,73           DI( BIGN1_DI(1-2) )
'               74,75           DI( BIGN2_DI(1-2) )

' Boiler Firing Mode
' 4/9/00
If VB_steam_logicals(130) Then
  If VB_steam_logicals(19) Then
        Boiler_mode(1).BackColor = &HFF00&   ' LIGHT GREEN
        Boiler_mode(2).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(3).BackColor = &H8000&   ' DARK GREEN
  ElseIf VB_steam_logicals(20) Then
        Boiler_mode(2).BackColor = &HFF00&   ' LIGHT GREEN
        Boiler_mode(1).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(3).BackColor = &H8000&   ' DARK GREEN
  ElseIf VB_steam_logicals(21) Then
        Boiler_mode(3).BackColor = &HFF00&   ' LIGHT GREEN
        Boiler_mode(1).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(2).BackColor = &H8000&   ' DARK GREEN
  Else
        Boiler_mode(1).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(2).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(3).BackColor = &H8000&   ' DARK GREEN
  End If
Else
        Boiler_mode(1).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(2).BackColor = &H8000&   ' DARK GREEN
        Boiler_mode(3).BackColor = &H8000&   ' DARK GREEN
End If


' Boiler Trips ( 7 )
' 4/9/00
If VB_steam_logicals(130) Then
  For i = 0 To 6
    If VB_steam_logicals(45 + i) Then
        B_trip(i).FillColor = &HFF&            ' BRIGHT RED
    Else
        B_trip(i).FillColor = &H80&            ' DARK RED
    End If
  Next i
Else
  For i = 0 To 6
        B_trip(i).FillColor = &H80&            ' DARK RED
  Next i
End If


' MOTV_closed
' 24/9/99
' 4/9/00
If VB_steam_logicals(130) Then
  Call update_vvalve_LSS(Me, 217, 17, -1)
  If vv_status_LSS(217) = 0 Then
        MOTV_Closed.FillColor = &HFF&            ' BRIGHT RED
  Else
        MOTV_Closed.FillColor = &H80&            ' DARK RED
  End If
Else
        MOTV_Closed.FillColor = &H80&            ' DARK RED
End If





' Boiler Logic Alarms ( 5 )
' 4/9/00
If VB_steam_logicals(130) Then
  For i = 0 To 4
    If VB_steam_logicals(61 + i) Then
        Logic_alarm(i).FillColor = &HFF&            ' BRIGHT RED
    Else
        Logic_alarm(i).FillColor = &H80&            ' DARK RED
    End If
  Next i
Else
  For i = 0 To 4
        Logic_alarm(i).FillColor = &H80&            ' DARK RED
  Next i
End If


' Permissives ( 10 )
' 4/9/00
If VB_steam_logicals(130) Then
  For i = 0 To 9
    If VB_steam_logicals(25 + i) Then
        Permissive(i).FillColor = &HFF00&   ' LIGHT GREEN
    Else
        Permissive(i).FillColor = &H8000&   ' DARK GREEN
    End If
  Next i
Else
  For i = 0 To 9
        Permissive(i).FillColor = &H8000&   ' DARK GREEN
  Next i
End If


' 4/9/00
If VB_steam_logicals(130) Then
  Call find_VV_status_LSS(215)
  Call LABEL_control(FO_recirc_label(0), FO_recirc_label(1), "       CLOSED", "       OPEN", 215)
Else
  FO_recirc_label(0).BackColor = &H80&           ' DARK RED
  FO_recirc_label(1).BackColor = &H8000&         ' DARK GREEN
  FO_recirc_label(0).Caption = " "
  FO_recirc_label(1).Caption = " "
End If

' 4/9/00
If VB_steam_logicals(130) Then
  Call Burner_LABEL_control(Burner1_label(0), Burner1_label(1), "       CLOSED", "       OPEN", 49, flash_burner_1_1)

  Call Burner_LABEL_control(Burner2_label(0), Burner2_label(1), "       CLOSED", "       OPEN", 52, flash_burner_1_2)
Else
  Burner1_label(0).BackColor = &H80&           ' DARK RED
  Burner1_label(1).BackColor = &H8000&         ' DARK GREEN
  Burner1_label(0).Caption = " "
  Burner1_label(1).Caption = " "
  Burner2_label(0).BackColor = &H80&           ' DARK RED
  Burner2_label(1).BackColor = &H8000&         ' DARK GREEN
  Burner2_label(0).Caption = " "
  Burner2_label(1).Caption = " "
End If


End Sub
Sub LABEL_control(label_0 As Object, label_1 As Object, _
OFF_string As String, ON_string As String, v_number As Integer)

Static flash As Boolean

flash = Not flash

If vv_status_LSS(v_number) = 0 Then
    label_0.BackColor = &HFF&             ' BRIGHT RED
    label_0.Caption = OFF_string
    label_1.BackColor = &H8000&          ' DARK GREEN
    label_1.Caption = " "
ElseIf vv_status_LSS(v_number) = 1 Then
    If flash Then
        label_0.BackColor = &H80&           ' DARK RED
        label_0.Caption = " "
        label_1.BackColor = &HFF00&           ' BRIGHT GREEN
        label_1.Caption = ON_string
    Else
        label_0.BackColor = &H80&            ' DARK RED
        label_0.Caption = " "
        label_1.BackColor = &H8000&          ' DARK GREEN
        label_1.Caption = " "
    End If
ElseIf vv_status_LSS(v_number) = 2 Then
    
    If flash Then
        label_0.BackColor = &HFF&             ' BRIGHT RED
        label_0.Caption = OFF_string
        label_1.BackColor = &H8000&          ' DARK GREEN
        label_1.Caption = " "
    Else
        label_0.BackColor = &H80&            ' DARK RED
        label_0.Caption = " "
        label_1.BackColor = &H8000&          ' DARK GREEN
        label_1.Caption = " "
    End If
Else
    label_0.BackColor = &H80&            ' DARK RED
    label_0.Caption = " "
    label_1.BackColor = &HFF00&          ' BRIGHT GREEN
    label_1.Caption = ON_string
End If

End Sub

Private Sub TRIP_label_Click()

' 24/9/99
Call setPCDI(89)

End Sub

Private Sub TRIP_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 27/8/99
Emergency_Trip_shape.BorderWidth = 1
Emergency_Trip_shape.BorderStyle = 3  ' DOT

End Sub

Private Sub TRIP_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 27/8/99
Emergency_Trip_shape.BorderWidth = 2
Emergency_Trip_shape.BorderStyle = 1  ' SOLID

End Sub
Sub Burner_LABEL_control(label_0 As Object, label_1 As Object, _
OFF_string As String, ON_string As String, nlog As Integer, flash_flag As Boolean)

flash_flag = Not flash_flag


'cc       STBD BOILER Burner # 1    STOP Lamp
'       LSS_LOGICALS(49) = DO( Burner_stop_LP(1,1)+1 )
'
'cc       STBD BOILER Burner # 1    START Lamp ( FLASHING )
'       LSS_LOGICALS(50) = DO( Burner_start_LP(1,1)   )
'
'cc       STBD BOILER Burner # 1    START Lamp ( STEADY )
'       LSS_LOGICALS(51) = DO( Burner_start_LP(1,1)+1 )

' 17/2/00
'If vv_status_LSS(v_number) = 0 Then
If LSS_LOGICALS(nlog) Then

    label_0.BackColor = &HFF&             ' BRIGHT RED
    label_0.Caption = OFF_string
    label_1.BackColor = &H8000&          ' DARK GREEN
    label_1.Caption = " "
    
'ElseIf vv_status_LSS(v_number) = 1 Then
ElseIf LSS_LOGICALS(nlog + 1) Then
    If flash_flag Then
        label_0.BackColor = &H80&           ' DARK RED
        label_0.Caption = " "
        label_1.BackColor = &HFF00&           ' BRIGHT GREEN
        label_1.Caption = ON_string
    Else
        label_0.BackColor = &H80&            ' DARK RED
        label_0.Caption = " "
        label_1.BackColor = &H8000&          ' DARK GREEN
        label_1.Caption = " "
    End If
'ElseIf vv_status_LSS(v_number) = 2 Then
'
'    If flash_flag Then
'        label_0.BackColor = &HFF&             ' BRIGHT RED
'        label_0.Caption = OFF_string
'        label_1.BackColor = &H8000&          ' DARK GREEN
'        label_1.Caption = " "
'    Else
'        label_0.BackColor = &H80&            ' DARK RED
'        label_0.Caption = " "
'        label_1.BackColor = &H8000&          ' DARK GREEN
'        label_1.Caption = " "
'    End If
'Else
ElseIf LSS_LOGICALS(nlog + 2) Then
    label_0.BackColor = &H80&            ' DARK RED
    label_0.Caption = " "
    label_1.BackColor = &HFF00&          ' BRIGHT GREEN
    label_1.Caption = ON_string
End If

End Sub
