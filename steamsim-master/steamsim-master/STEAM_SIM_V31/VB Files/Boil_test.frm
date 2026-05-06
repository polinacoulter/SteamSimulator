VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Boil_test 
   ClientHeight    =   12015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   12015
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows Default
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   11895
      Left            =   0
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   20981
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   540
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RUN (COLD)"
         Height          =   495
         Left            =   60
         TabIndex        =   35
         Top             =   540
         Width           =   1035
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1635
         Left            =   2100
         Max             =   20000
         TabIndex        =   34
         Top             =   3480
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Tsat to Psat"
         Height          =   435
         Left            =   300
         TabIndex        =   32
         Top             =   10680
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Index           =   0
         Left            =   1740
         TabIndex        =   31
         Text            =   "200 deg C"
         Top             =   10680
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Open COM LSS"
         Height          =   375
         Left            =   300
         TabIndex        =   30
         Top             =   10140
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Index           =   1
         Left            =   3600
         TabIndex        =   29
         Text            =   "xxx Bar"
         Top             =   10680
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Psat to Tsat"
         Height          =   435
         Left            =   300
         TabIndex        =   28
         Top             =   11220
         Width           =   1395
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   1740
         TabIndex        =   27
         Text            =   "40.0 Bar"
         Top             =   11220
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   26
         Text            =   "xxx deg C"
         Top             =   11280
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "    250 mS Normal Time"
         Height          =   495
         Left            =   60
         TabIndex        =   24
         Top             =   3060
         Width           =   1035
      End
      Begin VB.CommandButton Command6 
         Caption         =   "    1 sec  Slow Time"
         Height          =   495
         Left            =   60
         TabIndex        =   23
         Top             =   3660
         Width           =   1035
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1635
         Left            =   10920
         Max             =   1000
         TabIndex        =   21
         Top             =   7020
         Width           =   195
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Psat to hf  hfg   hg"
         Height          =   435
         Left            =   5640
         TabIndex        =   20
         Top             =   10920
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   7440
         TabIndex        =   19
         Text            =   "40.0 Bar"
         Top             =   10920
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   3
         Left            =   9300
         TabIndex        =   18
         Text            =   "xxx kJ/kg   xxx kJ/kg   xxx kJ/kg"
         Top             =   10920
         Width           =   3315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Psteam , Hsteam to Tsteam"
         Height          =   435
         Left            =   300
         TabIndex        =   17
         Top             =   9240
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   4
         Left            =   2460
         TabIndex        =   16
         Text            =   "40.0 Bar"
         Top             =   9000
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   6
         Left            =   2460
         TabIndex        =   15
         Text            =   "3094 kJ/kg"
         Top             =   9420
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Psteam , Tsteam qDryness to Hsteam"
         Height          =   435
         Left            =   300
         TabIndex        =   14
         Top             =   7260
         Width           =   4515
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   7
         Left            =   60
         TabIndex        =   13
         Text            =   "40.0 Bar"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   8
         Left            =   1860
         TabIndex        =   12
         Text            =   "300.0 degC"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   9
         Left            =   4860
         TabIndex        =   11
         Text            =   "xxxxx kJ/kg  xx %   xx.xxx kJ/kgdegC"
         Top             =   8280
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   10
         Left            =   3660
         TabIndex        =   10
         Text            =   "1.000"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   11
         Left            =   60
         TabIndex        =   9
         Top             =   8280
         Width           =   4695
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   1800
         TabIndex        =   8
         Top             =   5280
         Width           =   3315
      End
      Begin VB.CommandButton Burn_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner OFF/ON"
         Height          =   495
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   1035
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   10680
         TabIndex        =   6
         Top             =   5040
         Width           =   2475
      End
      Begin VB.CommandButton Command7 
         Caption         =   "    10 mS Fast Time"
         Height          =   495
         Left            =   60
         TabIndex        =   5
         Top             =   2460
         Width           =   1035
      End
      Begin VB.ListBox List3 
         Height          =   1815
         Left            =   6660
         TabIndex        =   4
         Top             =   3540
         Width           =   2895
      End
      Begin VB.CommandButton Command11 
         Caption         =   "RESUME"
         Height          =   495
         Left            =   60
         TabIndex        =   3
         Top             =   1740
         Width           =   1035
      End
      Begin VB.CommandButton Command12 
         Caption         =   "HALT"
         Height          =   495
         Left            =   60
         TabIndex        =   2
         Top             =   1200
         Width           =   1035
      End
      Begin VB.ListBox List4 
         Height          =   2205
         Left            =   5220
         TabIndex        =   1
         Top             =   5400
         Width           =   2895
      End
      Begin VB.ListBox List5 
         Height          =   1035
         Left            =   4380
         TabIndex        =   0
         Top             =   9000
         Width           =   6195
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve1 
         Height          =   375
         Index           =   0
         Left            =   9000
         TabIndex        =   22
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump Pump1 
         Height          =   495
         Left            =   9240
         TabIndex        =   25
         Top             =   105
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 HavenValveV1_21 
         Height          =   735
         Left            =   11760
         TabIndex        =   33
         Top             =   9720
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1296
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   10
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve1 
         Height          =   555
         Index           =   1
         Left            =   9000
         TabIndex        =   36
         Top             =   1200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   979
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve1 
         Height          =   375
         Index           =   2
         Left            =   11160
         TabIndex        =   37
         Top             =   7440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         TopColour       =   65280
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump Pump2 
         Height          =   495
         Left            =   1260
         TabIndex        =   38
         Top             =   3180
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrLineWidth  =   3
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   1260
         X2              =   8940
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Q_steel 
         Caption         =   "Qsteel(0)"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   70
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label P_steam 
         Caption         =   "Psteam(0)"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   69
         Top             =   5760
         Width           =   1515
      End
      Begin VB.Label T_steam 
         Caption         =   "Tsteam(1)"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   68
         Top             =   420
         Width           =   2355
      End
      Begin VB.Label M_steam 
         Caption         =   "Msteam(0)"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   67
         Top             =   5340
         Width           =   1455
      End
      Begin VB.Label P_sat 
         Caption         =   "Psat"
         Height          =   255
         Left            =   4800
         TabIndex        =   66
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label F_evap 
         Caption         =   "Fevap"
         Height          =   255
         Left            =   9360
         TabIndex        =   65
         Top             =   2100
         Width           =   2475
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3135
         Left            =   1260
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label M_water 
         Caption         =   "Mwater"
         Height          =   255
         Left            =   1860
         TabIndex        =   64
         Top             =   2940
         Width           =   2415
      End
      Begin VB.Label T_water 
         Caption         =   "Twater"
         Height          =   255
         Left            =   1860
         TabIndex        =   63
         Top             =   2580
         Width           =   2415
      End
      Begin VB.Label h_fg 
         Caption         =   "hfg"
         Height          =   255
         Left            =   3420
         TabIndex        =   62
         Top             =   2220
         Width           =   3495
      End
      Begin VB.Label P_actual 
         Caption         =   "Pactual"
         Height          =   195
         Left            =   7140
         TabIndex        =   61
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label P_air 
         Caption         =   "Pair"
         Height          =   195
         Left            =   4860
         TabIndex        =   60
         Top             =   1260
         Width           =   1875
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   8820
         X2              =   9240
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label M_air 
         Caption         =   "Mair"
         Height          =   255
         Left            =   4860
         TabIndex        =   59
         Top             =   780
         Width           =   1875
      End
      Begin VB.Label T_air 
         Caption         =   "Tair"
         Height          =   255
         Left            =   4860
         TabIndex        =   58
         Top             =   420
         Width           =   1875
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   1260
         Y2              =   1500
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   4440
         X2              =   4680
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   6840
         X2              =   7020
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   6840
         X2              =   7020
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   4380
         X2              =   4680
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4680
         Y1              =   2580
         Y2              =   2700
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4680
         Y1              =   2820
         Y2              =   2700
      End
      Begin VB.Label P_amb 
         Caption         =   "Pamb"
         Height          =   255
         Left            =   4380
         TabIndex        =   57
         Top             =   4140
         Width           =   1875
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   8820
         X2              =   9240
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label T_amb 
         Caption         =   "Tamb"
         Height          =   255
         Left            =   4380
         TabIndex        =   56
         Top             =   3780
         Width           =   1875
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   9180
         X2              =   9180
         Y1              =   2040
         Y2              =   2640
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   9180
         X2              =   9060
         Y1              =   2040
         Y2              =   2220
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   9300
         X2              =   9180
         Y1              =   2220
         Y2              =   2040
      End
      Begin VB.Line Line17 
         BorderWidth     =   4
         X1              =   8820
         X2              =   12060
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label T_sat 
         Caption         =   "Tsat"
         Height          =   255
         Left            =   6480
         TabIndex        =   55
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label F_out2 
         Caption         =   "Fout2"
         Height          =   255
         Left            =   10140
         TabIndex        =   54
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label F_out1 
         Caption         =   "Fout1"
         Height          =   255
         Left            =   10140
         TabIndex        =   53
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label F_out_pump 
         Caption         =   "Fout_pump"
         Height          =   255
         Left            =   9900
         TabIndex        =   52
         Top             =   60
         Width           =   3315
      End
      Begin VB.Line Line18 
         BorderWidth     =   4
         X1              =   12060
         X2              =   12060
         Y1              =   4920
         Y2              =   1620
      End
      Begin VB.Label V_pos 
         Alignment       =   1  'Right Justify
         Caption         =   "Vpos = xx %"
         Height          =   255
         Left            =   9000
         TabIndex        =   51
         Top             =   7680
         Width           =   1875
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   2700
         X2              =   2580
         Y1              =   3720
         Y2              =   3540
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   2580
         X2              =   2460
         Y1              =   3540
         Y2              =   3720
      End
      Begin VB.Line Line21 
         BorderWidth     =   2
         X1              =   2580
         X2              =   2580
         Y1              =   3540
         Y2              =   4140
      End
      Begin VB.Label h_f 
         Caption         =   "hf"
         Height          =   255
         Left            =   1380
         TabIndex        =   50
         Top             =   2220
         Width           =   1875
      End
      Begin VB.Label h_g 
         Caption         =   "hg"
         Height          =   255
         Left            =   7020
         TabIndex        =   49
         Top             =   2220
         Width           =   1875
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   1635
         Left            =   10500
         Top             =   4920
         Width           =   2775
      End
      Begin VB.Line Line22 
         BorderWidth     =   4
         X1              =   11340
         X2              =   11340
         Y1              =   8640
         Y2              =   6540
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   9780
         X2              =   10380
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line24 
         BorderWidth     =   2
         X1              =   10140
         X2              =   10380
         Y1              =   5400
         Y2              =   5520
      End
      Begin VB.Line Line25 
         BorderWidth     =   2
         X1              =   10140
         X2              =   10380
         Y1              =   5640
         Y2              =   5520
      End
      Begin VB.Label H_steam 
         Caption         =   "Hsteam(0)"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   48
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   12120
         X2              =   12600
         Y1              =   6540
         Y2              =   7920
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   12300
         X2              =   12600
         Y1              =   7740
         Y2              =   7920
      End
      Begin VB.Line Line28 
         BorderWidth     =   2
         X1              =   12600
         X2              =   12600
         Y1              =   7860
         Y2              =   7560
      End
      Begin VB.Label T_steel 
         Caption         =   "Tsteel(1)"
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   47
         Top             =   4740
         Width           =   2355
      End
      Begin VB.Label T_steam 
         Caption         =   "Tsteam(0)"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   46
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label M_steam 
         Caption         =   "Msteam(1)"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   45
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label P_steam 
         Caption         =   "Psteam(1)"
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   44
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label H_steam 
         Caption         =   "Hsteam(1)"
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   43
         Top             =   1680
         Width           =   2475
      End
      Begin VB.Label Q_steel 
         Caption         =   "Qsteel(1)"
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   42
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Q_steel 
         Caption         =   "Qsteel(2)"
         Height          =   255
         Index           =   2
         Left            =   8280
         TabIndex        =   41
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Label q_dry 
         Caption         =   "qDry(1)"
         Height          =   315
         Index           =   1
         Left            =   11040
         TabIndex        =   40
         Top             =   2880
         Width           =   2235
      End
      Begin VB.Label q_dry 
         Caption         =   "qDry(2)"
         Height          =   315
         Index           =   2
         Left            =   8940
         TabIndex        =   39
         Top             =   6660
         Width           =   2235
      End
   End
End
Attribute VB_Name = "Boil_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Burn_cmd_Click()

Burner_ON(1) = Not Burner_ON(1)
If Burner_ON(1) Then
    Burn_cmd.BackColor = RGB(0, &HFF, 0)
Else
    Burn_cmd.BackColor = RGB(&HC0, &HC0, &HC0)
End If


End Sub

Private Sub Command1_Click()
Dim i As Integer

Tamb = 20#      ' Ambient air Temp deg C
Pamb = 1.013    ' Ambient air Press Bar


' 4/6/99
R_HP(1) = 1.8333    ' Bar Seconds / kg
R_HP(2) = 0.5812    ' Bar Seconds / kg
R_LP(1) = 0.1389    ' Bar Seconds / kg
R_LP(2) = 0.02137   ' Bar Seconds / kg

' 7/6/99    R_HP_bleed = 17.4 Bar / ( 6.5% of 23.4kg/s ) = 17.4 / 1.521 = 11.44
R_HP_bleed = 11.44   ' Bar Seconds / kg

' 8/6/99    R_bleed = 3.8 Bar / ( 7% of 23.4kg/s ) = 3.8 / 1.638 = 2.32
R_bleed = 2.32   ' Bar Seconds / kg

' 8/6/99    R_LP_bleed = 0.55 Bar / ( 5% of 23.4kg/s ) = 0.55 / 1.17 = 0.47
R_LP_bleed = 0.47   ' Bar Seconds / kg



' Normally   R_turbine_A = 0.02137*0.47/(0.02137+0.47) = 0.02044  Bar Seconds / kg
R_turbine_A = (R_LP(2) * R_LP_bleed) / (R_LP(2) + R_LP_bleed)

' Normally   Res_A = 0.1389 + 0.02044 = 0.15934  Bar Seconds / kg
Res_A = R_LP(1) + R_turbine_A

' Normally   R_turbine_B = 0.15934*2.32/(0.15934+2.32) = 0.1491  Bar Seconds / kg
R_turbine_B = (Res_A * R_bleed) / (Res_A + R_bleed)

' Normally   Res_B = 0.5812 + 0.1491 = 0.7303   Bar Seconds / kg
Res_B = R_HP(2) + R_turbine_B

' Normally   R_turbine_C = 0.7303*11.44/(0.7303+11.44) = 0.6865  Bar Seconds / kg
R_turbine_C = (Res_B * R_HP_bleed) / (Res_B + R_HP_bleed)


' 8/6/99
'R_turbines = R_HP(1) + R_HP(2) + R_LP(1) + R_LP(2)
' Normally   R_turbines = 1.8333 + 0.6865 = 2.5198
R_turbines = R_HP(1) + R_turbine_C



' 15/4/99
'Mwater(1) = 500#   ' kg
Mwater(1) = 10000#  ' kg = 10 Tons ( occupies approx 10m^3 volume )


Vwater = 10#        ' m^3    ( 10000 kg occupies 10m^3 )

Msteam(1) = 0.0001
Psteam(1) = 0.0001     ' Water vapour pressure in the drum ( Bar )

For i = 1 To 10

    Fsteam(i) = 0#         ' kg/sec
    Psteam(i) = 0.0001
    
    Twater(i) = 20#    ' deg C
    Tsteam(i) = 20#    ' deg C
        
    Tsteel(i) = 20#    ' deg C
    
    Qsteel(i) = 0#
    Qsteel_amb(i) = 0#
    
    Fwater(i) = 0#
    
Next i

Psteam(14) = 0.05       ' Condenser pressure Bar

For i = 1 To 2  ' 2 Boilers

    'Pair = 1.013  ' Air pressure in the drum  ( Bar ) NB Pair is calculated from Mair
    
    Tair(i) = 20#  ' Air temp ( deg C )
    Vair(i) = 10#  ' m^3


' 15/4/99
' Mair = PV/RT      NB P is in kN/m^2 because gas constant Rair is in kJ/(kg degK)
' Also T MUST be in deg K and NOT in deg C.....
    Mair(i) = (1.013 * 100#) * Vair(i) / (Rair * (Tair(i) + 273.15))  ' 12.036 kg for 10.0 m^3 at 20 deg C

Next i


Psat = 0.02337  ' Psat in Bar (at 20.0 deg C)

CPW = 4.2   ' Specific Heat of water    kJ / (kg degK)

' CPSTEAM varies , default value is....
CPsteam = 2.2   ' Specific Heat of steam    kJ / (kg degK)


CPmetal = 4.2   ' Specific Heat of Iron    kJ / (kg degK)


boil_run_flag = True
Timer1.Enabled = True

End Sub

Private Sub Command10_Click()

' 19/5/99
Dim P1_steam As Single
Dim T1_steam As Single
Dim H1_steam As Single

Dim q1_dryness As Single
Dim T1_sat As Single
Dim hf_1 As Single
Dim hfg_1 As Single
Dim hg_1 As Single

' 28/4/99
P1_steam = Val(Text2(7).text)
T1_steam = Val(Text2(8).text)

q1_dryness = Val(Text2(10).text)

Call get_Tsat(P1_steam, T1_sat)

Call get_hf_hfg(P1_steam, hf_1, hfg_1)


hg_1 = hf_1 + hfg_1

Text2(11).text = "hf = " & Format(hf_1, "###0") & "  hfg = " & Format(hfg_1, "###0") & "  hg = " & Format(hg_1, "###0")




Call get_Hsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_Steam)

Text2(9).text = "H=" & Format(H1_steam, "###0.0") & " kJ/kg " & _
" qDryn=" & Format(q1_dryness, "###0.000") & " CP_steam=" & Format(CP_Steam, "###0.000")


End Sub

Private Sub Command11_Click()

boil_run_flag = True

End Sub

Private Sub Command12_Click()

boil_run_flag = False

End Sub

Private Sub Command2_Click()

Tsat = Val(Text1(0).text)
Call get_Psat(Psat, Tsat)
Text1(1).text = Psat & " Bar"

End Sub

Private Sub Command3_Click()

Call open_com_LSS(1)

End Sub

Private Sub Command4_Click()

Psat = Val(Text2(0).text)
Call get_Tsat(Psat, Tsat)
Text2(1).text = Tsat & " deg C"

End Sub

Private Sub Command5_Click()

'Timer1.Interval = 100
Timer1.Interval = 250

End Sub

Private Sub Command6_Click()

Timer1.Interval = 1000

End Sub

Private Sub Command7_Click()

Timer1.Interval = 10

End Sub

Private Sub Command8_Click()

' 20/4/99
Psat = Val(Text2(2).text)
Call get_hf_hfg(Psat, hf, hfg)
hg = hf + hfg
Text2(3).text = Format(hf, "###0") & "  " & Format(hfg, "###0") & "  " & Format(hg, "###0")


End Sub

Private Sub Command9_Click()
' 19/5/99
Dim P1_steam As Single
Dim T1_steam As Single
Dim H1_steam As Single

Dim q1_dryness As Single
Dim T1_sat As Single
Dim hf_1 As Single
Dim hfg_1 As Single
Dim hg_1 As Single


' 20/4/99
P1_steam = Val(Text2(4).text)
H1_steam = Val(Text2(6).text)

Call get_Tsat(P1_steam, T1_sat)

Call get_hf_hfg(P1_steam, hf_1, hfg_1)

hg_1 = hf_1 + hfg_1


Call get_Tsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_Steam)

' 28/6/99
List5.clear
List5.AddItem "T1_steam = " & Format(T1_steam, "###0.00000000") & "  deg C "
List5.AddItem "q1_dryness = " & Format(q1_dryness, "###0.000") & "    CP_steam = " & Format(CP_Steam, "###0.000")

List5.AddItem "Tsat =  " & Format(T1_sat, "###0.00000000") & "  deg C "
List5.AddItem "  hf =  " & Format(hf_1, "###0.0") & "  hfg =  " & Format(hfg_1, "###0.0")


End Sub

Private Sub Form_Load()


' For Normal time simulation...
STEP = Timer1.Interval * 0.001


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub


Private Sub Pump1_Click()


Pump1.PumpOn = Not Pump1.PumpOn

End Sub

Private Sub Pump2_Click()

Pump2.PumpOn = Not Pump2.PumpOn

End Sub



Private Sub Timer1_Timer()

' 27/5/99
'Dim T_flame_average As Single
'Dim CQ_iron As Single


If boil_run_flag Then
        
        
        ' Pair is in kN/m^2...
         Pair(1) = Mair(1) * Rair * (Tair(1) + 273.15) / Vair(1)   '  from PV = mRT , P = mRT/V

        ' Convert Pair from kN/m^2 to BAR...( 100 kN/m^2 = 1 Bar )
         Pair(1) = Pair(1) * 0.01
         
                ' 17/5/99
                 Tair(1) = Tsteam(1)

              
        ' Dalton's Law of Partial Pressures...
        Pactual(1) = Pair(1) + Psteam(1)
        

    ''' 11/5/99  BOILER
    ' Call Boiler_drum_sub(Pact, T1, F1, Mwat, Twat, Mstm, Tstm, Miron, Tiron, _
    '          CQiron, CQwat, CQstm, Cevap, _
    '            Tx, Ta, H2, T2, q2)
    
          ' 26/5/99
          '  CQ_iron = 200#
            ' Need about 19,000 kW at present to generate 11.7 kg/sec at 60 Bar
           
         '             Boiler...CQiron = 19000 kW / ( 1300.0 - 282.0 ) deg C = 19.0
         ' First Super-heater...CQiron =  3295 kW / ( 1300.0 - 410.0 ) deg C = 3.702
           
           Qsteel_amb(1) = (Tsteel(1) - Tamb) * 19#    ' Heat loss to ambient
           
        '  Qsteel_amb(2) = (Tsteel(2) - Tamb) * 3.702  ' Heat loss to ambient
        '  Qsteel_amb(3) = (Tsteel(3) - Tamb) * 3.702  ' Heat loss to ambient
           Qsteel_amb(2) = (Tsteel(2) - Tamb) * 2.702  ' Heat loss to ambient
           Qsteel_amb(3) = (Tsteel(3) - Tamb) * 2.702  ' Heat loss to ambient
           
           ' Qsteel(4) is defined in Boiler_drum_sub
           
           ' Attemperator/Super-Heater Mixing....
         ' Qsteel_amb(5) = (Tsteel(5) - Tamb) * 3.702  ' Heat loss to ambient
           Qsteel_amb(5) = (Tsteel(5) - Tamb) * 1.2    ' Heat loss to ambient
           Qsteel(5) = -Qsteel_amb(5)                  ' Heat loss to ambient
          
            If Burner_ON(1) Then
            
                    ' 27/5/99
               ' T_flame_average = VScroll1.Value * 13000# / 20000#        ' 0 to 13,000.0 deg C
               ' If T_flame_average < Tamb Then T_flame_average = Tamb
                
          ' Boiler Heat...
              '  Qsteel(1) = VScroll1.Value * 19000# / 20000# - Qsteel_amb(1)     ' 0 to 19,000 kW (3500 LOSS)
                 Qsteel(1) = VScroll1.Value * 19000# / 20000#      ' 0 to 19,000 kW
          
             ' FIRST SUPER-HEATER...CQiron = 3295 kW / ( 1300.0 - 410.0 ) deg C = 3.702 (make CQ_steam same..)
              '  Qsteel(2) = VScroll1.Value * 3295# / 20000# - Qsteel_amb(2)   ' 0  to 3295 kW
              ' 1/6/99 Need another 6 deg C ( approx 6% more heat ...)
              ' 1/6/99 Need another 3 deg C ( approx 9% more heat ...)
                 Qsteel(2) = VScroll1.Value * 3295# * 1.09 / 20000#     ' 0  to 3295 kW
                 
           ' Super-heater second stage...
                  Qsteel(3) = VScroll1.Value * 3295# / 20000#       ' 0  to 3295 kW
              '   Qsteel(3) = VScroll1.Value * 1600# / 20000#       ' 0  to 1600 kW
                                            
            Else
            
               '  T_flame_average = Tamb
                 
                Qsteel(1) = -Qsteel_amb(1) ' Heat loss to ambient
                Qsteel(2) = -Qsteel_amb(2) ' Heat loss to ambient
                Qsteel(3) = -Qsteel_amb(3) ' Heat loss to ambient
                 
            End If
            
                   
                                                                        
         
         ' 26/5/99 Changed Mmetal from 30.0 to 10000.0
         ' 26/5/99 Added:
         If Pump2.PumpOn Then
            Fwater(10) = 12#
            Twater(10) = Twater(1) - 2#
         Else
            Fwater(10) = 0#
            Twater(10) = 20#
         End If
         
         
          ' Boiler_drum_sub  WORKS OK ( Even with Cevap = 100.0 )
         
         '  NB Qwat is now multiplied by Mwat....( So CQwat is reduced from 10000.0 to 1.0 )
         '  NB Qstm is now multiplied by Mstm....( So CQstm is reduced from    50.0 to 1.0 )
         
         
         
        '   1/6/99    Tattemp = Tsteam(4) , Fattemp = Fsteam(2)
        
 
            Call Boiler_drum_sub(Pactual(1), Tsteam(4), Fsteam(2), Qsteel(4), _
            Mwater(1), 10000#, Twater(1), Fwater(10), Twater(10), _
            Msteam(1), 60#, Tsteam(1), 10000#, Tsteel(1), Qsteel(1), _
             1#, 1#, 100#, Fevap(1), _
            Hsteam(1), qDry(1))
            
                        
                    List1.clear
        List1.AddItem "Tsteam(1) = " & Format(Tsteam(1), "##0.0") & Chr(176) & "C  " & Format(Hsteam(1), "###0") & " kJ/kg"
        List1.AddItem "Tsteam(2) = " & Format(Tsteam(2), "##0.0") & Chr(176) & "C  " & Format(Hsteam(2), "###0") & " kJ/kg"
        List1.AddItem "Tsteam(3) = " & Format(Tsteam(3), "##0.0") & Chr(176) & "C  " & Format(Hsteam(3), "###0") & " kJ/kg"
        List1.AddItem "Tsteam(4) = " & Format(Tsteam(4), "##0.0") & Chr(176) & "C  " & Format(Hsteam(4), "###0") & " kJ/kg"
        List1.AddItem "Tsteam(5) = " & Format(Tsteam(5), "##0.0") & Chr(176) & "C  " & Format(Hsteam(5), "###0") & " kJ/kg"
            
              
           '  VALVE1(1)  (Big exhaust valve)
        If valve_status(1) = 1 Then
            Psteam(2) = Pactual(1)
            Psteam(3) = Pactual(1)
            Psteam(4) = Pactual(1)
            Psteam(5) = Pactual(1)
       End If


          '  sub Heater_sub(Pact, T1, F1, Mstm, Tstm, Mmetal, Tmetal, Qmetal, _
                            CQstm, Hstm, qDryness)
                            
        ' FIRST SUPER-HEATER...CQiron = 3295 kW / ( 1300.0 - 410.0 ) deg C = 3.702 (make CQ_steam same..)
        Call Heater_sub(Psteam(2), Tsteam(1), Fsteam(1), 10#, Tsteam(2), 500#, Tsteel(2), Qsteel(2), _
                    3.702, Hsteam(2), qDry(2))
                                     
                
                List2.clear
                   
           '   List2.AddItem "R_turbine_A = " & Format(R_turbine_A, "##0.0000")
           '   List2.AddItem "R_turbine_B = " & Format(R_turbine_B, "##0.0000")
           '   List2.AddItem "R_turbine_C = " & Format(R_turbine_C, "##0.0000")
           '   List2.AddItem "Res_A = " & Format(Res_A, "##0.0000")
           '   List2.AddItem "Res_B = " & Format(Res_B, "##0.0000")
           '   List2.AddItem "R_turbines = " & Format(R_turbines, "##0.0000")
                  
                  
            ' List2.AddItem "Psteam(2) = " & Format(Psteam(2), "##0.00")
            ' List2.AddItem "Fsteam(1) = " & Format(Fsteam(1), "##0.00")
            ' List2.AddItem "Tsteel(2) = " & Format(Tsteel(2), "##0.0")
            ' List2.AddItem "Qsteel(2) = " & Format(Qsteel(2), "#####0")
            ' List2.AddItem "Hsteam(2) = " & Format(Hsteam(2), "####0")
             
             
             
        '  ATTEMPARATOR CONTROL VALVE  ( 506 to 514 degC )
        If Tsteam(5) < 506# Then
            V_attemp_pos_set(1) = 0#
        ElseIf Tsteam(5) > 514# Then
            V_attemp_pos_set(1) = 100#
        Else
            V_attemp_pos_set(1) = (Tsteam(5) - 506#) / 8# * 100#    ' Linear operation
        End If
        
        'ATTEMPARATOR CONTROL VALVE
            V_attemp_pos_dot = (V_attemp_pos_set(1) - Vattemp_POS(1)) * 0.2
            Vattemp_POS(1) = Vattemp_POS(1) + V_attemp_pos_dot * STEP
        
        
            Fsteam(2) = Fsteam(1) * Vattemp_POS(1) * 0.01
            Fsteam(3) = Fsteam(1) - Fsteam(2)
            
        
        ' SECOND SUPER-HEATER...CQiron = 3295 kW / ( 1300.0 - 538.0 ) deg C = 4.324 (make CQ_steam same..)
        Call Heater_sub(Psteam(3), Tsteam(2), Fsteam(3), 10#, Tsteam(3), 500#, Tsteel(3), Qsteel(3), _
                    3.702, Hsteam(3), qDry(3))
                    
                                 
        '  ATTEMPARATOR
        ' SEE Boiler_drum_sub....     Qattemp = qsteel(4)
        Call Heater_sub(Psteam(4), Tsteam(2), Fsteam(2), 10#, Tsteam(4), 500#, Tsteel(4), -Qsteel(4), _
                    3.702, Hsteam(4), qDry(4))
                    
                    
         '  ATTEMPARATOR mixes with the SECOND Stage Super-Heater
        Call Heater_mix_sub(Psteam(5), Tsteam(4), Fsteam(2), Tsteam(3), Fsteam(3), 10#, Tsteam(5), 500#, Tsteel(5), Qsteel(5), _
                    3.702, Hsteam(5), qDry(5))
                    
       
        
                   List3.clear
             List3.AddItem "Psteam(3) = " & Format(Psteam(3), "##0.00")
             List3.AddItem "Tsteel(3) = " & Format(Tsteel(3), "##0.0")
             List3.AddItem "Qsteel(3) = " & Format(Qsteel(3), "#####0")
             List3.AddItem "Hsteam(3) = " & Format(Hsteam(3), "####0")
        
              List3.AddItem "Fsteam(2) = " & Format(Fsteam(2), "####0.0")
              List3.AddItem "Fsteam(3) = " & Format(Fsteam(3), "####0.0")
              List3.AddItem "Vattemp_pos(1) = " & Format(Vattemp_POS(1), "####0.0")
       
          
    
           
                       
           Q_steel(1).Caption = "Qsteel(1)  = " & Format(Qsteel(1), "##0") & " kW"
           Q_steel(2).Caption = "Qsteel(2)  = " & Format(Qsteel(2), "##0") & " kW"
                        
            q_dry(1).Caption = "qDry(1)  = " & Format(qDry(1), "##0.00")
            q_dry(2).Caption = "qDry(2)  = " & Format(qDry(2), "##0.00")
            
         '  VALVE1(2)  (Throttling  valve)
            Valve1(2).ColourOpen = RGB(0, V_ahead_pos * 2.5, 0)
        
        ' 4/6/99  Throttling Valve ( V_ahead_pos % )
            If V_ahead_pos = 100 Then
                    RV_ahead = 0.0000001
            Else
                    RV_ahead = 1# / (V_ahead_pos + 0.0000001)
            End If
            
                    
         '  4/6/99
         '   There is NO Enthalpy drop across the throttling valve ( V_ahead_pos ).....
         
                Hsteam(6) = Hsteam(5)
                
                 ' 4/6/99   Psteam(1) in the DRUM , Psteam(14) at the Condenser...
                    Pdiff = Psteam(1) - Psteam(14)
                    Rtotal = RV_ahead + R_turbines
                    
                    Psteam(13) = R_LP(2) * Pdiff / Rtotal + Psteam(14)
                    Psteam(12) = (R_LP(1) + R_LP(2)) * Pdiff / Rtotal + Psteam(14)
                    Psteam(11) = (R_LP(1) + R_LP(2) + R_HP(2)) * Pdiff / Rtotal + Psteam(14)
                    Psteam(6) = R_turbines * Pdiff / Rtotal + Psteam(14)
                    
               
         '  ===  HP TURBINE FIRST STAGE ===
         
         '  Torque_HP(1) = 7200kW / 100 rpm = 7200 kJ/s / (100*2*Pi/60) rads/s
         '  = 7200 kNm/s / 10.47 rads/s = 687.7 kNm
         
         '  Turbo_K = 687.7 kNm / ( 60.3Bar-17.4Bar)   =  16.03 kNm/Bar
         
         '    Turbine_sub(Pin, Pout, Turbo_K, Fsteam, Hin, Hout, Turbo_Torque, Turbo_power, Tout, qDry_out)
         '               (IN , IN  ,    IN  ,    IN ,  IN,  OUT,          OUT,         OUT,  OUT,    OUT  )
    
    
            ' NOTE: ONLY ONE BOILER at present so Fsteam(1)*2.0
            Call Turbine_sub(Psteam(6), Psteam(11), 16.03, Fsteam(1) * 2#, Hsteam(6), Hsteam(11), Torque_HP(1), Power_HP(1), Tsteam(11), qDry(11))
            
            
         '  ===  HP TURBINE SECOND STAGE ===
         
         '  Torque_HP(2) = 5700kW / 100 rpm = 5700 kJ/s / (100*2*Pi/60) rads/s
         '  = 5700 kNm/s / 10.47 rads/s = 544.4 kNm
         
         '  Turbo_K = 544.4 kNm / ( 17.4 Bar - 3.8 Bar)   =  40.03 kNm/Bar
         
            ' NOTE: ONLY ONE BOILER at present so Fsteam(1)*2.0
            Call Turbine_sub(Psteam(11), Psteam(12), 40.03, Fsteam(1) * 2#, Hsteam(11), Hsteam(12), Torque_HP(2), Power_HP(2), Tsteam(12), qDry(12))
            
            
         '  ===  LP TURBINE FIRST STAGE ===
         
         '  Torque_LP(1) = 5300kW / 100 rpm = 5300 kJ/s / (100*2*Pi/60) rads/s
         '  = 5300 kNm/s / 10.47 rads/s = 506.2 kNm
         
         '  Turbo_K = 506.2 kNm / ( 3.8 Bar - 0.55 Bar)   =  155.75 kNm/Bar
         
         '    Turbine_sub(Pin, Pout, Turbo_K, Fsteam, Hin, Hout, Turbo_Torque, Turbo_power, Tout, qDry_out)
         '               (IN , IN  ,    IN  ,    IN ,  IN,  OUT,          OUT,         OUT,  OUT,    OUT  )
    
    
            ' NOTE: ONLY ONE BOILER at present so Fsteam(1)*2.0
            Call Turbine_sub(Psteam(12), Psteam(13), 155.75, Fsteam(1) * 2#, Hsteam(12), Hsteam(13), Torque_LP(1), Power_LP(1), Tsteam(13), qDry(13))
                        
            
            
            
            
            ' PROP SHAFT TORQUE kNm   ( 2136.8 kNm at 22,380kW , 100rpm )
            TQ_shaft = Torque_HP(1) + Torque_HP(2) + Torque_LP(1) + Torque_LP(2)
            
            ' PROP SHAFT LOAD TORQUE kNm  ( 2136.8 kNm at 22,380kW , 100rpm )
            TQ_shaft_load = Shaft_rpm * Shaft_rpm * Shaft_rpm * 2.1368 * 0.001
            
            ' Torque = Moment of Inertia * Angular Acceleration
            shaft_rpm_dot = (TQ_shaft - TQ_shaft_load) / 4000#
            
            Shaft_rpm = Shaft_rpm + shaft_rpm_dot * STEP
            
            
             List4.clear
             List4.AddItem "Psteam(6) = " & Format(Psteam(6), "##0.00") & " Bar"
             List4.AddItem "psteam(11) = " & Format(Psteam(11), "##0.00") & " Bar"
             List4.AddItem "Hsteam(6) = " & Format(Hsteam(6), "##0.00") & " kJ/kg"
             List4.AddItem "hsteam(11) = " & Format(Hsteam(11), "##0.00") & " kJ/kg"
             List4.AddItem "Torque_HP(1) = " & Format(Torque_HP(1), "##0.0") & " kNm"
             List4.AddItem "Power_HP(1) = " & Format(Power_HP(1), "##0.0") & " kW"
             List4.AddItem "tsteam(11) = " & Format(Tsteam(11), "##0.0") & Chr(176) & "C"
             List4.AddItem " "
             List4.AddItem "TQ_shaft_load = " & Format(TQ_shaft_load, "##0.0") & " kNm"
             List4.AddItem "shaft_rpm = " & Format(Shaft_rpm, "##0.0") & " rpm"
             
            
'    Global shaft_rpm As Single
'    Global shaft_rpm_dot As Single
'    Global TQ_shaft As Single
'    Global TQ_shaft_load As Single

    
    
    ' Psteam(1) in kNm...( NB Vsteam = Vair )
    Psteam(1) = Msteam(1) * Rsteam * (Tsteam(1) + 273.15) / Vair(1)   '  from PV = mRT , P = mRT/V
    ' Psteam(1) in BAR...
    Psteam(1) = Psteam(1) * 0.01
        
        
    ' Super-heated Steam Temp deg C :
    ' NOTE: There is a mass of steel surrounding the super-heated steam
    ' Assume for now that there is 10 kg of steam in the super heater...
    ' Assume this mass of steel has an equivalent water mass (thermally) of 20kg...
    
    
        Fout_air = 0#
        Fout_steam = 0#
        Fout_pump = 0#
        Fout1 = 0#
        Fsteam(1) = 0#
           
    '  PUMP:
    If Pump1.PumpOn Then
    
        Fout_air = Mair(1) * 0.2       ' We are losing air
        Fout_steam = Msteam(1) * 0.2   ' We are losing steam
        
        Fout_pump = Fout_air + Fout_steam
               
    End If


       
    
    '  VALVE1(1)  (Big exhaust valve)
    If valve_status(1) = 1 Then
    
           ' ==== 4/6/99 ====
       'If (Pactual(1) - Pamb > 0#) Then       ' ONE WAY Valve
        If (Pactual(1) - Psteam(14) > 0#) Then       ' ONE WAY Valve
        
            ' Trottling valve position = V_ahead_pos %
            
                
            
           ' ==== 4/6/99 ====
           '  Fout_air = Mair(1) * 0.005 * V_ahead_pos      ' We are losing air
           '  Fout_steam = Msteam(1) * 0.005 * V_ahead_pos  ' We are losing steam
             
              
                Fout_air = (Pair(1) - Psteam(14)) / Rtotal
              Fout_steam = (Psteam(1) - Psteam(14)) / Rtotal
                                 
             Fsteam(1) = Fout_air + Fout_steam
                                
        End If
        
    End If
    
    
            ' AIR Loss
            M_dot = Fout_air
            Mair(1) = Mair(1) - M_dot * STEP
        
            ' Steam Loss
            M_dot = Fout_steam
            Msteam(1) = Msteam(1) - M_dot * STEP
  
    
    ' === DISPLAYS ===
    
    V_pos.Caption = "V_ahead_pos = " & Format(V_ahead_pos, "##0.0") & " %"
    
    F_out1.Caption = "Fout1 = " & Format(Fout1, "##0.0000") & " kg/sec"
    F_out2.Caption = "Fsteam(1) = " & Format(Fsteam(1), "##0.0000") & " kg/sec"
    F_out_pump.Caption = "Fout_pump = " & Format(Fout_pump, "##0.0000") & " kg/sec"

    
    T_amb.Caption = "Tamb = " & Format(Tamb, "##0.0") & Chr(176) & "C"
    P_amb.Caption = "Pamb = " & Format(Pamb, "##0.000") & " Bar"
        

    M_water.Caption = "Mwater(1) = " & Format(Mwater(1), "##0.000") & " kg"
    T_water.Caption = "Twater(1) = " & Format(Twater(1), "##0.000") & Chr(176) & "C"
     
    P_sat.Caption = "Psat = " & Format(Psat, "##0.00") & " Bar"
    T_sat.Caption = "Tsat = " & Format(Tsat, "##0.0") & Chr(176) & "C"

    
    M_steam(1).Caption = "Msteam(1) = " & Format(Msteam(1), "##0.000") & " kg"
    P_steam(1).Caption = "Psteam(1) = " & Format(Psteam(1), "##0.00") & " Bar"
    
    
    T_steam(1).Caption = "Tsteam(1) = " & Format(Tsteam(1), "##0.0") & Chr(176) & "C"
    
      
    M_air.Caption = "Mair(1) = " & Format(Mair(1), "##0.000") & " kg"
    P_air.Caption = "Pair(1) = " & Format(Pair(1), "##0.00") & " Bar"
    T_air.Caption = "Tair(1) = " & Format(Tair(1), "##0.0") & Chr(176) & "C"
    
    ' 11/5/99
     T_steel(1).Caption = "Tsteel(1) = " & Format(Tsteel(1), "##0.0") & Chr(176) & "C"
   
    P_actual.Caption = "Pactual(1) = " & Format(Pactual(1), "##0.00") & " Bar"
    
    F_evap.Caption = "Fevap(1) = " & Format(Fevap(1), "##0.0000") & " kg/sec"
    

    h_f.Caption = "hf = " & Format(hf, "####0") & " kJ/kg"
    
    h_fg.Caption = "hfg = " & Format(hfg, "####0") & " kJ/kg"
    
    ' 20/4/99
 ''' 28/4/99   hg = hf + hfg
    h_g.Caption = "hg = " & Format(hg, "####0") & " kJ/kg"
    
    ' 21/4/99
    H_steam(1).Caption = "Hsteam(1)  = " & Format(Hsteam(1), "####0") & " kJ/kg"
         
End If

If valve_status(0) = 1 Then
    Valve1(0).OpenTopLeft = True
Else
    Valve1(0).OpenTopLeft = False
End If
If valve_status(1) = 1 Then
    Valve1(1).OpenTopLeft = True
Else
    Valve1(1).OpenTopLeft = False
End If



End Sub

Private Sub Valve1_Click(Index As Integer)

'Valve1.OpenBtmRight = Not Valve1.OpenBtmRight
'Valve1.OpenTopLeft = Not Valve1.OpenTopLeft
If valve_status(Index) = 0 Then
    valve_status(Index) = 1
Else
    valve_status(Index) = 0
End If

End Sub

Private Sub valve2_Click()

'valve2.OpenBtmRight = Not valve2.OpenBtmRight
'valve2.OpenTopLeft = Not valve2.OpenTopLeft

End Sub

Private Sub VScroll3_Change()

V_ahead_pos = VScroll3.Value * 0.1

End Sub

Private Sub VScroll4_Change()


End Sub
