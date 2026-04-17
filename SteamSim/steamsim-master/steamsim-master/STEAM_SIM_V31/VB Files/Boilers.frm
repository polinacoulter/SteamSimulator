VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Boilers 
   Caption         =   "Form1"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18105
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   18105
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   12855
      Left            =   60
      Top             =   0
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   22675
      Begin VB.ListBox List1 
         Height          =   1620
         Index           =   1
         Left            =   480
         TabIndex        =   39
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Height          =   1620
         Index           =   1
         Left            =   540
         TabIndex        =   38
         Top             =   300
         Width           =   3255
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   60
         Top             =   1860
      End
      Begin VB.ListBox List3 
         Height          =   5130
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   6840
         Width           =   5535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RUN (COLD)"
         Height          =   495
         Left            =   360
         TabIndex        =   36
         Top             =   12180
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Command7 
         Caption         =   "    10 mS Fast Time"
         Height          =   495
         Left            =   4440
         TabIndex        =   35
         Top             =   12240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Command5 
         Caption         =   "    250 mS Normal Time"
         Height          =   495
         Left            =   5580
         TabIndex        =   34
         Top             =   12240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1 OFF/ON"
         Height          =   495
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1635
         Index           =   2
         Left            =   9840
         Max             =   20000
         TabIndex        =   32
         Top             =   4140
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1635
         Left            =   6480
         Max             =   1000
         TabIndex        =   31
         Top             =   3540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton Command12 
         Caption         =   "HALT"
         Height          =   495
         Left            =   1740
         TabIndex        =   30
         Top             =   12180
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Command11 
         Caption         =   "RESUME"
         Height          =   495
         Left            =   2880
         TabIndex        =   29
         Top             =   12180
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.ListBox List8 
         Height          =   4155
         Left            =   6720
         TabIndex        =   28
         Top             =   5280
         Width           =   2715
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1 OFF/ON"
         Height          =   495
         Index           =   3
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 2 OFF/ON"
         Height          =   495
         Index           =   4
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 2 OFF/ON"
         Height          =   495
         Index           =   2
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1635
         Index           =   1
         Left            =   780
         Max             =   20000
         TabIndex        =   24
         Top             =   4140
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ListBox List2 
         Height          =   1620
         Index           =   2
         Left            =   9720
         TabIndex        =   23
         Top             =   300
         Width           =   3255
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Index           =   2
         Left            =   9780
         TabIndex        =   22
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ListBox List3 
         Height          =   3375
         Index           =   2
         Left            =   9600
         TabIndex        =   21
         Top             =   5880
         Width           =   5535
      End
      Begin VB.VScrollBar TQscroll 
         Height          =   1515
         Index           =   1
         Left            =   11700
         Max             =   2000
         TabIndex        =   20
         Top             =   10800
         Value           =   1000
         Width           =   195
      End
      Begin VB.VScrollBar Rscroll 
         Height          =   1515
         Index           =   1
         Left            =   7560
         Max             =   2000
         Min             =   10
         TabIndex        =   19
         Top             =   10860
         Value           =   1000
         Width           =   195
      End
      Begin VB.VScrollBar Rscroll 
         Height          =   1515
         Index           =   2
         Left            =   8460
         Max             =   2000
         Min             =   10
         TabIndex        =   18
         Top             =   10860
         Value           =   1000
         Width           =   195
      End
      Begin VB.VScrollBar Rscroll 
         Height          =   1515
         Index           =   3
         Left            =   9300
         Max             =   2000
         Min             =   10
         TabIndex        =   17
         Top             =   10860
         Value           =   1097
         Width           =   195
      End
      Begin VB.VScrollBar Rscroll 
         Height          =   1515
         Index           =   4
         Left            =   10080
         Max             =   2000
         Min             =   10
         TabIndex        =   16
         Top             =   10860
         Value           =   1152
         Width           =   195
      End
      Begin VB.VScrollBar TQscroll 
         Height          =   1515
         Index           =   2
         Left            =   12480
         Max             =   2000
         TabIndex        =   15
         Top             =   10800
         Value           =   1242
         Width           =   195
      End
      Begin VB.VScrollBar TQscroll 
         Height          =   1515
         Index           =   3
         Left            =   13320
         Max             =   2000
         TabIndex        =   14
         Top             =   10800
         Value           =   1112
         Width           =   195
      End
      Begin VB.VScrollBar TQscroll 
         Height          =   1515
         Index           =   4
         Left            =   14100
         Max             =   2000
         TabIndex        =   13
         Top             =   10800
         Value           =   1315
         Width           =   195
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Index           =   0
         Left            =   3480
         TabIndex        =   12
         Top             =   2100
         Width           =   2535
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Mollier"
         Height          =   615
         Left            =   7440
         TabIndex        =   11
         Top             =   3060
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1155
         Left            =   10920
         Max             =   2000
         Min             =   10
         TabIndex        =   10
         Top             =   10560
         Value           =   10
         Width           =   195
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Condenser"
         Height          =   555
         Left            =   13380
         TabIndex        =   9
         Top             =   2100
         Width           =   1335
      End
      Begin VB.ListBox List4 
         Height          =   2790
         Left            =   12420
         TabIndex        =   8
         Top             =   2760
         Width           =   5475
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1155
         Left            =   13140
         Max             =   1000
         Min             =   50
         TabIndex        =   7
         Top             =   1440
         Value           =   50
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   1155
         Left            =   6360
         Max             =   1000
         Min             =   100
         TabIndex        =   6
         Top             =   9720
         Value           =   1000
         Width           =   195
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Psteam , Tsteam qDryness to Hsteam"
         Height          =   435
         Left            =   1260
         TabIndex        =   5
         Top             =   5280
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   10
         Left            =   3900
         TabIndex        =   4
         Text            =   "1.000"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   8
         Left            =   2100
         TabIndex        =   3
         Text            =   "300.0 degC"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   7
         Left            =   300
         TabIndex        =   2
         Text            =   "40.0 Bar"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   11
         Left            =   0
         TabIndex        =   1
         Top             =   6360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   9
         Left            =   2580
         TabIndex        =   0
         Text            =   "xxxxx kJ/kg  xx %   xx.xxx kJ/kgdegC"
         Top             =   6360
         Visible         =   0   'False
         Width           =   4095
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   1
         Left            =   3960
         TabIndex        =   40
         Top             =   420
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   2
         Left            =   9120
         TabIndex        =   41
         Top             =   420
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump Pump2 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   42
         Top             =   3420
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrLineWidth  =   3
      End
      Begin PumpPrj.Pump Pump2 
         Height          =   495
         Index           =   2
         Left            =   9300
         TabIndex        =   43
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   0
         Left            =   6720
         TabIndex        =   44
         Top             =   4140
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         TopColour       =   32768
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ColourOpen      =   32768
      End
      Begin PumpPrj.Pump Pump2 
         Height          =   495
         Index           =   3
         Left            =   13740
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   3
         Left            =   4800
         TabIndex        =   46
         Top             =   420
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   4
         Left            =   8280
         TabIndex        =   47
         Top             =   420
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Port Boiler"
         Height          =   195
         Left            =   9780
         TabIndex        =   84
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Stbd Boiler"
         Height          =   195
         Left            =   660
         TabIndex        =   83
         Top             =   0
         Width           =   2895
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         BorderWidth     =   4
         X1              =   6900
         X2              =   6900
         Y1              =   5280
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   4
         X1              =   9720
         X2              =   3780
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label F_evap 
         Caption         =   "Fevap"
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   82
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label V_pos 
         Caption         =   "Vpos = xx %"
         Height          =   255
         Left            =   6960
         TabIndex        =   81
         Top             =   4860
         Width           =   1875
      End
      Begin VB.Label F_evap 
         Caption         =   "Fevap"
         Height          =   255
         Index           =   2
         Left            =   9780
         TabIndex        =   80
         Top             =   2040
         Width           =   2475
      End
      Begin VB.Label F_steam 
         Alignment       =   2  'Center
         Caption         =   "Fsteam(2)"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   79
         Top             =   960
         Width           =   2475
      End
      Begin VB.Label F_steam 
         Alignment       =   2  'Center
         Caption         =   "Fsteam(1)"
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   78
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label F_steam 
         Alignment       =   2  'Center
         Caption         =   "Fsteam(3)"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   77
         Top             =   1320
         Width           =   2475
      End
      Begin VB.Label V_attemp_pos 
         Alignment       =   2  'Center
         Caption         =   "V_attemp_pos(1) = xx %"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   76
         Top             =   1680
         Width           =   2475
      End
      Begin VB.Label V_attemp_pos 
         Alignment       =   2  'Center
         Caption         =   "V_attemp_pos(2) = xx %"
         Height          =   255
         Index           =   2
         Left            =   7020
         TabIndex        =   75
         Top             =   1680
         Width           =   2475
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   15
         Left            =   7200
         TabIndex        =   74
         Top             =   10380
         Width           =   75
      End
      Begin VB.Label F_steam_stbd 
         Alignment       =   2  'Center
         Caption         =   "Fsteam(102)"
         Height          =   255
         Index           =   2
         Left            =   7020
         TabIndex        =   73
         Top             =   960
         Width           =   2475
      End
      Begin VB.Label F_steam_stbd 
         Alignment       =   2  'Center
         Caption         =   "Fsteam(103)"
         Height          =   255
         Index           =   3
         Left            =   7020
         TabIndex        =   72
         Top             =   1320
         Width           =   2475
      End
      Begin VB.Label F_steam_stbd 
         Alignment       =   2  'Center
         Caption         =   "Fsteam(101)"
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   71
         Top             =   0
         Width           =   2475
      End
      Begin VB.Label TQ_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   1
         Left            =   11520
         TabIndex        =   70
         Top             =   12420
         Width           =   615
      End
      Begin VB.Label R_pcent 
         Alignment       =   1  'Right Justify
         Caption         =   "1.000"
         Height          =   255
         Index           =   1
         Left            =   7260
         TabIndex        =   69
         Top             =   12420
         Width           =   615
      End
      Begin VB.Label R_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   68
         Top             =   12420
         Width           =   615
      End
      Begin VB.Label R_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   3
         Left            =   9060
         TabIndex        =   67
         Top             =   12420
         Width           =   615
      End
      Begin VB.Label R_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   4
         Left            =   9900
         TabIndex        =   66
         Top             =   12420
         Width           =   555
      End
      Begin VB.Label TQ_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   2
         Left            =   12300
         TabIndex        =   65
         Top             =   12420
         Width           =   615
      End
      Begin VB.Label TQ_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   3
         Left            =   13200
         TabIndex        =   64
         Top             =   12420
         Width           =   555
      End
      Begin VB.Label TQ_pcent 
         Alignment       =   2  'Center
         Caption         =   "xx %"
         Height          =   255
         Index           =   4
         Left            =   13920
         TabIndex        =   63
         Top             =   12360
         Width           =   615
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   5820
         Picture         =   "Boilers.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   7200
         Picture         =   "Boilers.frx":030A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "R   Bar seconds/ kg"
         Height          =   255
         Left            =   7260
         TabIndex        =   62
         Top             =   9780
         Width           =   3315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Torque kNm/Bar"
         Height          =   255
         Left            =   11340
         TabIndex        =   61
         Top             =   9720
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "HP Turbine (1)"
         Height          =   615
         Left            =   7320
         TabIndex        =   60
         Top             =   10140
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "HP Turbine (1)"
         Height          =   615
         Left            =   11340
         TabIndex        =   59
         Top             =   10080
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "HP Turbine (2)"
         Height          =   615
         Left            =   8160
         TabIndex        =   58
         Top             =   10140
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "HP Turbine (2)"
         Height          =   615
         Left            =   12180
         TabIndex        =   57
         Top             =   10080
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "LP Turbine (1)"
         Height          =   615
         Left            =   9000
         TabIndex        =   56
         Top             =   10140
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "LP Turbine (1)"
         Height          =   615
         Left            =   13020
         TabIndex        =   55
         Top             =   10080
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "LP Turbine (2)"
         Height          =   615
         Left            =   9840
         TabIndex        =   54
         Top             =   10140
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "LP Turbine (2)"
         Height          =   615
         Left            =   13860
         TabIndex        =   53
         Top             =   10080
         Width           =   735
      End
      Begin VB.Label Msteam6 
         Caption         =   "Msteam(6)"
         Height          =   255
         Index           =   0
         Left            =   6960
         TabIndex        =   52
         Top             =   2340
         Width           =   2535
      End
      Begin VB.Label Psteam14 
         Caption         =   "Psteam(14)"
         Height          =   255
         Left            =   10440
         TabIndex        =   51
         Top             =   11760
         Width           =   1275
      End
      Begin VB.Label Turb_K 
         Caption         =   "Turbine_K"
         Height          =   255
         Left            =   6240
         TabIndex        =   50
         Top             =   10920
         Width           =   1035
      End
      Begin VB.Label Turb_efficiency 
         Caption         =   "Turbine_efficiency"
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Top             =   11220
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "DIS-ABLED ( TEST ONLY )"
         Height          =   255
         Left            =   9360
         TabIndex        =   48
         Top             =   9480
         Width           =   3315
      End
   End
End
Attribute VB_Name = "Boilers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Burner_ON_OFF_Click(Index As Integer)



Burner_ON(Index) = Not Burner_ON(Index)
If Burner_ON(Index) Then
    Burner_ON_OFF(Index).BackColor = RGB(0, &HFF, 0)
Else
    Burner_ON_OFF(Index).BackColor = RGB(&HC0, &HC0, &HC0)
End If

' 7/7/99 SET in Steam_Sim.Dll
Call set_STEAM_logicals(Index, Burner_ON(Index))

End Sub

Private Sub Command1_Click()
Dim i As Integer


' ==== GAS Constants ====

' Universal Gas Constant R0 = 8.3143 kJ/ (kmol deg K )
'
' Molecular Mass for AIR = MMair = 28.96 kg/kmol
' Gas constant for AIR , Rair = R0/MMair = 8.3143/28.96 = 0.2871  kJ / (kg deg K )
Rair = 0.2871

' Molecular Mass for STEAM = MM_steam = 18.015 kg/kmol
' Gas constant for STEAM , Rsteam = R0/MM_steam = 8.3143/18.015 = 0.4615  kJ / (kg deg K )
Rsteam = 0.4615

' =========================


' 1/7/99
Turbine_K = 1#

' 21/6/99
For i = 1 To 4
    R_pcent(i).Caption = Format(Rscroll(i).Value * 0.001, "#0.000")
    TQ_pcent(i).Caption = Format(TQscroll(i).Value * 0.001, "#0.000")
Next i



Tamb_steam = 20#      ' Ambient air Temp deg C
Pamb_steam = 1.013    ' Ambient air Press Bar


R_HP(1) = 1.8333    ' Bar Seconds / kg
R_HP(2) = 0.5812    ' Bar Seconds / kg
R_LP(1) = 0.1389    ' Bar Seconds / kg
R_LP(2) = 0.02137   ' Bar Seconds / kg

' 21/6/99
R_turb(1) = R_HP(1)
R_turb(2) = R_HP(2)
R_turb(3) = R_LP(1)
R_turb(4) = R_LP(2)

TQ_HP(1) = 16.03
TQ_HP(2) = 40.03
TQ_LP(1) = 166.6
TQ_LP(2) = 798.47

TQ_turb(1) = TQ_HP(1)
TQ_turb(2) = TQ_HP(2)
TQ_turb(3) = TQ_LP(1)
TQ_turb(4) = TQ_LP(2)



'    R_HP_bleed = 17.4 Bar / ( 6.5% of 23.4kg/s ) = 17.4 / 1.521 = 11.44
R_HP_bleed = 11.44   ' Bar Seconds / kg

'   R_bleed = 3.8 Bar / ( 7% of 23.4kg/s ) = 3.8 / 1.638 = 2.32
R_bleed = 2.32   ' Bar Seconds / kg

'   R_LP_bleed = 0.55 Bar / ( 5% of 23.4kg/s ) = 0.55 / 1.17 = 0.47
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



Mwater(1) = 10000#  ' kg = 10 Tons ( occupies approx 10m^3 volume )
Mwater(101) = 10000#  ' kg = 10 Tons ( occupies approx 10m^3 volume )


'Vwater = 10#        ' m^3    ( 10000 kg occupies 10m^3 )

Msteam(1) = 0.0001
Psteam(1) = 0.0001     ' Water vapour pressure in the drum ( Bar )
Msteam(101) = 0.0001
Psteam(101) = 0.0001     ' Water vapour pressure in the drum ( Bar )


For i = 1 To 200

    Fsteam(i) = 0#         ' kg/sec
    Psteam(i) = 0.0001
    
    Twater(i) = 20#    ' deg C
    Tsteam(i) = 20#    ' deg C
        
    Tsteel(i) = 20#    ' deg C
    
    Qsteel(i) = 0#
    Qsteel_amb(i) = 0#
    
    Fwater(i) = 0#
    
Next i

' Sea Water Temp
    Twater(2) = 17.5   ' deg C
' 30/6/99  Condenser SW Outlet Temp
    Twater(3) = 17.5   ' deg C
    

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

End Sub

Private Sub Command10_Click()
Dim P1_steam As Single
Dim T1_steam As Single
Dim H1_steam As Single

Dim q1_dryness As Single
Dim T1_sat As Single
Dim hf_1 As Single
Dim hfg_1 As Single
Dim hg_1 As Single


P1_steam = Val(Text2(7).text)
T1_steam = Val(Text2(8).text)

q1_dryness = Val(Text2(10).text)

Call get_Tsat(P1_steam, T1_sat)

Call get_hf_hfg(P1_steam, hf_1, hfg_1)


hg_1 = hf_1 + hfg_1


Text2(11).text = "hf = " & Format(hf_1, "###0") & "  hfg = " & Format(hfg_1, "###0") & "  hg = " & Format(hg_1, "###0")


Call get_Hsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_Steam)

Text2(9).text = "H =" & Format(H1_steam, "###0.0") & " kJ/kg " & _
"  qDry=" & Format(q1_dryness, "###0.000") & "   CP_steam=" & Format(CP_Steam, "###0.000")

End Sub

Private Sub Command11_Click()

boil_run_flag = True

End Sub

Private Sub Command12_Click()

boil_run_flag = False

End Sub

Private Sub Command16_Click()

' 12/7/99
'RWJ_condense_form.Show
Main_Condenser.Show

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command22_Click()

'SIUnits = True
Mollier.Show

End Sub

Private Sub Command5_Click()


If boil_run_flag Then
    Timer1.Interval = 250
    Fast_time = False
Else
    status_flag = False
    Call set_STEAM_logicals(5, status_flag)
End If


End Sub

Private Sub Command7_Click()

If boil_run_flag Then
    Timer1.Interval = 10
    Fast_time = True
Else
    status_flag = True
    Call set_STEAM_logicals(5, status_flag)
End If


End Sub

Private Sub Form_Load()

' For Normal time simulation...
STEP = Timer1.Interval * 0.001


If Not I_AM_INSTRUCTOR And Not PTT Then
            ' I MUST be an ACT Terminal (Full Mission) so
            Command22.Visible = False
Else
            Command22.Visible = True
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Pump2_Click(Index As Integer)

Pump2(Index).PumpOn = Not Pump2(Index).PumpOn

End Sub

Private Sub Rscroll_Change(Index As Integer)

    R_pcent(Index).Caption = Format(Rscroll(Index).Value * 0.001, "#0.000")

End Sub

Private Sub Timer1_Timer()

' 7/7/99  RUNNING the Boiler as a DLL....
If Not boil_run_flag Then

    ' update_vvalve_LSS(syxxx,nValve,index1,index2)

  ' OLD LSS
   '  Call update_vvalve_LSS(Boilers, 16, 1, -1)    ' VV16
   '  Call update_vvalve_LSS(Boilers, 18, 2, -1)    ' VV18
     
     Call update_vvalve_LSS(Boilers, 438, 1, -1)
     
     
      ' 21/10/99
  '  Call update_vvalve_LSS(Boilers, 318, 2, -1)
     Call update_vvalve_LSS(Boilers, 468, 2, -1)
     
     Call update_vvalve_LSS(Boilers, 284, 3, -1)
     Call update_vvalve_LSS(Boilers, 283, 4, -1)
     
 
    F_evap(1).Caption = "Fevap(1) = " & Format(LSS_reals(318), "##0.0000")
' 8/12/99
'   F_evap(2).Caption = "Fevap(1) = " & Format(LSS_reals(319), "##0.0000")
    F_evap(2).Caption = "Fevap(2) = " & Format(LSS_reals(319), "##0.0000")
    
    V_pos.Caption = "V_ahead_pos = " & Format(LSS_reals(81), "#####0.00")
     
     F_steam(1).Caption = "Fsteam(1) = " & Format(LSS_reals(758), "#####0.00")
     F_steam(2).Caption = "Fsteam(2) = " & Format(LSS_reals(759), "#####0.00")
     F_steam(3).Caption = "Fsteam(3) = " & Format(LSS_reals(760), "#####0.00")
      
    F_steam_stbd(1).Caption = "Fsteam(101) = " & Format(LSS_reals(775), "#####0.00")
    F_steam_stbd(2).Caption = "Fsteam(102) = " & Format(LSS_reals(776), "#####0.00")
    F_steam_stbd(3).Caption = "Fsteam(103) = " & Format(LSS_reals(777), "#####0.00")
    
     
     V_attemp_pos(1).Caption = "V_attemp_pos(1) = " & Format(LSS_reals(786), "#####0.00")
     V_attemp_pos(2).Caption = "V_attemp_pos(2) = " & Format(LSS_reals(787), "#####0.00")
     

'   Boiler 1
    List1(0).clear
    List1(0).AddItem "Hsteam(5) = " & Format(LSS_reals(768), "#####0.00")
    List1(0).AddItem "Tsteam(5) = " & Format(LSS_reals(767), "#####0.00")
    List1(0).AddItem " "
    List1(0).AddItem "Msteam(6) = " & Format(LSS_reals(82), "#####0.00")
    List1(0).AddItem "Psteam(6) = " & Format(LSS_reals(788), "#####0.00")
    List1(0).AddItem "Hsteam(6) = " & Format(LSS_reals(789), "#####0.00")
    List1(0).AddItem "Tsteam(6) = " & Format(LSS_reals(790), "#####0.00")
    List1(0).AddItem " "
    List1(0).AddItem "Tamb = " & Format(LSS_reals(58), "#####0.00")
    List1(0).AddItem "Tsea = " & Format(LSS_reals(59), "#####0.00")
    List1(0).AddItem "Twater(3) " & Format(LSS_reals(326), "#####0.00")
    List1(0).AddItem "Fwater(3) " & Format(LSS_reals(327), "#####0.00")

    List1(1).clear
    List1(1).AddItem "Twater " & Format(LSS_reals(752), "#####0.00")
    List1(1).AddItem "Mwater " & Format(LSS_reals(753), "#####0.00")
    
    List2(1).clear
    List2(1).AddItem "Psteam " & Format(LSS_reals(754), "#####0.00")
    List2(1).AddItem "Tsteam " & Format(LSS_reals(755), "#####0.00")
    List2(1).AddItem "Msteam " & Format(LSS_reals(756), "#####0.00")
    List2(1).AddItem "Hsteam " & Format(LSS_reals(757), "#####0.00")
    List2(1).AddItem " "
    List2(1).AddItem "Mair " & Format(LSS_reals(761), "#####0.00")
    List2(1).AddItem "Pair " & Format(LSS_reals(762), "#####0.00")
    List2(1).AddItem "Tair " & Format(LSS_reals(763), "#####0.00")
   
'   Boiler 2
    List1(2).clear
    List1(2).AddItem "Twater " & Format(LSS_reals(769), "#####0.00")
    List1(2).AddItem "Mwater " & Format(LSS_reals(770), "#####0.00")
    
    List2(2).clear
    List2(2).AddItem "Psteam " & Format(LSS_reals(771), "#####0.00")
    List2(2).AddItem "Tsteam " & Format(LSS_reals(772), "#####0.00")
    List2(2).AddItem "Msteam " & Format(LSS_reals(773), "#####0.00")
    List2(2).AddItem "Hsteam " & Format(LSS_reals(774), "#####0.00")
    List2(2).AddItem " "
    List2(2).AddItem "Mair " & Format(LSS_reals(778), "#####0.00")
    List2(2).AddItem "Pair " & Format(LSS_reals(779), "#####0.00")
    List2(2).AddItem "Tair " & Format(LSS_reals(780), "#####0.00")
     
    List8.clear
    List8.AddItem "Torque_HP(1) " & Format(LSS_reals(793), "#####0.00")
    List8.AddItem "Torque_HP(2) " & Format(LSS_reals(794), "#####0.00")
    List8.AddItem "Torque_LP(1) " & Format(LSS_reals(795), "#####0.00")
    List8.AddItem "Torque_LP(2) " & Format(LSS_reals(796), "#####0.00")
    List8.AddItem " "
    List8.AddItem "Power_HP(1) " & Format(LSS_reals(797), "#####0.00")
    List8.AddItem "Power_HP(2) " & Format(LSS_reals(798), "#####0.00")
    List8.AddItem "Power_LP(1) " & Format(LSS_reals(799), "#####0.00")
    List8.AddItem "Power_LP(2) " & Format(LSS_reals(800), "#####0.00")
    
    List8.AddItem " "
    
    List8.AddItem "shaft rpm " & Format(LSS_reals(80), "#####0.00")
    
    Pshaft = LSS_reals(797) + LSS_reals(798) + LSS_reals(799) + LSS_reals(800)
    
    List8.AddItem "shaft power = " & Format(Pshaft, "##0.0") & " kW"
    List8.AddItem "shaft power = " & Format(Pshaft / 0.746, "##0.0") & " hp"
    
    ' 12/12/99
    'List8.AddItem " "
    'List8.AddItem "Pfeed_water " & Format(LSS_reals(791), "#####0.00")
    'List8.AddItem "Pcirc_water " & Format(LSS_reals(792), "#####0.00")
    
            For i = 1 To 5
            
                  Psteam(i + 9) = LSS_reals(59 + i)
                  Hsteam(i + 9) = LSS_reals(64 + i)
                  Tsteam(i + 9) = LSS_reals(69 + i)
                    qDry(i + 9) = LSS_reals(74 + i)
                
                '  20/7/99 Done in Mollier.frm..
                '    Pressure(i) = Psteam(i + 9)
                '    Enthalpy(i) = Hsteam(i + 9)
                ' Temperature(i) = Tsteam(i + 9)
                '     Dryness(i) = qDry(i + 9)
                
                
            Next i
            
            ' 5/10/99
            Fsteam(10) = LSS_reals(88)
            Fsteam(11) = LSS_reals(89)
            Fsteam(12) = LSS_reals(157)
            Fsteam(13) = LSS_reals(158)
            Fsteam(14) = LSS_reals(159)
            
     List3(1).clear
     List3(1).AddItem "           Psteam( )    Tsteam( )    Hsteam( )       Fsteam( )      qDry( ) "
 
 For i = 10 To 14
    List3(1).AddItem "(" & Format(i, "00") & ")   " & Format(Psteam(i), "000.00") _
    & " Bar    " & Format(Tsteam(i), "000.0") & Chr(176) & "C     " & Format(Hsteam(i), "0000") & " kJ/kg    " _
    & Format(Fsteam(i), "000.00") & " kg/s   " & Format(qDry(i), "##0.000")
  Next i
  
  
    ' 7/12/00
     Tsteam(18) = LSS_reals(814)
     Fsteam(17) = LSS_reals(815)
     List3(1).AddItem " "
     List3(1).AddItem "  Tsteam(18) = " & Format(Tsteam(18), "000.0") & Chr(176) & "C "
     List3(1).AddItem "  Fsteam(17) = " & Format(Fsteam(17), "000.0") & " kg/s"
            
     
End If


' 7/7/99  RUNNING the Boiler from VB....
If boil_run_flag Then
            

            For i = 1 To 5
                Pressure(i) = Psteam(i + 9)
                Enthalpy(i) = Hsteam(i + 9)
            'Global Entropy(5) As Single
            Temperature(i) = Tsteam(i + 9)
                Dryness(i) = qDry(i + 9)
            Next i
 

            ' 21/6/99
            R_HP(1) = R_turb(1) * Rscroll(1).Value * 0.001  '  1.000
            R_HP(2) = R_turb(2) * Rscroll(2).Value * 0.001  '  1.000
            R_LP(1) = R_turb(3) * Rscroll(3).Value * 0.001  '  1.097
            R_LP(2) = R_turb(4) * Rscroll(4).Value * 0.001  '  1.152
            
            TQ_HP(1) = TQ_turb(1) * TQscroll(1).Value * 0.001 '  1.000
            TQ_HP(2) = TQ_turb(2) * TQscroll(2).Value * 0.001 '  1.242
            TQ_LP(1) = TQ_turb(3) * TQscroll(3).Value * 0.001 '  1.112
            TQ_LP(2) = TQ_turb(4) * TQscroll(4).Value * 0.001 '  1.315
            
                        
        ' Port Boiler filling
         If Pump2(1).PumpOn Then
            Fwater(10) = 12#
            Twater(10) = Twater(1) - 2#
         Else
            Fwater(10) = 0#
            Twater(10) = 20#
         End If
         
        ' Stbd Boiler filling  ( Boffset = 100 )
         If Pump2(2).PumpOn Then
            Fwater(110) = 12#
            Twater(110) = Twater(1) - 2#
         Else
            Fwater(110) = 0#
            Twater(110) = 20#
         End If
         
        ' Sea Water Pump
         If Pump2(3).PumpOn Then
            Fwater(2) = 1228#  ' kg/sec
         Else
            Fwater(2) = 0#
         End If
         




        '  Stbd Boiler
        Call Boiler_sub(1)
        
       
        Psteam(2) = Psteam(1)
        Psteam(3) = Psteam(1)
        Psteam(4) = Psteam(1)
        Psteam(5) = Psteam(1)
        
        '  Port Boiler
        Call Boiler_sub(2)


        Psteam(102) = Psteam(101)
        Psteam(103) = Psteam(101)
        Psteam(104) = Psteam(101)
        Psteam(105) = Psteam(101)
        
       
        '  The PORT Boiler Steam mixes Steam with the STBD Boiler Steam...
                
        M_dot = Fsteam(1) + Fsteam(101) - Fsteam(6)
        Msteam(6) = Msteam(6) + M_dot * STEP
        
        
        ' 15/6/99 Port & Stbd Mixed Steam...
        Qsteel_amb(6) = (Tsteel(6) - Tamb_steam) * 0.4    ' Heat loss to ambient
        Qsteel(6) = -Qsteel_amb(6)                  ' Heat loss to ambient


        'Heater_mix_sub(Pact, Tstm_IN_a, Fstm_IN_a, Tstm_IN_b, Fstm_IN_b, Mstm, Tstm, Mmetal, Tmetal, Qmetal, _
        '       CQstm, Hstm, qDryness)

        Call Heater_mix_sub(Psteam(6), Tsteam(5), Fsteam(1), Tsteam(105), Fsteam(101), 10#, Tsteam(6), 500#, Tsteel(6), Qsteel(6), _
                    3.702, Hsteam(6), qDry(6))
                                             
                                             
     ' Psteam(6) oscillates abobe 50% Throttle....
     '  Psteam(6) = Msteam(6) * Rsteam * (Tsteam(6) + 273.15) / 0.5        '  from PV = mRT , P = mRT/V
      ' Psteam(6) = Msteam(6) * Rsteam * (Tsteam(6) + 273.15) / 5#         '  from PV = mRT , P = mRT/V
        Psteam(6) = Msteam(6) * Rsteam * (Tsteam(6) + 273.15) / 10#         '  from PV = mRT , P = mRT/V
        
         
        ' Psteam(6) in BAR...
        Psteam(6) = Psteam(6) * 0.01
                       
         
        '   Throttling Valve ( V_ahead_pos % )
            If V_ahead_pos = 100 Then
                    RV_ahead = 0.0000001
            Else
                    RV_ahead = 2.5 / (V_ahead_pos * 0.01 + 0.0000001) - 2.5
            End If
            
            '   There is NO Enthalpy drop across the throttling valve ( V_ahead_pos ).....
         
                Hsteam(10) = Hsteam(6)
                Tsteam(10) = Tsteam(6)
                qDry(10) = qDry(6)
                
                
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

                ' Normally   R_turbines = 1.8333 + 0.6865 = 2.5198
                R_turbines = R_HP(1) + R_turbine_C

                
                    Pdiff = Psteam(6) - Psteam(14)
                    
                    Rtotal = RV_ahead + R_turbines
                    

                    Psteam(11) = R_turbine_C * Pdiff / Rtotal + Psteam(14)
                    Psteam(12) = R_turbine_B * (Psteam(11) - Psteam(14)) / Res_B + Psteam(14)
                    Psteam(13) = R_turbine_A * (Psteam(12) - Psteam(14)) / Res_A + Psteam(14)
                                                                      
                    Psteam(10) = R_turbines * Pdiff / Rtotal + Psteam(14)
                   
                                     
                    Fsteam(6) = Pdiff / Rtotal
                    
                    Fsteam(10) = Fsteam(6)                              ' HP Turbine First stage
                    Fsteam(11) = (Psteam(11) - Psteam(12)) / R_HP(2)    ' HP Turbine Second stage
                    Fsteam(12) = (Psteam(12) - Psteam(13)) / R_LP(1)    ' LP Turbine First stage
                    Fsteam(13) = (Psteam(13) - Psteam(14)) / R_LP(2)    ' LP Turbine Second stage
                    
                    ' 17/6/99
                    Fsteam(14) = (Psteam(11) - Psteam(14)) / R_HP_bleed   ' HP Turbine Bled Steam
                    Fsteam(15) = (Psteam(12) - Psteam(14)) / R_bleed      ' Turbine Cross-over Bled Steam
                    Fsteam(16) = (Psteam(13) - Psteam(14)) / R_LP_bleed   ' LP Turbine Bled Steam
                    
                              
               
         '    Turbine_sub(Pin, Pout, Turbo_K, Fsteam, Hin, Hout, Turbo_Torque, Turbo_power, Tout, qDry_out)
         '               (IN , IN  ,    IN  ,    IN ,  IN,  OUT,          OUT,         OUT,  OUT,    OUT  )
         
         '  ===  HP TURBINE FIRST STAGE ===
         
           Turbine_efficiency = 1#
       
         '  Torque_HP(1) = 7200kW / 100 rpm = 7200 kJ/s / (100*2*Pi/60) rads/s
         '  = 7200 kNm/s / 10.47 rads/s = 687.7 kNm
         
         '  Turbo_K = 687.7 kNm / ( 60.3Bar-17.4Bar)   =  TQ_HP(1) = 16.03 kNm/Bar
         
          Call Turbine_sub(Psteam(10), Psteam(11), TQ_HP(1), Fsteam(10), Hsteam(10), Hsteam(11), Torque_HP(1), Power_HP(1), Tsteam(11), qDry(11))
            
            
         '  ===  HP TURBINE SECOND STAGE ===
         
         
         '  Torque_HP(2) = 5700kW / 100 rpm = 5700 kJ/s / (100*2*Pi/60) rads/s
         '  = 5700 kNm/s / 10.47 rads/s = 544.4 kNm
         
         '  Turbo_K = 544.4 kNm / ( 17.4 Bar - 3.8 Bar)  =  TQ_HP(2)  =  40.03 kNm/Bar
         
          Call Turbine_sub(Psteam(11), Psteam(12), TQ_HP(2), Fsteam(11), Hsteam(11), Hsteam(12), Torque_HP(2), Power_HP(2), Tsteam(12), qDry(12))
            
            
         '  ===  LP TURBINE FIRST STAGE ===
         
         ' 1/7/99
         Turbine_efficiency = (Turbine_K + 0.047 - Psteam(14)) / Turbine_K
         If Turbine_efficiency < 0.3 Then Turbine_efficiency = 0.3
         
         
         '  Torque_LP(1) = 5300kW / 100 rpm = 5300 kJ/s / (100*2*Pi/60) rads/s
         '  = 5300 kNm/s / 10.47 rads/s = 506.2 kNm
         
         '  Turbo_K = 506.2 kNm / ( 3.8 Bar - 0.55 Bar)  =  TQ_LP(1) =  155.75 kNm/Bar
          Call Turbine_sub(Psteam(12), Psteam(13), TQ_LP(1), Fsteam(12), Hsteam(12), Hsteam(13), Torque_LP(1), Power_LP(1), Tsteam(13), qDry(13))
                        
                        
          '  ===  LP TURBINE SECOND STAGE ===
         
         '  Torque_LP(2) = 4180kW / 100 rpm = 4180 kJ/s / (100*2*Pi/60) rads/s
         '  = 4180 kNm/s / 10.47 rads/s = 399.24 kNm
         
         '  Turbo_K = 399.24 kNm / ( 0.55 Bar - 0.05 Bar)    =  TQ_LP(2) =  798.47 kNm/Bar
         
           Call Turbine_sub(Psteam(13), Psteam(14), TQ_LP(2), Fsteam(13), Hsteam(13), Hsteam(14), Torque_LP(2), Power_LP(2), Tsteam(14), qDry(14))
                        
            
                       
'===========30/6/99 ===================

'    ====   Main Condenser ====

'  Note:
'   At 0.047 Bar there is VERY LITTLE Steam Mass in the condenser ( e.g 0.1 kg )
'   To calculate ACTUAL steam mass in the condeser when approx 24 kg of steam is flowing in
'   and approx 24 kg of water is flowing to the Hotwell would result in un-stable
'   differential equations. (You would need a VERY small timer interval to resolve this )

' So use a different approach viz:

                
'   Condense_sub(Pact, Mwat, Mwat_norm, Twat, Fwat_out, _
'                Mstm, Mstm_norm, Tstm_IN, Fstm_IN, Tstm, Mmetal, Tmetal, Qmetal, _
'               CQwat, CQstm, Cevap, F_evap, _
'                Hstm, qDryness)
                            
 '   Call Condense_sub(Psteam(14), Mwater(2), 10000#, Twater(2), 0#, 27#, _
 '           Msteam(15), 60#, Tsteam(15), 10000#, Tsteel(15), 0#, _
 '            1#, 1#, 100#, Fevap(3), _
 '           Hsteam(15), qDry(15))
 
            ' 30/6/99
            ' Finf Psat for Tsteam(14) ( LP turbine exhaust )
         '   Call get_Psat(Psat, Tsteam(14))
         '   Call get_hf_hfg(Psat, hf, hfg)
         '   Qcondenser = (Hsteam(14) - hf) * Fsteam(13)   ' Heat transfer to Sea Water
         
            '   Heat transfer to Sea Water 50,000kW when Tsteam(14) = 33 degC , Twater(3) = 27.5 degC
             Qcondenser = (Tsteam(14) - Twater(3)) * 9091#  ' Heat transfer to Sea Water
             
            If (Fsteam(13) > 0.1) Then
            
                dEnth_possible = Qcondenser / Fsteam(13)   ' Enthlapy drop possible due to SW cooling
                
                ' Find Psat , hf and hfg for Tsteam(14) ( LP turbine exhaust )
                Call get_Psat(Psat, Tsteam(14))
                Call get_hf_hfg(Psat, hf, hfg)
                dEnth_required = Hsteam(14) - hf    ' Enthalpy required to condense ALL the steam
                
                If (dEnth_possible >= dEnth_required) Then
                    dEnth_actual = dEnth_required       ' ALL the steam is condensed
                Else
                    dEnth_actual = dEnth_possible       ' Some of the steam is condensed
                End If
                
                  '  Psteam14_set = Psteam(13) - dEnth_actual / 2083#   ' Went to 16 Bar??
                     Psteam14_set = 1# - dEnth_actual / 2083#
                     If Psteam14_set < 0# Then Psteam14_set = 0#
                
            Else
            
                    dEnth_actual = 0#
                    Psteam14_set = 0#
                
            End If
            
                         Qcondenser = dEnth_actual * Fsteam(13)
                         
                   Hsteam(17) = Hsteam(14) - dEnth_actual  ' Enthalpy IN the condenser
                   
            ' Condensed steam
            '    T_dot = (-Qcondenser / CP_Steam + Fsteam(13) * (Tsteam(14) - Tsteam(17))) / 100000#
            '    Tsteam(17) = Tsteam(17) + T_dot * STEP
                   
                                  
               P_dot = (Psteam14_set - Psteam(14)) * 0.1
               Psteam(14) = Psteam(14) + P_dot * STEP

            '   Call get_Tsat(Psteam(14), Tsteam(14))
                Call get_Tsat(Psteam(14), Tsteam(17))
                
                
 
            ' Sea Water Cooling....
            ' Sea Water Desity = 1024 kg/m^3   ( see Marine Engines and Turbines Page 260 )
            ' CPsea = 3.935 kJ / kg degC
            ' Flow = 1.2 m^3/sec = 1.2 m^3/sec * 1024 kg/m^3 = 1228 kg/sec

            ' Sea Water Temp = Twater(2)
            
            T_dot = (Qcondenser / 3.935 + Fwater(2) * (Twater(2) - Twater(3))) / 5000#
            Twater(3) = Twater(3) + T_dot * STEP
 

'======================================
            
            
            ' PROP SHAFT TORQUE kNm   ( 2136.8 kNm at 22,380kW , 100rpm )
            TQ_shaft = Torque_HP(1) + Torque_HP(2) + Torque_LP(1) + Torque_LP(2)
            
            ' PROP SHAFT LOAD TORQUE kNm  ( 2136.8 kNm at 22,380kW , 100rpm )
         '  TQ_shaft_load = shaft_rpm * shaft_rpm * shaft_rpm * 2.1368 * 0.001
         
            TQ_shaft_load = Shaft_rpm * Shaft_rpm * 0.21368
            
            ' 9/6/99 Define a MINIMUM load torque
            If TQ_shaft_load < 100# Then
                If Shaft_rpm > 0.1 Then
                    TQ_shaft_load = 100#
                End If
            End If
            
            
            ' Angular Acceleration = Torque / Moment of Inertia
            shaft_rpm_dot = (TQ_shaft - TQ_shaft_load) / 4000#
            
            
            Shaft_rpm = Shaft_rpm + shaft_rpm_dot * STEP
            
                                                 
   
        Fout_air = 0#
        Fout_steam = 0#
        
'        Fout_pump = 0#
'        Fout1 = 0#

        Fsteam(1) = 0#
                
  
           
' Port Boiler
    If ((valve_status(1) = 1) And (Pactual(1) > Psteam(6))) Then
    
           ' 16/6/99
           '    Fout_air = (Pair(1) - Psteam(6)) / 2#
                Fout_air = Pair(1) / 2#
                
          ' 16/6/99
          ' 87.5 Bar at boiler , 6269 Bar at HP first Turbine   Throttle 100% )....
          '   Fout_steam = (Psteam(1) - Psteam(6)) / 2#
           If Psteam(1) > Psteam(6) Then
              Fout_steam = (Psteam(1) - Psteam(6)) * 3#      ' e.g 64 Bar - 60 Bar * 3.0 = 12 kg/sec
              
              ' 16/6/99 CAN CRASH IF Psteam(1) = 30 Bar and Psteam(6)=0 Bar  SO....
              If Fout_steam > 15# Then Fout_steam = 15#
              
           Else
              Fout_steam = 0#
           End If
           
               Fsteam(1) = Fout_air + Fout_steam

    Else
                Fout_air = 0#
              Fout_steam = 0#
               Fsteam(1) = 0#
    End If
            ' AIR Loss
            M_dot = Fout_air
            Mair(1) = Mair(1) - M_dot * STEP
        
            ' Steam Loss
            M_dot = Fout_steam
            Msteam(1) = Msteam(1) - M_dot * STEP
            
' Stbd Boiler
    If ((valve_status(2) = 1) And (Pactual(2) > Psteam(6))) Then
    
           ' 16/6/99
          '     Fout_air = (Pair(2) - Psteam(6)) / 2#
                Fout_air = Pair(2) / 2#
                
          ' 16/6/99
          ' 87.5 Bar at boiler , 6269 Bar at HP first Turbine   Throttle 100% )....
          '   Fout_steam = (Psteam(101) - Psteam(6)) / 2#
           If Psteam(101) > Psteam(6) Then
              Fout_steam = (Psteam(101) - Psteam(6)) * 3#      ' e.g 64 Bar - 60 Bar * 3.0 = 12 kg/sec
              
              ' 16/6/99 CAN CRASH IF Psteam(1) = 30 Bar and Psteam(6)=0 Bar  SO....
              If Fout_steam > 15# Then Fout_steam = 15#
              
           Else
              Fout_steam = 0#
           End If
           
               Fsteam(101) = Fout_air + Fout_steam

    Else
                Fout_air = 0#
              Fout_steam = 0#
               Fsteam(101) = 0#
    End If
            ' AIR Loss
            M_dot = Fout_air
            Mair(2) = Mair(2) - M_dot * STEP
        
            ' Steam Loss
            M_dot = Fout_steam
            Msteam(101) = Msteam(101) - M_dot * STEP
            
'==================
  
  
    ' Throttling Valve
    'Valve1(0).ColourOpen = RGB(0, V_ahead_pos * 2.5, 0)
    If V_ahead_pos = 0 Then
        Valve(0).ColourOpen = RGB(0, 0, 0)                         ' BLACK
    Else
        Valve(0).ColourOpen = RGB(0, V_ahead_pos * 1.25 + 125, 0)  ' SHADE OF GREEN
    End If


    ' Port Steam valve
    If valve_status(1) = 1 Then
        Valve(1).OpenTopLeft = True
    Else
        Valve(1).OpenTopLeft = False
    End If

    ' Stbd Steam valve
    If valve_status(2) = 1 Then
        Valve(2).OpenBtmRight = True
    Else
        Valve(2).OpenBtmRight = False
    End If
  
  
    
    ' === DISPLAYS ===

                Fast_time_count = Fast_time_count + 1
                
 If Not Fast_time Or Fast_time_count > 12 Then

            Fast_time_count = 1

        
        Turb_efficiency.Caption = "T_effic = " & Format(Turbine_efficiency, "##0.00")
 
 
    V_pos.Caption = "V_ahead_pos = " & Format(V_ahead_pos, "##0.0") & " %"
    
  
     F_steam(1).Caption = "Fsteam(1) = " & Format(Fsteam(1), "##0.0000") & " kg/sec"
     
     F_steam(2).Caption = "Fsteam(2) = " & Format(Fsteam(2), "##0.0000") & " kg/sec"
     F_steam(3).Caption = "Fsteam(3) = " & Format(Fsteam(3), "##0.0000") & " kg/sec"
      
    F_steam_stbd(1).Caption = "Fsteam(101) = " & Format(Fsteam(101), "##0.0000") & " kg/sec"
    F_steam_stbd(2).Caption = "Fsteam(102) = " & Format(Fsteam(102), "##0.0000") & " kg/sec"
    F_steam_stbd(3).Caption = "Fsteam(103) = " & Format(Fsteam(103), "##0.0000") & " kg/sec"
    
     
     V_attemp_pos(1).Caption = "V_attemp_pos(1) = " & Format(Vattemp_POS(1), "##0.0") & " %"
     V_attemp_pos(2).Caption = "V_attemp_pos(2) = " & Format(Vattemp_POS(2), "##0.0") & " %"
     
     
 
  '  F_out_pump.Caption = "Fout_pump = " & Format(Fout_pump, "##0.0000") & " kg/sec"

    
  '  T_amb.Caption = "Tamb_steam = " & Format(Tamb_steam, "##0.0") & Chr(176) & "C"
  '  P_amb.Caption = "Pamb_steam = " & Format(Pamb_steam, "##0.000") & " Bar"
  
  
            List1(0).clear
            
         '   List1(0).AddItem Format(R_HP(1), "#0.0000")
         '   List1(0).AddItem Format(R_HP(2), "#0.0000")
         '   List1(0).AddItem Format(R_LP(1), "#0.0000")
         '   List1(0).AddItem Format(R_LP(2), "#0.0000")
            
         '   List1(0).AddItem Format(TQ_HP(1), "###0.00")
         '   List1(0).AddItem Format(TQ_HP(2), "###0.00")
         '   List1(0).AddItem Format(TQ_LP(1), "###0.00")
         '   List1(0).AddItem Format(TQ_LP(2), "###0.00")
            
            List1(0).AddItem Format(Qsteel(1), "###0.0") & " kW"
            List1(0).AddItem Format(Qsteel(2), "###0.0") & " kW"
            List1(0).AddItem Format(Qsteel(3), "###0.0") & " kW"
            List1(0).AddItem Format(Qsteel(101), "###0.0") & " kW"
            List1(0).AddItem Format(Qsteel(102), "###0.0") & " kW"
            List1(0).AddItem Format(Qsteel(103), "###0.0") & " kW"
            
            
            
     List1(0).AddItem " "
            
     List1(0).AddItem "Msteam(6) = " & Format(Msteam(6), "##0.000") & " kg"
    List1(0).AddItem "Psteam(6) = " & Format(Psteam(6), "##0.00") & " Bar"
    List1(0).AddItem "Tsteam(6) = " & Format(Tsteam(6), "##0.0") & Chr(176) & "C"
  
                                    
    List1(1).clear
    List1(1).AddItem "Mwater(1) = " & Format(Mwater(1), "##0.0") & " kg"
    List1(1).AddItem "Twater(1) = " & Format(Twater(1), "##0.0") & Chr(176) & "C"
    
   ' List1(1).AddItem "        Psat = " & Format(Psat, "##0.00") & " Bar"
   ' Call get_Tsat(Psat, Tsat)
   ' List1(1).AddItem "        Tsat = " & Format(Tsat, "##0.0") & Chr(176) & "C"
    
    List1(2).clear
    List1(2).AddItem "Mwater(101) = " & Format(Mwater(101), "##0.0") & " kg"
    List1(2).AddItem "Twater(101) = " & Format(Twater(101), "##0.0") & Chr(176) & "C"
    
   ' List1(2).AddItem "        Psat = " & Format(Psat, "##0.00") & " Bar"
   ' Call get_Tsat(Psat, Tsat)
   ' List1(2).AddItem "        Tsat = " & Format(Tsat, "##0.0") & Chr(176) & "C"
    

    List2(1).clear
    List2(1).AddItem "Msteam(1) = " & Format(Msteam(1), "##0.000") & " kg"
    List2(1).AddItem "Psteam(1) = " & Format(Psteam(1), "##0.00") & " Bar"
    List2(1).AddItem "Tsteam(1) = " & Format(Tsteam(1), "##0.0") & Chr(176) & "C"
    List2(1).AddItem " "
    List2(1).AddItem "     Mair(1) = " & Format(Mair(1), "##0.000") & " kg"
    List2(1).AddItem "     Pair(1) = " & Format(Pair(1), "##0.00") & " Bar"
    List2(1).AddItem "     Tair(1) = " & Format(Tair(1), "##0.0") & Chr(176) & "C"
    
    List2(2).clear
    List2(2).AddItem "Msteam(101) = " & Format(Msteam(101), "##0.000") & " kg"
    List2(2).AddItem "Psteam(101) = " & Format(Psteam(101), "##0.00") & " Bar"
    List2(2).AddItem "Tsteam(101) = " & Format(Tsteam(101), "##0.0") & Chr(176) & "C"
    List2(2).AddItem " "
    List2(2).AddItem "     Mair(2) = " & Format(Mair(2), "##0.000") & " kg"
    List2(2).AddItem "     Pair(2) = " & Format(Pair(2), "##0.00") & " Bar"
    List2(2).AddItem "     Tair(2) = " & Format(Tair(2), "##0.0") & Chr(176) & "C"
    
    
    List3(1).clear
    List3(1).AddItem "           Psteam( )    Tsteam( )    Hsteam( )       Fsteam()       qDry( )    Tsteel( )"
  For i = 1 To 16
    List3(1).AddItem "(" & Format(i, "00") & ")   " & Format(Psteam(i), "000.00") _
    & " Bar    " & Format(Tsteam(i), "000.0") & Chr(176) & "C     " & Format(Hsteam(i), "0000") & " kJ/kg    " _
    & Format(Fsteam(i), "00.00") & " kg/s    " _
    & Format(qDry(i), "##0.000") & "   " _
    & Format(Tsteel(i), "000.0") & Chr(176) & "C"
  Next i
     
    List3(2).clear
   'List3(2).AddItem "           Psteam( )    Tsteam( )    Hsteam( )       qDry( )    Tsteel( )"
    List3(2).AddItem "           Psteam( )    Tsteam( )    Hsteam( )       Fsteam()       qDry( )    Tsteel( )"
  For i = 101 To 106
    List3(2).AddItem "(" & Format(i, "000") & ")   " & Format(Psteam(i), "000.00") _
    & " Bar    " & Format(Tsteam(i), "000.0") & Chr(176) & "C     " & Format(Hsteam(i), "0000") & " kJ/kg    " _
    & Format(Fsteam(i), "00.00") & " kg/s    " _
    & Format(qDry(i), "##0.000") & "   " _
    & Format(Tsteel(i), "000.0") & Chr(176) & "C"
  Next i
      
     
    List4.clear
    List4.AddItem "           Psteam( )    Tsteam( )    Hsteam( )       Fsteam()       qDry( )    Tsteel( )"
  For i = 13 To 17
    List4.AddItem "(" & Format(i, "000") & ")   " & Format(Psteam(i), "000.00") _
    & " Bar    " & Format(Tsteam(i), "000.0") & Chr(176) & "C     " & Format(Hsteam(i), "0000") & " kJ/kg    " _
    & Format(Fsteam(i), "00.00") & " kg/s    " _
    & Format(qDry(i), "##0.000") & "   " _
    & Format(Tsteel(i), "000.0") & Chr(176) & "C"
  Next i
    List4.AddItem "Twater(2) = " & Format(Twater(2), "##0.0") & Chr(176) & "C"
    List4.AddItem "Twater(3) = " & Format(Twater(3), "##0.0") & Chr(176) & "C"
    List4.AddItem "Psteam14_set = " & Format(Psteam14_set, "##0.000") & " Bar"
    
       
    List8.clear
    List8.AddItem "Torque_HP(1) = " & Format(Torque_HP(1), "##0.0") & " kNm"
    List8.AddItem "Torque_HP(2) = " & Format(Torque_HP(2), "##0.0") & " kNm"
    List8.AddItem "Torque_LP(1) = " & Format(Torque_LP(1), "##0.0") & " kNm"
    List8.AddItem "Torque_LP(2) = " & Format(Torque_LP(2), "##0.0") & " kNm"
    List8.AddItem " "
    List8.AddItem "Power_HP(1) = " & Format(Power_HP(1), "##0.0") & " kW"
    List8.AddItem "Power_HP(2) = " & Format(Power_HP(2), "##0.0") & " kW"
    List8.AddItem "Power_LP(1) = " & Format(Power_LP(1), "##0.0") & " kW"
    List8.AddItem "Power_LP(2) = " & Format(Power_LP(2), "##0.0") & " kW"
    List8.AddItem " "
    
    ' 22/6/99
   ' List8.AddItem "TQ_shaft_load = " & Format(TQ_shaft_load, "##0.0") & " kNm"
    
    List8.AddItem "shaft_rpm = " & Format(Shaft_rpm, "##0.0") & " rpm"
    
    Pshaft = Power_HP(1) + Power_HP(2) + Power_LP(1) + Power_LP(2)
    List8.AddItem "shaft power = " & Format(Pshaft, "##0.0") & " kW"
    List8.AddItem "shaft power = " & Format(Pshaft / 0.746, "##0.0") & " hp"
    
    
    
    
 '   T_steel(1).Caption = "Tsteel(1) = " & Format(Tsteel(1), "##0.0") & Chr(176) & "C"
   
 '   P_actual.Caption = "Pactual(1) = " & Format(Pactual(1), "##0.00") & " Bar"
    
    F_evap(1).Caption = "Fevap(1) = " & Format(Fevap(1), "##0.0000") & " kg/sec"
    
    F_evap(2).Caption = "Fevap(2) = " & Format(Fevap(2), "##0.0000") & " kg/sec"

 '   h_f.Caption = "hf = " & Format(hf, "####0") & " kJ/kg"
 '   h_fg.Caption = "hfg = " & Format(hfg, "####0") & " kJ/kg"
    
 '   h_g.Caption = "hg = " & Format(hg, "####0") & " kJ/kg"
    
    
    
 End If



End If


End Sub

Private Sub TQscroll_Change(Index As Integer)

    TQ_pcent(Index).Caption = Format(TQscroll(Index).Value * 0.001, "#0.000")

End Sub

Private Sub valve_Click(Index As Integer)


' OLD LSS
'Select Case Index
'
'Case 1
'' 15/7/99..
''    Call valve_pcdi_LSS(27)
'     Call valve_pcdi_LSS(16)
'Case 2
''    Call valve_pcdi_LSS(29)
'     Call valve_pcdi_LSS(18)
'
'End Select


'If valve_status(Index) = 0 Then
'    valve_status(Index) = 1
'Else
'    valve_status(Index) = 0
'End If


Select Case Index
  Case 1
        Call valve_pcdi_LSS(438)
  Case 2
 ' 21/10/99
 '      Call valve_pcdi_LSS(318)
        Call valve_pcdi_LSS(468)
  Case 3
        Call valve_pcdi_LSS(284)
  Case 4
        Call valve_pcdi_LSS(283)
End Select


End Sub

Private Sub VScroll1_Change(Index As Integer)

' 14/6/99
Vscroll_val(Index) = VScroll1(Index).Value

' 7/7/99
Call set_VSCROLL(Index, VScroll1(Index).Value)


End Sub

Private Sub VScroll2_Change()

Psteam(14) = VScroll2.Value * 0.001
Psteam14.Caption = "Psteam(14) = " & Format(Psteam(14), "##0.000")


End Sub

Private Sub VScroll3_Change()

V_ahead_pos = VScroll3.Value * 0.1

' 7/7/99
Call set_VSCROLL(3, VScroll3.Value)

End Sub

Private Sub VScroll4_Change()

Twater(2) = VScroll4.Value * 0.1

End Sub

Private Sub VScroll5_Change()


Turbine_K = VScroll5.Value * 0.01
Turb_K.Caption = Format(Turbine_K, "##0.00")


End Sub
