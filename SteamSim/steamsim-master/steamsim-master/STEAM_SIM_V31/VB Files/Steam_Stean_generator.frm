VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Stm_Stm_Generator 
   Caption         =   "Contaminated Steam Evaporator"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11760
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6225
      Left            =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10980
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         TabIndex        =   48
         Text            =   "P"
         Top             =   1560
         Width           =   210
      End
      Begin VB.TextBox P_stm_BUNK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   42
         Text            =   "120 psi"
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox P_stm_DECK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6180
         TabIndex        =   41
         Text            =   "120 psi"
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1740
         TabIndex        =   37
         Text            =   "P"
         Top             =   2400
         Width           =   210
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   480
         Top             =   4920
      End
      Begin VB.TextBox P_875_desup_out 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   36
         Text            =   "875 psi"
         Top             =   2160
         Width           =   795
      End
      Begin VB.TextBox P_225 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   31
         Text            =   "225 psi"
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3900
         TabIndex        =   27
         Text            =   "P"
         Top             =   1380
         Width           =   210
      End
      Begin VB.TextBox P_120 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   22
         Text            =   "120 psi"
         Top             =   1040
         Width           =   795
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 HavenValveV1_24 
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         Top             =   3270
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 HavenValveV1_23 
         Height          =   315
         Left            =   5460
         TabIndex        =   14
         Top             =   3270
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         TopColour       =   65280
         BottomColour    =   65280
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 HavenValveV1_22 
         Height          =   495
         Left            =   6720
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   873
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         BackColour      =   12632256
         ValveType       =   10
         ColourOpen      =   32768
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   580
         Left            =   4020
         TabIndex        =   11
         Top             =   1980
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   1
         Left            =   2910
         TabIndex        =   0
         Top             =   1830
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   2
         Left            =   5700
         TabIndex        =   1
         Top             =   1620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   975
         Index           =   21
         Left            =   6420
         TabIndex        =   2
         Top             =   2070
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         TopColour       =   16711680
         BottomColour    =   16711680
         LineWidth       =   3
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   0
         Left            =   5250
         TabIndex        =   9
         Top             =   1620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 HavenValveV1_25 
         Height          =   285
         Left            =   4020
         TabIndex        =   16
         Top             =   3270
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   581
         Left            =   5880
         TabIndex        =   29
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   582
         Left            =   7140
         TabIndex        =   30
         Top             =   1980
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   577
         Left            =   720
         TabIndex        =   46
         Top             =   2460
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   3
         Left            =   7720
         TabIndex        =   47
         Top             =   1980
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   16384
         BottomColour    =   16384
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
         ColourOpen      =   16384
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   591
         Left            =   8460
         TabIndex        =   49
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   30
         Left            =   10260
         TabIndex        =   51
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   593
         Left            =   8460
         TabIndex        =   52
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   31
         Left            =   10260
         TabIndex        =   54
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   190
         Left            =   8460
         TabIndex        =   55
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   32
         Left            =   10260
         TabIndex        =   57
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   189
         Left            =   8460
         TabIndex        =   58
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   33
         Left            =   10260
         TabIndex        =   60
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   554
         Left            =   8460
         TabIndex        =   61
         Top             =   2880
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   34
         Left            =   10260
         TabIndex        =   63
         Top             =   2880
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   590
         Left            =   8460
         TabIndex        =   64
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   35
         Left            =   10260
         TabIndex        =   66
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   599
         Left            =   8460
         TabIndex        =   67
         Top             =   4080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   36
         Left            =   10260
         TabIndex        =   69
         Top             =   4080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   573
         Left            =   8460
         TabIndex        =   70
         Top             =   4620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   37
         Left            =   10260
         TabIndex        =   72
         Top             =   4620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   592
         Left            =   8460
         TabIndex        =   74
         Top             =   5160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   5
         Left            =   10260
         TabIndex        =   75
         Top             =   5160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   375
         Left            =   6780
         TabIndex        =   77
         Top             =   4680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "FO Heater"
         Height          =   255
         Index           =   8
         Left            =   9060
         TabIndex        =   76
         Top             =   5340
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   8280
         X2              =   10920
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Line Line41 
         X1              =   6480
         X2              =   6480
         Y1              =   420
         Y2              =   660
      End
      Begin VB.Line Line40 
         X1              =   7620
         X2              =   7620
         Y1              =   2280
         Y2              =   2640
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   10800
         Picture         =   "Steam_Stean_generator.frx":0000
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   10920
         X2              =   10920
         Y1              =   5640
         Y2              =   780
      End
      Begin VB.Label Label21 
         Caption         =   "To CDT"
         Height          =   255
         Left            =   10500
         TabIndex        =   73
         Top             =   5940
         Width           =   765
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "SSTG LO"
         Height          =   255
         Index           =   7
         Left            =   9060
         TabIndex        =   71
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   8280
         X2              =   10920
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Slop Tk"
         Height          =   255
         Index           =   6
         Left            =   9060
         TabIndex        =   68
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   8280
         X2              =   10920
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "ME Sump"
         Height          =   255
         Index           =   5
         Left            =   9060
         TabIndex        =   65
         Top             =   3660
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   8280
         X2              =   10920
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "LO Purifier"
         Height          =   255
         Index           =   4
         Left            =   9060
         TabIndex        =   62
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   8
         X1              =   8280
         X2              =   10920
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Stbd Sett Tk"
         Height          =   255
         Index           =   3
         Left            =   9060
         TabIndex        =   59
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   8280
         X2              =   10920
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Port Sett Tk"
         Height          =   255
         Index           =   2
         Left            =   9060
         TabIndex        =   56
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   8280
         X2              =   10920
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Stbd Deep Tk"
         Height          =   255
         Index           =   1
         Left            =   9060
         TabIndex        =   53
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   8280
         X2              =   10920
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   6600
         Picture         =   "Steam_Stean_generator.frx":030A
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Port Deep Tk"
         Height          =   255
         Index           =   0
         Left            =   9060
         TabIndex        =   50
         Top             =   660
         Width           =   1095
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   8280
         X2              =   10920
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line38 
         BorderStyle     =   3  'Dot
         X1              =   7920
         X2              =   7920
         Y1              =   2100
         Y2              =   1680
      End
      Begin VB.Line Line37 
         BorderStyle     =   3  'Dot
         X1              =   7920
         X2              =   8280
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   8280
         X2              =   8280
         Y1              =   5460
         Y2              =   780
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   7020
         X2              =   8280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   6600
         Picture         =   "Steam_Stean_generator.frx":0614
         Stretch         =   -1  'True
         Top             =   540
         Width           =   255
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   5760
         X2              =   6840
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   4140
         X2              =   7020
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   7020
         X2              =   7020
         Y1              =   2250
         Y2              =   960
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4140
         X2              =   4140
         Y1              =   2250
         Y2              =   960
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5340
         X2              =   5760
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   360
         Picture         =   "Steam_Stean_generator.frx":091E
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   255
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   240
         X2              =   1200
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   1560
         X2              =   2160
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label18 
         Caption         =   "# 120"
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label17 
         Caption         =   "# 225"
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   2820
         Width           =   465
      End
      Begin VB.Label Label15 
         Caption         =   "# 875"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "# 225 Regulator"
         Height          =   495
         Left            =   960
         TabIndex        =   40
         Top             =   3180
         Width           =   855
      End
      Begin VB.Label REG_120_pos 
         Alignment       =   2  'Center
         Caption         =   "xxx %"
         Height          =   255
         Left            =   2160
         TabIndex        =   39
         Top             =   2940
         Width           =   495
      End
      Begin VB.Label Reg_875_225_pos 
         Alignment       =   2  'Center
         Caption         =   "xxx %"
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   2940
         Width           =   495
      End
      Begin VB.Line Line34 
         BorderStyle     =   3  'Dot
         X1              =   1800
         X2              =   1800
         Y1              =   2640
         Y2              =   2220
      End
      Begin VB.Line Line33 
         BorderStyle     =   3  'Dot
         X1              =   1440
         X2              =   1800
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line32 
         BorderStyle     =   3  'Dot
         X1              =   1395
         X2              =   1395
         Y1              =   2340
         Y2              =   2220
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   5760
         X2              =   5760
         Y1              =   960
         Y2              =   660
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         Height          =   255
         Left            =   4695
         TabIndex        =   28
         Top             =   1800
         Width           =   195
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   4560
         Shape           =   3  'Circle
         Top             =   1740
         Width           =   375
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   4740
         X2              =   4740
         Y1              =   1980
         Y2              =   2340
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "De-sup Steam"
         Height          =   735
         Left            =   0
         TabIndex        =   21
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Primary Steam Control"
         Height          =   495
         Left            =   1860
         TabIndex        =   20
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Direct Feed From Condensate"
         Height          =   675
         Left            =   5160
         TabIndex        =   19
         Top             =   4080
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "Blow Down Valve"
         Height          =   645
         Left            =   4680
         TabIndex        =   18
         Top             =   4110
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Scumming Valve"
         Height          =   465
         Left            =   3840
         TabIndex        =   17
         Top             =   4170
         Width           =   765
      End
      Begin VB.Line Line28 
         X1              =   4110
         X2              =   4110
         Y1              =   3270
         Y2              =   2670
      End
      Begin VB.Shape Shape3 
         Height          =   135
         Left            =   4050
         Shape           =   2  'Oval
         Top             =   2580
         Width           =   345
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4020
         Picture         =   "Steam_Stean_generator.frx":0C28
         Stretch         =   -1  'True
         Top             =   3690
         Width           =   255
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4140
         X2              =   4140
         Y1              =   3480
         Y2              =   4140
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4680
         Picture         =   "Steam_Stean_generator.frx":0F32
         Stretch         =   -1  'True
         Top             =   3690
         Width           =   255
      End
      Begin VB.Image Image31 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5460
         Picture         =   "Steam_Stean_generator.frx":123C
         Stretch         =   -1  'True
         Top             =   3690
         Width           =   255
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   3480
         Y2              =   4140
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5580
         X2              =   5580
         Y1              =   3360
         Y2              =   4020
      End
      Begin VB.Label Label3 
         Caption         =   "Steam To Deck"
         Height          =   255
         Left            =   6900
         TabIndex        =   13
         Top             =   600
         Width           =   1185
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   2880
         Picture         =   "Steam_Stean_generator.frx":1546
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   5940
         Picture         =   "Steam_Stean_generator.frx":1850
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   255
      End
      Begin VB.Line Line22 
         BorderStyle     =   3  'Dot
         X1              =   2400
         X2              =   4080
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "To Atmospheric Drain Tank"
         Height          =   795
         Left            =   2760
         TabIndex        =   10
         Top             =   4140
         Width           =   1065
      End
      Begin VB.Image Image23 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3120
         Picture         =   "Steam_Stean_generator.frx":1B5A
         Stretch         =   -1  'True
         Top             =   3750
         Width           =   255
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5760
         X2              =   5760
         Y1              =   2100
         Y2              =   1740
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3240
         X2              =   3240
         Y1              =   4290
         Y2              =   3090
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5340
         X2              =   5340
         Y1              =   2100
         Y2              =   1740
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3240
         X2              =   5790
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5790
         X2              =   5790
         Y1              =   3090
         Y2              =   2760
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3750
         X2              =   3450
         Y1              =   2340
         Y2              =   2190
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   5970
         X2              =   6450
         Y1              =   2700
         Y2              =   2820
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5970
         X2              =   6390
         Y1              =   2430
         Y2              =   2190
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5580
         X2              =   5580
         Y1              =   2100
         Y2              =   2280
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5760
         X2              =   5760
         Y1              =   1380
         Y2              =   1650
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   2820
         X2              =   3420
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5340
         X2              =   5940
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CONTAMINATED STEAM EVAPORATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2700
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   60
         Width           =   3330
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GAUGE GLASS"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6180
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   3300
         Width           =   990
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AIR VENT"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   2310
         Width           =   645
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Valves"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   5160
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   1140
         Width           =   960
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   1620
         Y2              =   1380
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4380
         X2              =   4380
         Y1              =   2790
         Y2              =   2910
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3900
         X2              =   4140
         Y1              =   3030
         Y2              =   3030
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3900
         X2              =   4380
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3900
         X2              =   3900
         Y1              =   3030
         Y2              =   2910
      End
      Begin VB.Image Image21 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2550
         Picture         =   "Steam_Stean_generator.frx":1E64
         Stretch         =   -1  'True
         Top             =   2070
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "9 inch"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6480
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "- 9 inc"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6450
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   3060
         Width           =   390
      End
      Begin VB.Line Line16 
         BorderStyle     =   3  'Dot
         X1              =   2355
         X2              =   2355
         Y1              =   2340
         Y2              =   1440
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2520
         X2              =   5790
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   3720
         Shape           =   4  'Rounded Rectangle
         Top             =   2310
         Width           =   2265
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   2160
         OleObjectBlob   =   "Steam_Stean_generator.frx":216E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   25
         Top             =   2340
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   2160
         OleObjectBlob   =   "Steam_Stean_generator.frx":A386
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   24
         Top             =   2340
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   2160
         OleObjectBlob   =   "Steam_Stean_generator.frx":12D9E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   23
         Top             =   2340
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   2160
         OleObjectBlob   =   "Steam_Stean_generator.frx":1B3B6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   26
         Top             =   2340
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   1200
         OleObjectBlob   =   "Steam_Stean_generator.frx":231CE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   34
         Top             =   2340
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   1200
         OleObjectBlob   =   "Steam_Stean_generator.frx":2B7E6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   33
         Top             =   2340
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   1200
         OleObjectBlob   =   "Steam_Stean_generator.frx":341FE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   32
         Top             =   2340
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   1200
         OleObjectBlob   =   "Steam_Stean_generator.frx":3C416
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   35
         Top             =   2340
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   360
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Stm_Stm_Generator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


On Error Resume Next
Call JYAScreenInit(Me, id_STM_STM)


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_STM_STM)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
' === 27/08/02 ===
Dim green_colour As Integer
Dim i_count As Integer


' === 27/08/02 ===
' Simple model for this "new" valve...
i_count = 0
If LSS_reals(150) < 40# / 14.5 Then
    green_colour = 255
Else
    If vv_status_LSS(591) = 3 Then i_count = i_count + 1
    If vv_status_LSS(593) = 3 Then i_count = i_count + 1
    If vv_status_LSS(190) = 3 Then i_count = i_count + 1
    If vv_status_LSS(189) = 3 Then i_count = i_count + 1
    If vv_status_LSS(554) = 3 Then i_count = i_count + 1
    If vv_status_LSS(590) = 3 Then i_count = i_count + 1
    If vv_status_LSS(599) = 3 Then i_count = i_count + 1
    If vv_status_LSS(573) = 3 Then i_count = i_count + 1
    If vv_status_LSS(592) = 3 Then i_count = i_count + 1
    green_colour = 40 + i_count * 20
End If
    Valve(3).ColourOpen = RGB(0, green_colour, 0)
    
'c      === 27/08/02 Z_stm_stm ( 0.0 to 1.0 ) ===
'           PC_REALS(858) = Z_stm_stm
If LSS_reals(858) > 1# Then
    Valve(21).LineWidth = 50
ElseIf LSS_reals(858) < 0# Then
    Valve(21).LineWidth = 0
Else
    Valve(21).LineWidth = LSS_reals(858) * 50#
End If


' 8/8/00

If SI_units Then

    P_875_desup_out.text = Format(LSS_reals(363) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    P_225.text = Format(LSS_reals(145) * Press_mult, "##0.0") & Press_chars     ' Bar
    P_120.text = Format(LSS_reals(147) * Press_mult, "##0.0") & Press_chars     ' Bar
    P_stm_DECK.text = Format(LSS_reals(149) * Press_mult, "##0.0") & Press_chars     ' Bar
    P_stm_BUNK.text = Format(LSS_reals(150) * Press_mult, "##0.0") & Press_chars     ' Bar

Else
    
    P_875_desup_out.text = Format(LSS_reals(363) * Press_mult, "###0") & Press_chars       ' psi
    P_225.text = Format(LSS_reals(145) * Press_mult, "###0") & Press_chars       ' psi
    P_120.text = Format(LSS_reals(147) * Press_mult, "###0") & Press_chars       ' psi
    P_stm_DECK.text = Format(LSS_reals(149) * Press_mult, "###0") & Press_chars       ' psi
    P_stm_BUNK.text = Format(LSS_reals(150) * Press_mult, "###0") & Press_chars       ' psi

End If


Reg_875_225_pos.Caption = Format(LSS_reals(146), "###0") & " %"
Call update_regulator(146, Reg_875_225(0), Reg_875_225(1), Reg_875_225(2), Reg_875_225(3))

REG_120_pos.Caption = Format(LSS_reals(148), "###0") & " %"
Call update_regulator(148, Reg_120(0), Reg_120(1), Reg_120(2), Reg_120(3))

'c       === 8/8/00 ===
'          PC_REALS(145) = P_225
'          PC_REALS(146) = REG_875_225_pos
'          PC_REALS(147) = P_120
'          PC_REALS(148) = REG_120_pos
'          PC_REALS(149) = P_stm_DECK
'          PC_REALS(150) = P_stm_BUNK



For i = 580 To 582
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

' 14/9/00   See Point 8.85 NEW VALVE...
   Call update_vvalve_LSS(Me, 577, 577, -1)

' === 19/08/02 ===
   Call update_vvalve_LSS(Me, 189, 189, -1)
   Call update_vvalve_LSS(Me, 190, 190, -1)
   Call update_vvalve_LSS(Me, 554, 554, -1)
   Call update_vvalve_LSS(Me, 590, 590, -1)
   
 ' === 20/08/02 Save valves...
'For i = 591 To 599
'   Call update_vvalve_LSS(Me, i, i, -1)
'Next i
'
'For i = 572 To 574
'   Call update_vvalve_LSS(Me, i, i, -1)
'Next i
   Call update_vvalve_LSS(Me, 591, 591, -1)
   Call update_vvalve_LSS(Me, 593, 593, -1)
   Call update_vvalve_LSS(Me, 599, 599, -1)
   Call update_vvalve_LSS(Me, 573, 573, -1)
   
   '=== 23/08/02 ===
   Call update_vvalve_LSS(Me, 592, 592, -1)
   
   Call copy_valve(Me, 591, 30)
   Call copy_valve(Me, 593, 31)
   Call copy_valve(Me, 190, 32)
   Call copy_valve(Me, 189, 33)
   Call copy_valve(Me, 554, 34)
   Call copy_valve(Me, 590, 35)
   Call copy_valve(Me, 599, 36)
   Call copy_valve(Me, 573, 37)
   
   '=== 23/08/02 ===
   Call copy_valve(Me, 592, 5)


End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
  
' 14/9/00

'=== 19/08/02 ===
'If Index > 576 Then
If Index > 188 Then
    Call valve_pcdi_LSS(Index)
End If

If Index = 30 Then
    Call valve_pcdi_LSS(591)
ElseIf Index = 31 Then
    Call valve_pcdi_LSS(593)
ElseIf Index = 32 Then
    Call valve_pcdi_LSS(190)
ElseIf Index = 33 Then
    Call valve_pcdi_LSS(189)
ElseIf Index = 34 Then
    Call valve_pcdi_LSS(554)
ElseIf Index = 35 Then
    Call valve_pcdi_LSS(590)
ElseIf Index = 36 Then
    Call valve_pcdi_LSS(599)
ElseIf Index = 37 Then
    Call valve_pcdi_LSS(573)
    
' === 23/08/02 ===
ElseIf Index = 5 Then
    Call valve_pcdi_LSS(592)
End If




End Sub
Sub update_regulator(nReals As Integer, obj_0 As Object, obj_1 As Object, obj_2 As Object, obj_3 As Object)

' 6/11/99
If LSS_reals(nReals) > 99# Then
      If (obj_3.Visible = False) Then
          obj_0.Visible = False
          obj_1.Visible = False
          obj_2.Visible = False
          obj_3.Visible = True
      End If
ElseIf LSS_reals(nReals) > 49# Then
      If (obj_2.Visible = False) Then
          obj_0.Visible = False
          obj_1.Visible = False
          obj_2.Visible = True
          obj_3.Visible = False
      End If
ElseIf LSS_reals(nReals) > 0.5 Then
      If (obj_1.Visible = False) Then
          obj_0.Visible = False
          obj_1.Visible = True
          obj_2.Visible = False
          obj_3.Visible = False
      End If
Else
      If (obj_0.Visible = False) Then
          obj_0.Visible = True
          obj_1.Visible = False
          obj_2.Visible = False
          obj_3.Visible = False
      End If
End If



End Sub

