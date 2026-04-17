VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Port_Boiler 
   Caption         =   "Port Boiler"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "Port_Boiler_V2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   11835
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9315
      Left            =   0
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   16431
      Begin VB.TextBox F_MAIN_stm 
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
         Height          =   285
         Left            =   1260
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Text            =   "xx kg/s"
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   2220
         MousePointer    =   1  'Arrow
         TabIndex        =   68
         Text            =   "xx "
         Top             =   8820
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   2220
         MousePointer    =   1  'Arrow
         TabIndex        =   67
         Text            =   "xx "
         Top             =   8520
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox F_875_desup_boiler 
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
         Height          =   285
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   66
         Text            =   "xx kg/s"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox P_875_desup_boiler 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Text            =   "xx Bar"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox F_SSTG_stm 
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
         Height          =   285
         Left            =   1200
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Text            =   "xx kg/s"
         Top             =   7560
         Width           =   855
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   461
         Left            =   4800
         TabIndex        =   63
         Top             =   6840
         Visible         =   0   'False
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
         ValveType       =   11
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   62
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox F_main_feed 
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
         Index           =   2
         Left            =   300
         TabIndex        =   61
         Text            =   "xx.x kg/s"
         Top             =   720
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox T_main_feed 
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
         Index           =   2
         Left            =   2640
         TabIndex        =   60
         Text            =   "xxx  F"
         Top             =   1440
         Width           =   675
      End
      Begin VB.TextBox T_econ_IN 
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
         Index           =   2
         Left            =   2880
         TabIndex        =   59
         Text            =   "xxx  F"
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox P_875_desup 
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
         Height          =   285
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   58
         Text            =   "xx Bar"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SSTG"
         Height          =   435
         Left            =   660
         TabIndex        =   45
         Top             =   8100
         Width           =   555
      End
      Begin VB.TextBox Psteam105 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         Text            =   "xx Bar"
         Top             =   5820
         Width           =   735
      End
      Begin VB.TextBox Tsteam105 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2820
         MousePointer    =   1  'Arrow
         TabIndex        =   32
         Text            =   "xx C"
         Top             =   4500
         Width           =   735
      End
      Begin VB.TextBox Fsteam101 
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
         Height          =   285
         Left            =   2820
         MousePointer    =   1  'Arrow
         TabIndex        =   31
         Text            =   "xx kg/s"
         Top             =   4200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   8580
      End
      Begin VB.TextBox Tsteam101 
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
         Height          =   285
         Left            =   3120
         MousePointer    =   1  'Arrow
         TabIndex        =   27
         Text            =   "xx C"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox Drum_water_level 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6060
         MousePointer    =   1  'Arrow
         TabIndex        =   26
         Text            =   "xx in"
         Top             =   2820
         Width           =   855
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   478
         Left            =   2400
         TabIndex        =   1
         Top             =   780
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
         Height          =   315
         Index           =   460
         Left            =   3660
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   479
         Left            =   1380
         TabIndex        =   3
         Top             =   1500
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   480
         Left            =   1380
         TabIndex        =   4
         Top             =   780
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   473
         Left            =   8160
         TabIndex        =   6
         Top             =   2260
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   469
         Left            =   3600
         TabIndex        =   7
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
         Height          =   375
         Index           =   476
         Left            =   6540
         TabIndex        =   8
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
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
         Height          =   375
         Index           =   472
         Left            =   6000
         TabIndex        =   9
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
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
         Height          =   435
         Index           =   466
         Left            =   3600
         TabIndex        =   10
         Top             =   840
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
         Index           =   467
         Left            =   6480
         TabIndex        =   11
         Top             =   7620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   463
         Left            =   6840
         TabIndex        =   12
         Top             =   7620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   468
         Left            =   1080
         TabIndex        =   14
         Top             =   6390
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   471
         Left            =   6900
         TabIndex        =   16
         Top             =   1480
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
         Height          =   555
         Index           =   0
         Left            =   6360
         TabIndex        =   23
         Top             =   2220
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   979
         TopColour       =   16711680
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   470
         Left            =   3360
         TabIndex        =   30
         Top             =   1320
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   464
         Left            =   1440
         TabIndex        =   34
         Top             =   8000
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   465
         Left            =   1380
         TabIndex        =   15
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   481
         Left            =   2820
         TabIndex        =   36
         Top             =   5640
         Width           =   435
         _ExtentX        =   767
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
         Height          =   375
         Index           =   475
         Left            =   2835
         TabIndex        =   38
         Top             =   7440
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   474
         Left            =   2080
         TabIndex        =   40
         Top             =   5640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   477
         Left            =   2100
         TabIndex        =   42
         Top             =   1480
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
         Index           =   462
         Left            =   10440
         TabIndex        =   13
         Top             =   8340
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   482
         Left            =   1740
         TabIndex        =   47
         Top             =   780
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
      End
      Begin VB.TextBox Psteam101 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Left            =   6060
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         Text            =   "xx Bar"
         Top             =   3120
         Width           =   855
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   9720
         Top             =   630
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   9720
         Top             =   360
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   9720
         Top             =   930
         Width           =   195
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Removed"
         Height          =   225
         Index           =   15
         Left            =   10050
         TabIndex        =   73
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "DO"
         Height          =   225
         Index           =   14
         Left            =   10050
         TabIndex        =   72
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "HO"
         Height          =   225
         Index           =   13
         Left            =   10050
         TabIndex        =   71
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner Tip"
         Height          =   225
         Index           =   12
         Left            =   9720
         TabIndex        =   70
         Top             =   90
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         Height          =   1155
         Left            =   9600
         Top             =   60
         Width           =   1455
      End
      Begin VB.Line Line8 
         X1              =   3300
         X2              =   3300
         Y1              =   4680
         Y2              =   5520
      End
      Begin VB.Line Line7 
         X1              =   2040
         X2              =   2160
         Y1              =   7740
         Y2              =   7740
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Index           =   1
         Left            =   4440
         OleObjectBlob   =   "Port_Boiler_V2.frx":0CFA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Vert_Left Green33_Mushrrom.vsd"
         TabIndex        =   49
         Top             =   5280
         Width           =   435
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Index           =   3
         Left            =   4440
         OleObjectBlob   =   "Port_Boiler_V2.frx":8F12
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Vert_Left_Green_Mushroom.vsd"
         TabIndex        =   51
         Top             =   5280
         Width           =   435
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Index           =   2
         Left            =   4440
         OleObjectBlob   =   "Port_Boiler_V2.frx":10F2A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Vert_Left_Green66_Mushroom.vsd"
         TabIndex        =   50
         Top             =   5280
         Width           =   435
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Index           =   0
         Left            =   4440
         OleObjectBlob   =   "Port_Boiler_V2.frx":19142
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Vert_Left_Black Mushroom.vsd"
         TabIndex        =   48
         Top             =   5280
         Width           =   435
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   960
         Picture         =   "Port_Boiler_V2.frx":2115A
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   255
      End
      Begin VB.Shape Burner2_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   4380
         Width           =   195
      End
      Begin VB.Shape Burner1_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   8460
         Shape           =   3  'Circle
         Top             =   6060
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   8280
         Top             =   4740
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   8700
         Top             =   6360
         Width           =   195
      End
      Begin VB.Line Line5 
         X1              =   3180
         X2              =   3180
         Y1              =   900
         Y2              =   1140
      End
      Begin VB.Line Line4 
         X1              =   2940
         X2              =   3480
         Y1              =   1740
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   1800
         Y1              =   6360
         Y2              =   6660
      End
      Begin VB.Label V_feed_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   54
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Feed Water Control"
         Height          =   315
         Index           =   11
         Left            =   1260
         TabIndex        =   52
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Feed"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Feed"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vent"
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   41
         Top             =   5220
         Width           =   555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   960
         X2              =   2160
         Y1              =   8340
         Y2              =   8340
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Drain"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   39
         Top             =   7920
         Width           =   615
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Safety"
         Height          =   315
         Index           =   1
         Left            =   2580
         TabIndex        =   37
         Top             =   5220
         Width           =   555
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   2940
         X2              =   2940
         Y1              =   5460
         Y2              =   8340
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   2160
         X2              =   2160
         Y1              =   5460
         Y2              =   8340
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Port Boiler SSTG Stop"
         Height          =   495
         Left            =   1140
         TabIndex        =   35
         Top             =   8580
         Width           =   1035
      End
      Begin VB.Label V_attemp_pos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xx %"
         Height          =   255
         Index           =   2
         Left            =   4020
         TabIndex        =   29
         Top             =   5100
         Width           =   555
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 2 Burner"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         Top             =   5400
         Width           =   750
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 1 Burner"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8400
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   7020
         Width           =   750
      End
      Begin VB.Shape Shape8 
         Height          =   1065
         Left            =   7770
         Shape           =   3  'Circle
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Shape Shape16 
         Height          =   1065
         Left            =   8190
         Shape           =   3  'Circle
         Top             =   5880
         Width           =   1155
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   8160
         Shape           =   2  'Oval
         Top             =   5010
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   8580
         Shape           =   2  'Oval
         Top             =   6630
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Height          =   1275
         Left            =   5835
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label Vent 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Attemperator Control"
         Height          =   495
         Index           =   8
         Left            =   3420
         TabIndex        =   22
         Top             =   4860
         Width           =   1095
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blow Down Valves"
         Height          =   495
         Index           =   7
         Left            =   6360
         TabIndex        =   21
         Top             =   8100
         Width           =   1155
      End
      Begin VB.Label Vent 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Surface Blow Valve"
         Height          =   375
         Index           =   6
         Left            =   8640
         TabIndex        =   20
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Air Cock"
         Height          =   315
         Index           =   5
         Left            =   7620
         TabIndex        =   19
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Steam Stop valve"
         Height          =   495
         Index           =   4
         Left            =   1140
         TabIndex        =   18
         Top             =   6840
         Width           =   915
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeSuperheated Stop valve"
         Height          =   495
         Index           =   3
         Left            =   1020
         TabIndex        =   17
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   540
         Width           =   255
      End
      Begin VB.OLE Port_Boiler_OLE 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   8475
         Left            =   840
         OleObjectBlob   =   "Port_Boiler_V2.frx":21464
         SizeMode        =   1  'Stretch
         SourceDoc       =   "\\Pc01\PC01_C\STEAM_SIM\Images\Port_boiler_WV1.vsd"
         TabIndex        =   0
         Top             =   660
         Width           =   10170
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5775
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   5895
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1935
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   540
         Width           =   1755
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   2100
         OleObjectBlob   =   "Port_Boiler_V2.frx":3067C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   46
         Top             =   720
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   2100
         OleObjectBlob   =   "Port_Boiler_V2.frx":38C94
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   57
         Top             =   720
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   2100
         OleObjectBlob   =   "Port_Boiler_V2.frx":40AAC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   56
         Top             =   720
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   2100
         OleObjectBlob   =   "Port_Boiler_V2.frx":48CC4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   55
         Top             =   720
         Width           =   315
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1620
      TabIndex        =   53
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Port_Boiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim set_dummy_focus As Integer

Private Sub Compressor_Updated(Code As Integer)

End Sub

Private Sub Command1_Click()

SSTG.Show

End Sub

Private Sub Drum_water_level_Click()

    Unload Alarm_set_point   ' Do this first
    
    nAlarm_HI = 52
    nAlarm_LO = 60  ' 6/1/00

    Alarm_set_point.Show

End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idsy155)


' === 13/11/02 ===
update_VBAI(5) = True


' 5/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        F_875_desup_boiler.Visible = False
        
        '21/9/00
        V_attemp_pos(2).Visible = False
        
Else
        F_875_desup_boiler.Visible = True
        
        '21/9/00
        V_attemp_pos(2).Visible = True

End If

old_greenness_LSS(5) = -123    ' To ensure an update
update_VBAI(35) = True
set_dummy_focus = False

' 1/12/99
HScroll2.Visible = PTT

' 7/1/00
Feed_reg(3).ZOrder
Feed_reg(2).ZOrder
Feed_reg(1).ZOrder
Feed_reg(0).ZOrder

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy155)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub HScroll2_Change()

Dim AIvalue As Integer

' 5/1/00   Stbd Aux Feed Valve(479)
' VALVE is now variable ( used to be clickable )

AIvalue = HScroll2.Value / 4
Call setAI(5, AIvalue)


End Sub

Private Sub P_875_desup_boiler_Click()

  ' 5/5/00 ADDED:
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 58

    Alarm_set_point.Show


End Sub

Private Sub P_875_desup_Click()


' 5/5/00 MOVED
'  ' 6/1/00
'    Unload Alarm_set_point   ' Do this first
'    nAlarm_LO = 58
'
'    Alarm_set_point.Show

End Sub

Private Sub Psteam101_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 53
    nAlarm_LO = 61  ' 6/1/99
    Alarm_set_point.Show

End Sub

Private Sub Psteam105_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 51
    nAlarm_LO = 59  ' 6/1/00

    Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim B_water_level As Integer
Dim VV_pos As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

'c        === 19/11/02 ===
'        PC_REALS(881) = F_SSTG_stm(1)
'        PC_REALS(882) = F_SSTG_stm(2)
'        PC_REALS(883) = QAVAILABLE(1) * 0.001
'        PC_REALS(884) = QAVAILABLE(2) * 0.001
'Text1.text = Format(LSS_reals(882), "##0.00")
'Text2.text = Format(LSS_reals(884), "####0")




' === 13/11/02 =======
If update_VBAI(5) Then
     update_VBAI(5) = False
     
     ' === 26/03/03 ===
     'HScroll2.Value = VB_AI(5) * 4
     Call update_the_slider(HScroll2, VB_AI(5) * 4)
     
End If



' 10/5/00 Do NOT need these...
'Call paint_alarm(Me.Psteam101, 53, 0, 0, 0)
'Call paint_alarm(Me.Psteam105, 51, 0, 0, 0)
'Call paint_alarm(Me.Drum_water_level, 52, 0, 0, 0)
'Call paint_alarm(Me.Tsteam105, 50, 0, 0, 0)

' 7/1/00
Call paint_alarm(Me.Psteam101, 53, 61, 0, 0)
Call paint_alarm(Me.Psteam105, 51, 59, 0, 0)
Call paint_alarm(Me.Drum_water_level, 52, 60, 0, 0)
Call paint_alarm(Me.Tsteam105, 50, 0, 0, 0)

' 5/5/00
'Call paint_alarm(Me.P_875_desup, 58, 0, 0, 0)
Call paint_alarm(Me.P_875_desup_boiler, 58, 0, 0, 0)


' 26/11/99   HO/DO Nozzles
'equivalence(STEAM_logicals(15),HO_nozzle(1))

For i = 2 To 3

    ' 12/10/09 REMOVED
    'If VB_steam_logicals(15 + i) Then
    '    Burner_tip(i).FillColor = &H4080&         ' DARK BROWN
    'Else
    '    Burner_tip(i).FillColor = &H80FF&         ' ORANGE
    'End If
    
    ' 12/10/09 ADDED
    If LSS_reals(877 + i) > 1.5 Then
        'HO Nozzle fited
         Burner_tip(i).FillColor = &H4080&         ' DARK BROWN
    ElseIf LSS_reals(877 + i) > 0.5 Then
        'DO Nozzle fited
        Burner_tip(i).FillColor = &H80FF&         ' ORANGE
    Else
         Burner_tip(i).FillColor = vbWhite        ' white
    End If
    
    
Next i


'        equivalence(STEAM_logicals(1),Burner1_flame(1))
'        equivalence(STEAM_logicals(3),Burner2_flame(1))
'        logical*1 Burner1_flame(2)
'        logical*1 Burner2_flame(2)
    If VB_steam_logicals(2) Then
        Burner1_shape(2).FillColor = RGB(&HFF, 0, 0)
    Else
        Burner1_shape(2).FillColor = RGB(&HC0, &HC0, &HC0)
    End If
    If VB_steam_logicals(4) Then
        Burner2_shape(2).FillColor = RGB(&HFF, 0, 0)
    Else
        Burner2_shape(2).FillColor = RGB(&HC0, &HC0, &HC0)
    End If

 
'          PC_REALS(340) = VV_pos_drum(1) * 100.0
'          PC_REALS(341) = VV_pos_drum(2) * 100.0



V_feed_pos(2).Caption = Format(LSS_reals(341), "###0") & " %"
If LSS_reals(341) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(341) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(341) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Feed_reg(i).Visible = True
    Else
        Feed_reg(i).Visible = False
    End If
Next i



V_attemp_pos(2).Caption = Format(LSS_reals(787), "###0") & " %"

If LSS_reals(787) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(787) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(787) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Attemp_Reg(i).Visible = True
    Else
        Attemp_Reg(i).Visible = False
    End If
Next i

'  OV015
'          PC_REALS(363) = P_875_desup_OUT


    ' 17/11/99
    F_main_feed(2).Visible = Advanced_mode
    
'        PC_REALS(386) = T_econ_IN(1)
'        PC_REALS(387) = T_econ_IN(2)
'        PC_REALS(388) = T_econ_OUT(1)
'        PC_REALS(389) = T_econ_OUT(2)
'        PC_REALS(390) = T_main_feed(1)
'        PC_REALS(391) = T_main_feed(2)

'cc 4 / 4 / 0
'    PC_REALS(127) = Psteam(5)
'    PC_REALS(128) = Psteam(105)


T_econ_IN(2).text = Format(LSS_reals(387) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_main_feed(2).text = Format(LSS_reals(391) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

'   3/9/99
Tsteam101.text = Format(LSS_reals(772) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam105.text = Format(LSS_reals(784) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

If SI_units Then

    ' 17/11/99
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult, "###0.0") & Flow_chars

    Psteam101.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    ' 4/5/00
    P_875_desup_boiler.text = Format(LSS_reals(493) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    ' 5/5/00
    F_875_desup_boiler.text = Format(LSS_reals(495) * Flow_mult, "###0.0") & Flow_chars
   
   ' 4/4/00
  ' Psteam105.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    Psteam105.text = Format(LSS_reals(128) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    Fsteam101.text = Format(LSS_reals(775) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
    
    ' === 20/11/02 ===
    '    PC_REALS(490) = F_MAIN_stm(1)
    '    PC_REALS(491) = F_MAIN_stm(2)
    F_MAIN_stm.text = Format(LSS_reals(491) * Flow_mult, "###0.0") & Flow_chars       ' kg/s

    
    'c        === 20/11/02 ===
    '        PC_REALS(881) = F_SSTG_stm(1)
    '        PC_REALS(882) = F_SSTG_stm(2)
    'F_SSTG_boiler.text = Format(LSS_reals(491) * Flow_mult, "###0.00") & Flow_chars       ' kg/s
     F_SSTG_stm.text = Format(LSS_reals(882) * Flow_mult, "###0.00") & Flow_chars       ' kg/s
   
    Drum_water_level.text = Format(LSS_reals(116) * 25.4, "###0") & " mm"
    
    P_875_desup.text = Format(LSS_reals(363) * Press_mult, "##0.0") & Press_chars     ' Bar
Else

     ' 17/11/99
   F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult / 3600#, "###0.0") & " lb/s"

    Psteam101.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    
    ' 4/5/00
    P_875_desup_boiler.text = Format(LSS_reals(493) * Press_mult, "###0") & Press_chars       ' psi
    ' 5/5/00
    F_875_desup_boiler.text = Format(LSS_reals(495) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
   ' 4/4/00
   'Psteam105.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    Psteam105.text = Format(LSS_reals(128) * Press_mult, "###0") & Press_chars       ' psi
    
    Fsteam101.text = Format(LSS_reals(775) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MAIN_stm.text = Format(LSS_reals(491) * Flow_mult / 3600#, "###0.0") & " lb/s"

    'c        === 20/11/02 ===
    '        PC_REALS(881) = F_SSTG_stm(1)
    '        PC_REALS(882) = F_SSTG_stm(2)
    'F_SSTG_boiler.text = Format(LSS_reals(491) * Flow_mult / 3600#, "###0.00") & " lb/s"
     F_SSTG_stm.text = Format(LSS_reals(882) * Flow_mult / 3600#, "###0.00") & " lb/s"
    
    
    Drum_water_level.text = Format(LSS_reals(116), "##0.0") & " in"
    
    P_875_desup.text = Format(LSS_reals(363) * Press_mult, "###0") & Press_chars       ' psi

End If

' Drum Water Level ( Tank )
If LSS_reals(116) > 8.9999 Then
    B_water_level = 100         ' FULL
ElseIf LSS_reals(116) < -8.9999 Then
    B_water_level = 0           ' EMPTY
Else
    B_water_level = Int((LSS_reals(116) + 9#) * 100# / 18#)
End If

' 30/11/99 VB ( or OCX bug )..
' when going from 20% to 0% the Blue water level does not change so..
If B_water_level < 1 Then
   B_water_level = 1
End If

    valve(0).LineWidth = B_water_level

' 28/10/99
'For i = 462 To 481
For i = 462 To 482

    ' 6/01/ 00
   'If i = 472 Or i = 476 Or i = 481 Then
    If i = 472 Or i = 476 Or i = 481 Or i = 479 Then
      ' Do NOT update safety valves here ( see code below )
    Else
           Call update_vvalve_LSS(Me, i, i, -1)
    End If
Next i


'c 6 / 1 / 0
'    PC_REALS(483) = VV_aux_feed_pos(1)
'    PC_REALS(484) = VV_aux_feed_pos(2)

' 6/1/00
Call update_green_VV(484, valve(479))


'       80,81     DRUM_Safety_VV_1(1-2)
'       89,90     DRUM_Safety_VV_2(1-2)
'       91,92     SUPER_Safety_VV(1-2)



'  DRUM Safety Valve # 1
If VB_steam_logicals(81) Then
    valve(472).OpenBtmRight = True
    valve(472).OpenBtmRight = True
Else
    valve(472).OpenBtmRight = False
    valve(472).OpenBtmRight = False  ' need to open other safety VV 476 & 477 as well
End If

'  DRUM Safety Valve # 2
If VB_steam_logicals(90) Then
    valve(476).OpenBtmRight = True
    valve(476).OpenBtmRight = True
Else
    valve(476).OpenBtmRight = False
    valve(476).OpenBtmRight = False  ' need to open other safety VV 476 & 477 as well
End If

'  SUPER-HEAT Safety Valve
If VB_steam_logicals(92) Then
    valve(481).OpenBtmRight = True
    valve(481).OpenBtmRight = True
Else
    valve(481).OpenBtmRight = False
    valve(481).OpenBtmRight = False  ' need to open other safety VV 476 & 477 as well
End If



End Sub

Private Sub Tsteam105_Click()
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 50
    Alarm_set_point.Show

End Sub

Private Sub valve_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
 ' 14/8/00 Block aux feed valve click on ACT Terminals
' === 23/09/02 Point 17 Block 448 and 450 on ACT Terminals as well ===
'If Index = 477 Then
If Index = 477 Or Index = 482 Or Index = 478 Then


  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
  
  
End If
  
   
   
   
' 28/10/99
'   If index = 472 Or index = 476 Or index = 477 Then
    If Index = 472 Or Index = 476 Or Index = 481 Then
        'skip safety valves (can't click these !)
    Else
        Call valve_pcdi_LSS(Index)
    End If

End Sub
Sub update_green_VV(nReals As Integer, obj As Object)
Dim green_colour As Integer

' 5/1/00
If LSS_reals(nReals) > 99# Then
    green_colour = 255
ElseIf LSS_reals(nReals) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(nReals) * 2.55
End If

obj.ColourClosed = RGB(0, green_colour, 0)

End Sub
Sub update_the_slider(Slider As Object, Pot_Value As Integer)

Dim S_max As Integer
Dim S_min As Integer

' === 26/03/03 (ADDED) ===

' TESTING...
S_max = Slider.max
S_min = Slider.Min

' === 19/02/03 ===
'VScroll1.Value = VB_AI(6) * 4
If Pot_Value >= Slider.max Then
    Slider.Value = Slider.max
ElseIf Pot_Value <= Slider.Min Then
    Slider.Value = Slider.Min
Else
    Slider.Value = Pot_Value
End If


End Sub

