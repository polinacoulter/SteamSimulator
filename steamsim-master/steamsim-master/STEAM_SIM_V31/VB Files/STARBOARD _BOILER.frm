VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Boiler_Stbd 
   Caption         =   "STARBOARD BOILER"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   13305
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8805
      Left            =   0
      Top             =   -120
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   15531
      Begin VB.TextBox Tsteam5 
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
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   45
         Text            =   "xx C"
         Top             =   7560
         Width           =   855
      End
      Begin VB.TextBox Psteam5 
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
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   44
         Text            =   "xx Bar"
         Top             =   7260
         Width           =   855
      End
      Begin VB.TextBox Fsteam1 
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
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   43
         Text            =   "xx kg/s"
         Top             =   7860
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   420
         Top             =   5160
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
         Left            =   6480
         MousePointer    =   1  'Arrow
         TabIndex        =   40
         Text            =   "-xx in"
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox Psteam1 
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
         Left            =   3780
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Text            =   "xx Bar"
         Top             =   660
         Width           =   795
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Left            =   4320
         Max             =   255
         TabIndex        =   2
         Top             =   0
         Width           =   1395
      End
      Begin VB.TextBox Tsteam1 
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
         Left            =   4620
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Text            =   "xx C"
         Top             =   660
         Width           =   795
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   2
         Left            =   8040
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BottomColour    =   65280
         EnabledTopLeft  =   0   'False
         ColourClosed    =   255
         BackColour      =   14745599
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   5
         Left            =   4530
         TabIndex        =   4
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   675
         Index           =   9
         Left            =   6090
         TabIndex        =   5
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1191
         TopColour       =   16711680
         BottomColour    =   16711680
         LineWidth       =   5
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   6
         Top             =   6840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   1
         Left            =   2460
         TabIndex        =   7
         Top             =   5820
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   661
         TopColour       =   65280
         BottomColour    =   65280
         ColourClosed    =   255
         BackColour      =   14745599
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   7
         Left            =   10020
         TabIndex        =   19
         Top             =   7200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   3
         Left            =   10380
         TabIndex        =   20
         Top             =   7200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   11
         Left            =   10800
         TabIndex        =   24
         Top             =   1230
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   767
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   12
         Left            =   9540
         TabIndex        =   25
         Top             =   1200
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   13
         Left            =   9840
         TabIndex        =   26
         Top             =   2310
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   16777215
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   14
         Left            =   10260
         TabIndex        =   27
         Top             =   4980
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   15
         Left            =   10410
         TabIndex        =   28
         Top             =   6150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   16
         Left            =   10080
         TabIndex        =   29
         Top             =   1230
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   0
         Left            =   9960
         TabIndex        =   39
         Top             =   3720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   46
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   6
         Left            =   9480
         TabIndex        =   47
         Top             =   8160
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   10
         Left            =   9960
         TabIndex        =   48
         Top             =   8160
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label Label4 
         Caption         =   "Spare Valves  >>>>>>>"
         Height          =   195
         Left            =   7800
         TabIndex        =   50
         Top             =   8400
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Boiler SSTG Bulkhead VV"
         Height          =   735
         Left            =   120
         TabIndex        =   49
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label V_attemp_pos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "xx %"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   42
         Top             =   6240
         Width           =   855
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9600
         X2              =   10440
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   915
         Left            =   10440
         Shape           =   3  'Circle
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "FEED WATER REGULATOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   38
         Top             =   3480
         Width           =   2085
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   3060
         X2              =   2880
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   3060
         X2              =   3060
         Y1              =   6000
         Y2              =   5640
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   2220
         Picture         =   "STARBOARD _BOILER.frx":0000
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   3480
         Picture         =   "STARBOARD _BOILER.frx":030A
         Stretch         =   -1  'True
         Top             =   5340
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   3060
         Picture         =   "STARBOARD _BOILER.frx":0614
         Stretch         =   -1  'True
         Top             =   2720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   2520
         Picture         =   "STARBOARD _BOILER.frx":091E
         Stretch         =   -1  'True
         Top             =   2340
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   2040
         Picture         =   "STARBOARD _BOILER.frx":0C28
         Stretch         =   -1  'True
         Top             =   900
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   60
         Picture         =   "STARBOARD _BOILER.frx":0F32
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   10740
         Picture         =   "STARBOARD _BOILER.frx":123C
         Stretch         =   -1  'True
         Top             =   7440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   10920
         Picture         =   "STARBOARD _BOILER.frx":1546
         Stretch         =   -1  'True
         Top             =   6360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   10740
         Picture         =   "STARBOARD _BOILER.frx":1850
         Stretch         =   -1  'True
         Top             =   5220
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   10380
         Picture         =   "STARBOARD _BOILER.frx":1B5A
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   5700
         Picture         =   "STARBOARD _BOILER.frx":1E64
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line89 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   10920
         X2              =   10920
         Y1              =   630
         Y2              =   1830
      End
      Begin VB.Line Line88 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9660
         X2              =   9660
         Y1              =   1890
         Y2              =   870
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O/BOARD"
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
         Left            =   10800
         MousePointer    =   1  'Arrow
         TabIndex        =   37
         Top             =   2550
         Width           =   900
      End
      Begin VB.Line Line87 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9360
         X2              =   9600
         Y1              =   2430
         Y2              =   2670
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SAFETY VALVES"
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
         Height          =   330
         Left            =   9720
         MousePointer    =   1  'Arrow
         TabIndex        =   36
         Top             =   480
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label28 
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
         Height          =   330
         Left            =   10980
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Top             =   690
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line86 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9600
         X2              =   10440
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SCUMMING VALVE"
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
         Left            =   9420
         MousePointer    =   1  'Arrow
         TabIndex        =   34
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Line Line85 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   10020
         X2              =   10860
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Line84 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9780
         X2              =   10020
         Y1              =   5700
         Y2              =   5340
      End
      Begin VB.Line Line75 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   10140
         X2              =   11040
         Y1              =   6510
         Y2              =   6510
      End
      Begin VB.Line Line69 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9900
         X2              =   10140
         Y1              =   6270
         Y2              =   6510
      End
      Begin VB.Label Label26 
         Caption         =   "SUPERHEATER CONTROLS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9420
         TabIndex        =   33
         Top             =   5850
         Width           =   2085
      End
      Begin VB.Label Label25 
         Caption         =   "VENT"
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
         Left            =   9840
         TabIndex        =   32
         Top             =   5010
         Width           =   435
      End
      Begin VB.Label Label20 
         Caption         =   "DRAIN"
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
         Left            =   9690
         TabIndex        =   31
         Top             =   6540
         Width           =   495
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1875
         Index           =   4
         Left            =   9300
         Top             =   4950
         Width           =   2355
      End
      Begin VB.Line Line68 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   10200
         X2              =   9690
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   4815
         Index           =   3
         Left            =   9300
         Top             =   60
         Width           =   2355
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "STEAM DRUM MOUNTINGS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9360
         TabIndex        =   30
         Tag             =   "Steam Drum"
         Top             =   150
         Width           =   2175
      End
      Begin VB.Line Line65 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   10200
         X2              =   10200
         Y1              =   1890
         Y2              =   870
      End
      Begin VB.Line Line64 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9930
         X2              =   9930
         Y1              =   1950
         Y2              =   2190
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1215
         Index           =   0
         Left            =   9300
         Top             =   6900
         Width           =   2355
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O/BOARD"
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
         Left            =   10980
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Top             =   7440
         Width           =   675
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WATER DRUM MOUNTING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9510
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Top             =   6960
         Width           =   1980
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BLOWDOWN"
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
         Left            =   9600
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Top             =   7860
         Width           =   945
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9540
         X2              =   9780
         Y1              =   7320
         Y2              =   7560
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9780
         X2              =   10680
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Superheater Headers"
         Height          =   435
         Left            =   4680
         TabIndex        =   12
         Top             =   3630
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   8070
         Shape           =   3  'Circle
         Top             =   7680
         Width           =   195
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   3450
         Shape           =   3  'Circle
         Top             =   3420
         Width           =   315
      End
      Begin VB.Line Line63 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4590
         X2              =   4500
         Y1              =   4530
         Y2              =   4500
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4320
         X2              =   4440
         Y1              =   5910
         Y2              =   5940
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4050
         X2              =   3930
         Y1              =   5130
         Y2              =   5070
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4560
         X2              =   4560
         Y1              =   2220
         Y2              =   2460
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4320
         X2              =   4320
         Y1              =   2580
         Y2              =   2820
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3000
         X2              =   4320
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4290
         X2              =   2580
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2160
         X2              =   4530
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1650
         X2              =   1650
         Y1              =   1230
         Y2              =   1410
      End
      Begin VB.Line Line57 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1680
         X2              =   2820
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Shape burner2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6780
         Shape           =   3  'Circle
         Top             =   4050
         Width           =   375
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 2 Burner"
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
         Left            =   7140
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   4320
         Width           =   645
      End
      Begin VB.Shape Shape8 
         Height          =   1065
         Left            =   6150
         Shape           =   3  'Circle
         Top             =   3870
         Width           =   1155
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6600
         Top             =   4260
         Width           =   255
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   6540
         Shape           =   2  'Oval
         Top             =   4620
         Width           =   375
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         BorderWidth     =   8
         X1              =   4200
         X2              =   3810
         Y1              =   3300
         Y2              =   6900
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00808080&
         BorderWidth     =   8
         X1              =   4650
         X2              =   4260
         Y1              =   3330
         Y2              =   7140
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   3780
         Shape           =   3  'Circle
         Top             =   7530
         Width           =   675
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1530
         X2              =   1530
         Y1              =   7830
         Y2              =   270
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4650
         X2              =   4890
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   3120
         X2              =   3930
         Y1              =   7050
         Y2              =   7650
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3960
         X2              =   3000
         Y1              =   5430
         Y2              =   5460
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   -540
         X2              =   4260
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4110
         X2              =   4110
         Y1              =   1080
         Y2              =   1620
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WATER DRUM"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3630
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Top             =   8130
         Width           =   1185
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   4710
         X2              =   5310
         Y1              =   2670
         Y2              =   3030
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   4380
         X2              =   4920
         Y1              =   7680
         Y2              =   7410
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   4830
         X2              =   8130
         Y1              =   2280
         Y2              =   3210
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3210
         X2              =   5970
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   4770
         X2              =   6090
         Y1              =   2160
         Y2              =   2310
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4710
         X2              =   6120
         Y1              =   1920
         Y2              =   1680
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   2910
         X2              =   2910
         Y1              =   1800
         Y2              =   600
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-9 in."
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
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Top             =   2220
         Width           =   285
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   1380
         X2              =   1380
         Y1              =   7050
         Y2              =   5640
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   3360
         X2              =   3900
         Y1              =   2940
         Y2              =   2520
      End
      Begin VB.Line Line48 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   3120
         X2              =   3360
         Y1              =   7020
         Y2              =   2940
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   8070
         X2              =   4170
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   8160
         X2              =   8160
         Y1              =   3210
         Y2              =   7710
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+9 in."
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
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   1620
         Width           =   300
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1350
         X2              =   2790
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1650
         X2              =   2760
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1650
         X2              =   2820
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2790
         X2              =   2790
         Y1              =   1020
         Y2              =   1200
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2820
         X2              =   2820
         Y1              =   1410
         Y2              =   1770
      End
      Begin VB.Line Line58 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1680
         X2              =   3060
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line59 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1650
         X2              =   1650
         Y1              =   1770
         Y2              =   1950
      End
      Begin VB.Line Line60 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3480
         X2              =   3300
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2490
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   1890
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Line Line61 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1530
         X2              =   3810
         Y1              =   7830
         Y2              =   7830
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STEAM DRUM"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3300
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   1260
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4920
         X2              =   4920
         Y1              =   1050
         Y2              =   3330
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4110
         X2              =   4890
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3000
         X2              =   3000
         Y1              =   2850
         Y2              =   5430
      End
      Begin VB.Label Label10 
         Caption         =   "WET STEAM"
         Height          =   255
         Left            =   6300
         TabIndex        =   13
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label M 
         Caption         =   "Main Feed"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "Economiser"
         Height          =   255
         Left            =   1710
         TabIndex        =   10
         Top             =   660
         Width           =   1155
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3240
         X2              =   4560
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3210
         X2              =   3210
         Y1              =   510
         Y2              =   2190
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   2910
         X2              =   3690
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line44 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   3030
         X2              =   1380
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line66 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4110
         X2              =   3390
         Y1              =   3870
         Y2              =   3540
      End
      Begin VB.Line Line67 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4590
         X2              =   3630
         Y1              =   3810
         Y2              =   3390
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3360
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   375
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   195
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2580
         X2              =   2580
         Y1              =   7080
         Y2              =   2580
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   4950
         X2              =   5310
         Y1              =   7410
         Y2              =   3060
      End
      Begin VB.Shape Shape7 
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   1290
         Top             =   6990
         Width           =   195
      End
      Begin VB.Line Line47 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4500
         X2              =   3510
         Y1              =   4320
         Y2              =   3870
      End
      Begin VB.Line Line62 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4500
         X2              =   3420
         Y1              =   4740
         Y2              =   4260
      End
      Begin VB.Line Line70 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4080
         X2              =   3570
         Y1              =   4380
         Y2              =   4170
      End
      Begin VB.Line Line71 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4020
         X2              =   3420
         Y1              =   4800
         Y2              =   4560
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 1 Burner"
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
         Left            =   6810
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   6030
         Width           =   645
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2160
         X2              =   2160
         Y1              =   2460
         Y2              =   7050
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soot Blowers"
         Height          =   255
         Left            =   540
         TabIndex        =   8
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Line Line72 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   3840
         X2              =   3240
         Y1              =   6870
         Y2              =   6600
      End
      Begin VB.Line Line73 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4320
         X2              =   3330
         Y1              =   6750
         Y2              =   6300
      End
      Begin VB.Line Line76 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   3870
         X2              =   3330
         Y1              =   6330
         Y2              =   6060
      End
      Begin VB.Line Line77 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4320
         X2              =   3270
         Y1              =   6240
         Y2              =   5790
      End
      Begin VB.Line Line78 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   3900
         X2              =   3330
         Y1              =   5790
         Y2              =   5490
      End
      Begin VB.Line Line79 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4350
         X2              =   3330
         Y1              =   5640
         Y2              =   5250
      End
      Begin VB.Line Line80 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   3990
         X2              =   3420
         Y1              =   5310
         Y2              =   5040
      End
      Begin VB.Line Line81 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4440
         X2              =   3390
         Y1              =   5190
         Y2              =   4740
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6450
         Shape           =   3  'Circle
         Top             =   5700
         Width           =   375
      End
      Begin VB.Shape Shape16 
         Height          =   1065
         Left            =   5820
         Shape           =   3  'Circle
         Top             =   5520
         Width           =   1155
      End
      Begin VB.Shape Shape17 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6240
         Top             =   5910
         Width           =   255
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   6180
         Shape           =   2  'Oval
         Top             =   6270
         Width           =   375
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   1395
         Left            =   3330
         Shape           =   3  'Circle
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00C0FFFF&
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   7155
         Left            =   1020
         Shape           =   4  'Rounded Rectangle
         Top             =   1140
         Width           =   7995
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   41
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Boiler_Stbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idsy156)

set_dummy_focus = True

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy156)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim B_water_level As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 6/9/99 Safety Valve...
If VB_steam_logicals(80) Then
    Valve(12).OpenBtmRight = True
    Valve(12).OpenBtmRight = True
Else
    Valve(12).OpenBtmRight = False
    Valve(12).OpenBtmRight = False
End If
    

V_attemp_pos(1).Caption = Format(LSS_reals(786), "###0") & " %"

'   3/9/99
Tsteam1.text = Format(LSS_reals(755) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam5.text = Format(LSS_reals(767) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

If SI_units Then
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    Psteam5.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    Fsteam1.text = Format(LSS_reals(758) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
    Drum_water_level.text = Format(LSS_reals(115) * 25.4, "###0") & " mm"
Else
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    Psteam5.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    Fsteam1.text = Format(LSS_reals(758) * Flow_mult / 3600#, "###0.0") & " lb/s"
    Drum_water_level.text = Format(LSS_reals(115), "##0.0") & " in"
End If

' Drum Water Level ( Tank )
If LSS_reals(115) > 8.9999 Then
    B_water_level = 100         ' FULL
ElseIf LSS_reals(115) < -8.9999 Then
    B_water_level = 0           ' EMPTY
Else
    B_water_level = Int((LSS_reals(115) + 9#) * 100# / 18#)
End If
    Valve(9).LineWidth = B_water_level

' update_vvalve_LSS(syxxx,nValve,index1,index2)
' OLD LSS
' Call update_vvalve_LSS(Me, 16, 8, -1)    ' VV16
For i = 2 To 5
  Call update_vvalve_LSS(Me, i + 430, i, -1)  ' VV16
Next i
For i = 7 To 9
  Call update_vvalve_LSS(Me, i + 430, i, -1)  ' VV16
Next i
For i = 11 To 12
  Call update_vvalve_LSS(Me, i + 430, i, -1)  ' VV16
Next i

'  Valves 13 & 14 have NO addresses in .xls
For i = 15 To 16
  Call update_vvalve_LSS(Me, i + 430, i, -1)  ' VV16
Next i


End Sub

Private Sub valve_Click(Index As Integer)

' OLD LSS
'Select Case Index
'    Case 8
'        Call valve_pcdi_LSS(16)
'End Select

' 4/10/99
        Call valve_pcdi_LSS(Index + 430)


End Sub
