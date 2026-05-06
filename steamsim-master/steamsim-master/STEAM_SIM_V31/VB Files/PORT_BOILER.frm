VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form BOILER_PORT 
   Caption         =   "PORT BOILER "
   ClientHeight    =   9540
   ClientLeft      =   1605
   ClientTop       =   585
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   12930
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8745
      Left            =   0
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   15425
      Begin VB.TextBox Fsteam101 
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
         Left            =   11100
         MousePointer    =   1  'Arrow
         TabIndex        =   48
         Text            =   "xx kg/s"
         Top             =   7920
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   120
         Top             =   8160
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
         Left            =   3780
         MousePointer    =   1  'Arrow
         TabIndex        =   45
         Text            =   "xx in"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Tsteam101 
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
         Left            =   7140
         MousePointer    =   1  'Arrow
         TabIndex        =   27
         Text            =   "xx C"
         Top             =   840
         Width           =   735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Left            =   10080
         TabIndex        =   25
         Top             =   6630
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   11100
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Text            =   "xx C"
         Top             =   7620
         Width           =   735
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
         Left            =   11100
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Text            =   "xx Bar"
         Top             =   7320
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   2
         Left            =   11160
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BottomColour    =   65280
         EnabledTopLeft  =   0   'False
         ColourClosed    =   255
         BackColour      =   14745599
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Left            =   5520
         Max             =   255
         TabIndex        =   13
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton Command3D4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -2460
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
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
         Height          =   285
         Left            =   6420
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Text            =   "xx Bar"
         Top             =   840
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   5
         Left            =   6300
         TabIndex        =   7
         Top             =   300
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
         Height          =   495
         Index           =   7
         Left            =   750
         TabIndex        =   8
         Top             =   7110
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
         Height          =   675
         Index           =   9
         Left            =   4920
         TabIndex        =   10
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1191
         TopColour       =   16711680
         BottomColour    =   16711680
         LineWidth       =   40
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
         Left            =   11280
         TabIndex        =   9
         Top             =   6900
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   1
         Left            =   8340
         TabIndex        =   23
         Top             =   5910
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
         Index           =   3
         Left            =   1170
         TabIndex        =   28
         Top             =   7110
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
         Left            =   1530
         TabIndex        =   29
         Top             =   1140
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
         Left            =   270
         TabIndex        =   30
         Top             =   1140
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
         Left            =   570
         TabIndex        =   31
         Top             =   2220
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
         Index           =   14
         Left            =   1050
         TabIndex        =   32
         Top             =   4890
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
         Left            =   1200
         TabIndex        =   33
         Top             =   6060
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
         Left            =   840
         TabIndex        =   34
         Top             =   1140
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
         Left            =   1140
         TabIndex        =   44
         Top             =   3480
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
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   4
         Left            =   10800
         TabIndex        =   49
         Top             =   3840
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
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Port Boiler SSTG Bulkhead VV"
         Height          =   735
         Left            =   10680
         TabIndex        =   50
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Label V_attemp_pos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "xx %"
         Height          =   255
         Index           =   2
         Left            =   8220
         TabIndex        =   47
         Top             =   6300
         Width           =   735
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1020
         X2              =   1860
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   915
         Left            =   300
         Shape           =   3  'Circle
         Top             =   3600
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
         Left            =   120
         TabIndex        =   43
         Top             =   3240
         Width           =   2085
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   8640
         Picture         =   "PORT_BOILER.frx":0000
         Stretch         =   -1  'True
         Top             =   7020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line44 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   9870
         X2              =   8670
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   11640
         Picture         =   "PORT_BOILER.frx":030A
         Stretch         =   -1  'True
         Top             =   7020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   7620
         Picture         =   "PORT_BOILER.frx":0614
         Stretch         =   -1  'True
         Top             =   5400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   5280
         Picture         =   "PORT_BOILER.frx":091E
         Stretch         =   -1  'True
         Top             =   540
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   7860
         Picture         =   "PORT_BOILER.frx":0C28
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   8460
         Picture         =   "PORT_BOILER.frx":0F32
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   8760
         Picture         =   "PORT_BOILER.frx":123C
         Stretch         =   -1  'True
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1500
         Picture         =   "PORT_BOILER.frx":1546
         Stretch         =   -1  'True
         Top             =   7320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   1740
         Picture         =   "PORT_BOILER.frx":1850
         Stretch         =   -1  'True
         Top             =   6300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   1560
         Picture         =   "PORT_BOILER.frx":1B5A
         Stretch         =   -1  'True
         Top             =   5100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   1080
         Picture         =   "PORT_BOILER.frx":1E64
         Stretch         =   -1  'True
         Top             =   2460
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line89 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   1650
         X2              =   1650
         Y1              =   540
         Y2              =   1740
      End
      Begin VB.Line Line88 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   390
         X2              =   390
         Y1              =   1800
         Y2              =   780
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
         Left            =   1530
         MousePointer    =   1  'Arrow
         TabIndex        =   42
         Top             =   2460
         Width           =   900
      End
      Begin VB.Line Line87 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   90
         X2              =   330
         Y1              =   2340
         Y2              =   2580
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
         Left            =   450
         MousePointer    =   1  'Arrow
         TabIndex        =   41
         Top             =   390
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
         Left            =   1710
         MousePointer    =   1  'Arrow
         TabIndex        =   40
         Top             =   600
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line86 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   330
         X2              =   1170
         Y1              =   2580
         Y2              =   2580
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
         Left            =   150
         MousePointer    =   1  'Arrow
         TabIndex        =   39
         Top             =   2730
         Width           =   1395
      End
      Begin VB.Line Line85 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   750
         X2              =   1590
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line Line84 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   510
         X2              =   750
         Y1              =   5610
         Y2              =   5250
      End
      Begin VB.Line Line75 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   870
         X2              =   1770
         Y1              =   6420
         Y2              =   6420
      End
      Begin VB.Line Line69 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   630
         X2              =   870
         Y1              =   6180
         Y2              =   6420
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
         Left            =   150
         TabIndex        =   38
         Top             =   5760
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
         Left            =   570
         TabIndex        =   37
         Top             =   4920
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
         Left            =   420
         TabIndex        =   36
         Top             =   6450
         Width           =   495
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1875
         Index           =   4
         Left            =   30
         Top             =   4860
         Width           =   2355
      End
      Begin VB.Line Line68 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   930
         X2              =   420
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   4755
         Index           =   3
         Left            =   60
         Top             =   0
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
         Left            =   90
         TabIndex        =   35
         Tag             =   "Steam Drum"
         Top             =   60
         Width           =   2175
      End
      Begin VB.Line Line65 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   930
         X2              =   930
         Y1              =   1800
         Y2              =   780
      End
      Begin VB.Line Line64 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   660
         X2              =   660
         Y1              =   1860
         Y2              =   2100
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   4740
         Shape           =   2  'Oval
         Top             =   6360
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   4320
         Shape           =   2  'Oval
         Top             =   4740
         Width           =   375
      End
      Begin VB.Shape Shape17 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4800
         Top             =   6000
         Width           =   255
      End
      Begin VB.Shape Shape16 
         Height          =   1065
         Left            =   4350
         Shape           =   3  'Circle
         Top             =   5610
         Width           =   1155
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4470
         Shape           =   3  'Circle
         Top             =   5790
         Width           =   375
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4380
         Top             =   4380
         Width           =   255
      End
      Begin VB.Shape Shape8 
         Height          =   1065
         Left            =   3930
         Shape           =   3  'Circle
         Top             =   3990
         Width           =   1155
      End
      Begin VB.Line Line81 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   6900
         X2              =   7920
         Y1              =   5250
         Y2              =   4830
      End
      Begin VB.Line Line80 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7380
         X2              =   7920
         Y1              =   5310
         Y2              =   5070
      End
      Begin VB.Line Line79 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   6900
         X2              =   7920
         Y1              =   5730
         Y2              =   5310
      End
      Begin VB.Line Line78 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7380
         X2              =   7920
         Y1              =   5790
         Y2              =   5550
      End
      Begin VB.Line Line77 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7020
         X2              =   8040
         Y1              =   6240
         Y2              =   5820
      End
      Begin VB.Line Line76 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7500
         X2              =   8040
         Y1              =   6300
         Y2              =   6060
      End
      Begin VB.Line Line73 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7080
         X2              =   8100
         Y1              =   6750
         Y2              =   6330
      End
      Begin VB.Line Line72 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7560
         X2              =   8100
         Y1              =   6810
         Y2              =   6570
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   6990
         X2              =   6870
         Y1              =   5910
         Y2              =   5940
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   7200
         X2              =   7320
         Y1              =   5160
         Y2              =   5130
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soot Blowers"
         Height          =   255
         Left            =   9480
         TabIndex        =   26
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9060
         X2              =   9060
         Y1              =   2580
         Y2              =   7140
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
         Left            =   4080
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   6090
         Width           =   645
      End
      Begin VB.Line Line71 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7290
         X2              =   7830
         Y1              =   4830
         Y2              =   4590
      End
      Begin VB.Line Line70 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7200
         X2              =   7740
         Y1              =   4380
         Y2              =   4140
      End
      Begin VB.Line Line62 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   6840
         X2              =   7860
         Y1              =   4770
         Y2              =   4350
      End
      Begin VB.Line Line47 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   6780
         X2              =   7800
         Y1              =   4320
         Y2              =   3900
      End
      Begin VB.Shape Shape7 
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   9780
         Top             =   7080
         Width           =   195
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   8220
         X2              =   7860
         Y1              =   7020
         Y2              =   3060
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   8460
         X2              =   8460
         Y1              =   7140
         Y2              =   2700
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   7500
         Shape           =   3  'Circle
         Top             =   3540
         Width           =   195
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7500
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   375
      End
      Begin VB.Line Line67 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   6660
         X2              =   7680
         Y1              =   3900
         Y2              =   3480
      End
      Begin VB.Line Line66 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   7080
         X2              =   7620
         Y1              =   3960
         Y2              =   3720
      End
      Begin VB.Line Line63 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   6720
         X2              =   6810
         Y1              =   4620
         Y2              =   4590
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   7500
         X2              =   8280
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6660
         X2              =   9060
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   7980
         X2              =   7980
         Y1              =   660
         Y2              =   2340
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6660
         X2              =   7980
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6660
         X2              =   6660
         Y1              =   2340
         Y2              =   2580
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   7800
         X2              =   7740
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label Label12 
         Caption         =   "Economiser"
         Height          =   255
         Left            =   8400
         TabIndex        =   20
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label M 
         Caption         =   "Main Feed"
         Height          =   195
         Left            =   9720
         TabIndex        =   19
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Superheater Headers"
         Height          =   435
         Left            =   5820
         TabIndex        =   18
         Top             =   3720
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "WET STEAM"
         Height          =   255
         Left            =   4140
         TabIndex        =   17
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPERHEATED STEAM"
         Height          =   195
         Left            =   10170
         TabIndex        =   16
         Top             =   6330
         Width           =   2115
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6960
         X2              =   8220
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   7860
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   8220
         X2              =   8220
         Y1              =   2940
         Y2              =   5520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6360
         X2              =   7140
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6360
         X2              =   6360
         Y1              =   1200
         Y2              =   3420
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
         Left            =   8100
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   2160
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line61 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   7440
         X2              =   9660
         Y1              =   7980
         Y2              =   7980
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   8580
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   9180
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Line Line60 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   7860
         X2              =   7680
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line59 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9480
         X2              =   9480
         Y1              =   1800
         Y2              =   1980
      End
      Begin VB.Line Line58 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8100
         X2              =   9480
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line57 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8340
         X2              =   9480
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8340
         X2              =   8340
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9480
         X2              =   9480
         Y1              =   1260
         Y2              =   1440
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8340
         X2              =   8340
         Y1              =   1080
         Y2              =   1260
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8340
         X2              =   9480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8340
         X2              =   9480
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8340
         X2              =   9780
         Y1              =   1080
         Y2              =   1080
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
         Left            =   4500
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   1800
         Width           =   300
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   3180
         X2              =   3180
         Y1              =   3360
         Y2              =   7860
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   7140
         X2              =   3240
         Y1              =   7980
         Y2              =   7980
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
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   4500
         Width           =   645
      End
      Begin VB.Line Line48 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   6480
         X2              =   5940
         Y1              =   7500
         Y2              =   3120
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   5940
         X2              =   6480
         Y1              =   3120
         Y2              =   2700
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   9870
         X2              =   9870
         Y1              =   7140
         Y2              =   6060
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
         Left            =   4500
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   2340
         Width           =   285
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   8280
         X2              =   8280
         Y1              =   1860
         Y2              =   660
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5160
         X2              =   6540
         Y1              =   1860
         Y2              =   2040
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   5160
         X2              =   6480
         Y1              =   2400
         Y2              =   2280
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   540
         X2              =   1380
         Y1              =   7470
         Y2              =   7470
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5220
         X2              =   7980
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   3180
         X2              =   6420
         Y1              =   3360
         Y2              =   2400
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   7500
         X2              =   8220
         Y1              =   7680
         Y2              =   7020
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   7680
         X2              =   7860
         Y1              =   2880
         Y2              =   3060
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   8460
         X2              =   6960
         Y1              =   2700
         Y2              =   2700
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
         Left            =   6780
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   8160
         Width           =   1185
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   7140
         X2              =   7140
         Y1              =   1200
         Y2              =   1740
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   7080
         X2              =   11880
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Shape burner2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   4170
         Width           =   375
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   330
         X2              =   570
         Y1              =   7230
         Y2              =   7470
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
         Left            =   330
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   7770
         Width           =   945
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
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   6870
         Width           =   1980
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
         Left            =   1770
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   7350
         Width           =   675
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1215
         Index           =   0
         Left            =   30
         Top             =   6810
         Width           =   2355
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   8220
         X2              =   7380
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6960
         X2              =   6960
         Y1              =   2700
         Y2              =   2940
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   6480
         X2              =   7020
         Y1              =   7500
         Y2              =   7740
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6360
         X2              =   6600
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   9690
         X2              =   9690
         Y1              =   7980
         Y2              =   420
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   6960
         Shape           =   3  'Circle
         Top             =   7560
         Width           =   675
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00808080&
         BorderWidth     =   8
         X1              =   7020
         X2              =   7500
         Y1              =   3390
         Y2              =   6930
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         BorderWidth     =   8
         X1              =   6600
         X2              =   7080
         Y1              =   3420
         Y2              =   7140
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   1395
         Left            =   6360
         Shape           =   3  'Circle
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00C0FFFF&
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   7155
         Left            =   2580
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   7995
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2820
      TabIndex        =   46
      Top             =   0
      Width           =   435
   End
   Begin VB.Line Line74 
      BorderColor     =   &H00000000&
      BorderStyle     =   2  'Dash
      X1              =   480
      X2              =   1020
      Y1              =   480
      Y2              =   240
   End
End
Attribute VB_Name = "BOILER_PORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim set_dummy_focus As Integer


Private Sub Command3D3_Click()


End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idsy155)

' 22/7/98
old_greenness_LSS(5) = -123    ' To ensure an update

' 22/7/98
   update_VBAI(35) = True
   
   set_dummy_focus = False

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy155)
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
If VB_steam_logicals(81) Then
    Valve(12).OpenBtmRight = True
    Valve(12).OpenBtmRight = True
Else
    Valve(12).OpenBtmRight = False
    Valve(12).OpenBtmRight = False
End If

V_attemp_pos(2).Caption = Format(LSS_reals(787), "###0") & " %"

'   3/9/99
Tsteam101.text = Format(LSS_reals(772) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam105.text = Format(LSS_reals(784) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

If SI_units Then
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    Psteam105.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    Fsteam101.text = Format(LSS_reals(775) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
    Drum_water_level.text = Format(LSS_reals(116) * 25.4, "###0") & " mm"
Else
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    Psteam105.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    Fsteam101.text = Format(LSS_reals(775) * Flow_mult / 3600#, "###0.0") & " lb/s"
    Drum_water_level.text = Format(LSS_reals(116), "##0.0") & " in"
End If

' Drum Water Level ( Tank )
If LSS_reals(116) > 8.9999 Then
    B_water_level = 100         ' FULL
ElseIf LSS_reals(116) < -8.9999 Then
    B_water_level = 0           ' EMPTY
Else
    B_water_level = Int((LSS_reals(116) + 9#) * 100# / 18#)
End If
    Valve(9).LineWidth = B_water_level

' update_vvalve_LSS(syxxx,nValve,index1,index2)
' OLD LSS
'Call update_vvalve_LSS(Me, 18, 8, -1)    ' VV18

' 5/10/99
For i = 2 To 5
  Call update_vvalve_LSS(Me, i + 310, i, -1)
Next i
For i = 7 To 8
  Call update_vvalve_LSS(Me, i + 310, i, -1)
Next i

' SKIP valve 11 ( DO's NOT defined )
For i = 12 To 16
  Call update_vvalve_LSS(Me, i + 310, i, -1)
Next i


End Sub

Private Sub Valve_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' OLD LSS
'Select Case Index
'    Case 8
'        Call valve_pcdi_LSS(18)
'End Select

' 5/10/99
        Call valve_pcdi_LSS(Index + 310)

End Sub
