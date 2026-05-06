VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.0#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form OLD_Feed_Water 
   Caption         =   "Feed Water"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   13170
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8955
      Left            =   0
      Top             =   0
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   15796
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   6
         Left            =   4920
         TabIndex        =   6
         Top             =   1680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   65280
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   22
         Left            =   4920
         TabIndex        =   42
         Top             =   2440
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   65280
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.TextBox Text6 
         Height          =   435
         Left            =   1020
         TabIndex        =   107
         Text            =   "Text6"
         Top             =   5400
         Width           =   1815
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1275
         Left            =   4320
         Max             =   100
         TabIndex        =   106
         Top             =   1320
         Width           =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   720
         Top             =   2220
      End
      Begin VB.TextBox Text5 
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
         Height          =   255
         Left            =   5100
         TabIndex        =   103
         Text            =   "P"
         Top             =   2820
         Width           =   195
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   135
         Left            =   6180
         Max             =   100
         TabIndex        =   102
         Top             =   6600
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   135
         Left            =   5700
         Max             =   100
         TabIndex        =   101
         Top             =   7380
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   10860
         Max             =   100
         TabIndex        =   98
         Top             =   1140
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   3240
         Max             =   100
         TabIndex        =   97
         Top             =   3680
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1155
         Left            =   6000
         Max             =   100
         TabIndex        =   96
         Top             =   1740
         Width           =   135
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   8220
         Max             =   100
         TabIndex        =   95
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text4 
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
         Height          =   255
         Left            =   3660
         TabIndex        =   94
         Text            =   "P"
         Top             =   4020
         Width           =   195
      End
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
         Height          =   240
         Left            =   10320
         TabIndex        =   92
         Text            =   "P"
         Top             =   280
         Width           =   195
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
         Height          =   240
         Left            =   9720
         TabIndex        =   89
         Text            =   "P"
         Top             =   280
         Width           =   195
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
         Height          =   255
         Left            =   4680
         TabIndex        =   88
         Text            =   "P"
         Top             =   2940
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "To Boiler"
         Height          =   495
         Left            =   1500
         TabIndex        =   37
         Top             =   6060
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Condensate System"
         Height          =   675
         Left            =   11760
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/B"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8400
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9300
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   8400
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8820
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   8400
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox P_feed_water 
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
         Left            =   7860
         TabIndex        =   13
         Text            =   "75 Bar"
         Top             =   7920
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   0
         Left            =   8820
         TabIndex        =   0
         Top             =   7560
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
         Index           =   1
         Left            =   9540
         TabIndex        =   1
         Top             =   6300
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
         Index           =   2
         Left            =   9780
         TabIndex        =   3
         Top             =   7800
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
         Left            =   6180
         TabIndex        =   4
         Top             =   6660
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
         Index           =   13
         Left            =   6720
         TabIndex        =   7
         Top             =   5580
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   14
         Left            =   6720
         TabIndex        =   8
         Top             =   4680
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   15
         Left            =   8760
         TabIndex        =   9
         Top             =   6060
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   9240
         TabIndex        =   18
         Top             =   7845
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   8940
         TabIndex        =   19
         Top             =   6360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   4
         Left            =   7560
         TabIndex        =   21
         Top             =   60
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   8
         Left            =   3660
         TabIndex        =   22
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
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   9
         Left            =   5640
         TabIndex        =   23
         Top             =   240
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   10
         Left            =   11040
         TabIndex        =   24
         Top             =   3780
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   3
         Left            =   8040
         TabIndex        =   30
         Top             =   6960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   17
         Left            =   8940
         TabIndex        =   31
         Top             =   6900
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
         Left            =   6180
         TabIndex        =   2
         Top             =   6060
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
         Height          =   255
         Index           =   16
         Left            =   7860
         TabIndex        =   33
         Top             =   6600
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   18
         Left            =   4020
         TabIndex        =   39
         Top             =   2880
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   19
         Left            =   3300
         TabIndex        =   40
         Top             =   2880
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   20
         Left            =   3660
         TabIndex        =   41
         Top             =   3240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   49152
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   23
         Left            =   5580
         TabIndex        =   43
         Top             =   2100
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   24
         Left            =   4920
         TabIndex        =   44
         Top             =   840
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   25
         Left            =   9240
         TabIndex        =   45
         Top             =   3060
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   26
         Left            =   8340
         TabIndex        =   46
         Top             =   3060
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   27
         Left            =   7080
         TabIndex        =   47
         Top             =   4680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   29
         Left            =   7080
         TabIndex        =   48
         Top             =   5580
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   30
         Left            =   6000
         TabIndex        =   49
         Top             =   5580
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   31
         Left            =   6000
         TabIndex        =   50
         Top             =   4680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   32
         Left            =   6360
         TabIndex        =   51
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
         Index           =   33
         Left            =   6300
         TabIndex        =   52
         Top             =   5100
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
         Left            =   5460
         TabIndex        =   53
         Top             =   3180
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
         Index           =   35
         Left            =   5460
         TabIndex        =   54
         Top             =   3720
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
         Index           =   36
         Left            =   4140
         TabIndex        =   55
         Top             =   5280
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
         Index           =   40
         Left            =   8520
         TabIndex        =   59
         Top             =   240
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   41
         Left            =   8880
         TabIndex        =   60
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   42
         Left            =   8160
         TabIndex        =   61
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   43
         Left            =   8580
         TabIndex        =   62
         Top             =   840
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
         Height          =   675
         Index           =   44
         Left            =   11580
         TabIndex        =   63
         Top             =   3120
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   1191
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   70
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   28
         Left            =   11040
         TabIndex        =   65
         Top             =   240
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   37
         Left            =   10680
         TabIndex        =   66
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   38
         Left            =   11400
         TabIndex        =   67
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   39
         Left            =   11100
         TabIndex        =   68
         Top             =   660
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
         Index           =   45
         Left            =   3060
         TabIndex        =   70
         Top             =   4060
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   46
         Left            =   2220
         TabIndex        =   71
         Top             =   4060
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   47
         Left            =   2700
         TabIndex        =   72
         Top             =   4560
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
         Index           =   48
         Left            =   6360
         TabIndex        =   74
         Top             =   4680
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
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   49
         Left            =   6360
         TabIndex        =   75
         Top             =   5580
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
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   12
         Left            =   7440
         TabIndex        =   99
         Top             =   6480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   51
         Left            =   6960
         TabIndex        =   100
         Top             =   7380
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Frame VV_frame 
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   4
         Left            =   2340
         TabIndex        =   110
         Top             =   3960
         Width           =   795
         Begin VB.OLE OLE_atomising 
            Appearance      =   0  'Flat
            AutoActivate    =   3  'Automatic
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Class           =   "Visio.Drawing.5"
            Enabled         =   0   'False
            Height          =   525
            Index           =   0
            Left            =   240
            OleObjectBlob   =   "OLD_Feed_Water.frx":0000
            OLETypeAllowed  =   1  'Embedded
            SizeMode        =   1  'Stretch
            SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
            TabIndex        =   111
            Top             =   60
            Width           =   510
         End
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   9240
         Picture         =   "OLD_Feed_Water.frx":8018
         Stretch         =   -1  'True
         Top             =   2700
         Width           =   255
      End
      Begin VB.OLE OLE_atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Index           =   4
         Left            =   180
         OleObjectBlob   =   "OLD_Feed_Water.frx":8322
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   118
         Top             =   4140
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE OLE_atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Index           =   3
         Left            =   180
         OleObjectBlob   =   "OLD_Feed_Water.frx":1033A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   117
         Top             =   3480
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE OLE_atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   180
         OleObjectBlob   =   "OLD_Feed_Water.frx":18552
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   116
         Top             =   2820
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE OLE8 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   465
         Index           =   8
         Left            =   4900
         OleObjectBlob   =   "OLD_Feed_Water.frx":20F6A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   115
         Top             =   2040
         Width           =   495
      End
      Begin VB.OLE OLE8 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   465
         Index           =   7
         Left            =   3600
         OleObjectBlob   =   "OLD_Feed_Water.frx":28F82
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   114
         Top             =   1860
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE OLE8 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   465
         Index           =   6
         Left            =   3600
         OleObjectBlob   =   "OLD_Feed_Water.frx":3179A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   113
         Top             =   1380
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE OLE8 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   465
         Index           =   5
         Left            =   3600
         OleObjectBlob   =   "OLD_Feed_Water.frx":397B2
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   112
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   4920
         Picture         =   "OLD_Feed_Water.frx":415CA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1935
         Index           =   1
         Left            =   4995
         Top             =   60
         Width           =   75
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   16
         X1              =   2820
         X2              =   2820
         Y1              =   3900
         Y2              =   4020
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   3420
         X2              =   2820
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   7740
         Picture         =   "OLD_Feed_Water.frx":418D4
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   255
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   1
         X1              =   7440
         X2              =   8900
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   9360
         X2              =   9360
         Y1              =   1620
         Y2              =   4260
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   10140
         Picture         =   "OLD_Feed_Water.frx":41BDE
         Stretch         =   -1  'True
         Top             =   6480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   17
         X1              =   7920
         X2              =   9960
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   15
         X1              =   5460
         X2              =   5460
         Y1              =   2280
         Y2              =   2700
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   9660
         TabIndex        =   105
         Top             =   3120
         Width           =   195
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   8760
         TabIndex        =   104
         Top             =   3120
         Width           =   195
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   14
         X1              =   5460
         X2              =   5280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   13
         X1              =   5460
         X2              =   5460
         Y1              =   2640
         Y2              =   2280
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   12
         X1              =   5460
         X2              =   5280
         Y1              =   2700
         Y2              =   2940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   7500
         X2              =   7500
         Y1              =   7860
         Y2              =   7080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   26
         X1              =   7020
         X2              =   7020
         Y1              =   7860
         Y2              =   6960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   25
         X1              =   7500
         X2              =   7500
         Y1              =   6840
         Y2              =   6360
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   11
         X1              =   2880
         X2              =   2880
         Y1              =   3960
         Y2              =   4020
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   10
         X1              =   3720
         X2              =   3420
         Y1              =   4140
         Y2              =   3900
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   7
         Left            =   9420
         Top             =   540
         Width           =   75
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   2
         X1              =   7440
         X2              =   8040
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label Label7 
         Caption         =   "Turbo Feed Pump Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   93
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   0
         X1              =   180
         X2              =   840
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   9
         Left            =   5040
         Top             =   3060
         Width           =   675
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   8
         Left            =   5040
         Top             =   1560
         Width           =   675
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   6
         Left            =   5640
         Top             =   1560
         Width           =   75
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   8
         X1              =   11220
         X2              =   11220
         Y1              =   180
         Y2              =   420
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   11220
         X2              =   10680
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   10380
         X2              =   10680
         Y1              =   420
         Y2              =   180
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   9180
         X2              =   10740
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   9780
         X2              =   9540
         Y1              =   420
         Y2              =   180
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   8700
         X2              =   8700
         Y1              =   180
         Y2              =   300
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   9540
         X2              =   8700
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   9
         X1              =   4560
         X2              =   4560
         Y1              =   3660
         Y2              =   3180
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "35 #"
         Height          =   315
         Left            =   9780
         TabIndex        =   91
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "150 #"
         Height          =   315
         Left            =   4560
         TabIndex        =   90
         Top             =   3240
         Width           =   435
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   3900
         X2              =   3840
         Y1              =   2820
         Y2              =   2880
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   4800
         X2              =   4620
         Y1              =   3060
         Y2              =   2820
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   4620
         X2              =   3840
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   5
         Left            =   7620
         Top             =   60
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   4
         Left            =   7920
         Top             =   540
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   7920
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   5040
         Top             =   500
         Width           =   3195
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   180
         Top             =   1860
         Width           =   675
      End
      Begin VB.Label Label32 
         Caption         =   "Bled Steam"
         Height          =   195
         Left            =   1020
         TabIndex        =   87
         Top             =   1800
         Width           =   1635
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   10500
         X2              =   11220
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   12060
         Picture         =   "OLD_Feed_Water.frx":41EE8
         Stretch         =   -1  'True
         Top             =   420
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   15
         X1              =   11700
         X2              =   12600
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "KEY"
         Height          =   195
         Left            =   1320
         TabIndex        =   86
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "  35 psi  Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   85
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "150 psi  Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   84
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "875 psi  Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   83
         Top             =   600
         Width           =   2175
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   24
         X1              =   840
         X2              =   180
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   8
         X1              =   840
         X2              =   180
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   14
         X1              =   180
         X2              =   840
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   2
         X1              =   3600
         X2              =   3000
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   4140
         Picture         =   "OLD_Feed_Water.frx":421F2
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   255
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   1
         X1              =   4380
         X2              =   3240
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   12
         X1              =   5040
         X2              =   5040
         Y1              =   6420
         Y2              =   2820
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   4560
         X2              =   3900
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   1500
         Picture         =   "OLD_Feed_Water.frx":424FC
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   255
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   1
         X1              =   1560
         X2              =   3420
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   7920
         TabIndex        =   82
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5760
         TabIndex        =   81
         Top             =   60
         Width           =   195
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5340
         TabIndex        =   80
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Water level"
         Height          =   495
         Left            =   11940
         TabIndex        =   79
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   6840
         TabIndex        =   78
         Top             =   5340
         Width           =   195
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   1800
         Picture         =   "OLD_Feed_Water.frx":42806
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "875 #"
         Height          =   255
         Left            =   960
         TabIndex        =   77
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   6840
         TabIndex        =   76
         Top             =   4440
         Width           =   195
      End
      Begin VB.Label Label20 
         Caption         =   "Atomising Steam to  burners"
         Height          =   675
         Left            =   3840
         TabIndex        =   73
         Top             =   4440
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   28
         X1              =   2820
         X2              =   2100
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   27
         X1              =   3600
         X2              =   3600
         Y1              =   4860
         Y2              =   4320
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   26
         X1              =   2100
         X2              =   2100
         Y1              =   4860
         Y2              =   4320
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   25
         X1              =   2400
         X2              =   1920
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   1920
         X2              =   1920
         Y1              =   4320
         Y2              =   3180
      End
      Begin VB.Label Label19 
         Caption         =   "875 # / 35 #"
         Height          =   255
         Left            =   11700
         TabIndex        =   69
         Top             =   60
         Width           =   1035
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   23
         X1              =   11940
         X2              =   11400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   22
         X1              =   11940
         X2              =   11940
         Y1              =   960
         Y2              =   540
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   21
         X1              =   10500
         X2              =   10500
         Y1              =   960
         Y2              =   540
      End
      Begin VB.Label Label15 
         Caption         =   "To DFT"
         Height          =   255
         Left            =   5760
         TabIndex        =   64
         Top             =   8280
         Width           =   675
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   5640
         Picture         =   "OLD_Feed_Water.frx":42B10
         Stretch         =   -1  'True
         Top             =   8100
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   24
         X1              =   5460
         X2              =   5940
         Y1              =   8220
         Y2              =   8220
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5280
         Top             =   8100
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   23
         X1              =   5340
         X2              =   5340
         Y1              =   8280
         Y2              =   7980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   22
         X1              =   5340
         X2              =   5340
         Y1              =   7800
         Y2              =   7680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   11
         X1              =   10920
         X2              =   11580
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   11160
         X2              =   11580
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   6060
         Picture         =   "OLD_Feed_Water.frx":42E1A
         Stretch         =   -1  'True
         Top             =   3900
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   6060
         Picture         =   "OLD_Feed_Water.frx":43124
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   7
         X1              =   5040
         X2              =   6420
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label13 
         Caption         =   "To Main Air Ejectors"
         Height          =   375
         Left            =   6000
         TabIndex        =   57
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "To Gland Seal Regulator"
         Height          =   435
         Left            =   3240
         TabIndex        =   56
         Top             =   5760
         Width           =   1755
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   3720
         Picture         =   "OLD_Feed_Water.frx":4342E
         Stretch         =   -1  'True
         Top             =   5460
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   11
         X1              =   3660
         X2              =   5040
         Y1              =   5580
         Y2              =   5580
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   8340
         Picture         =   "OLD_Feed_Water.frx":43738
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   16
         X1              =   7620
         X2              =   7620
         Y1              =   5880
         Y2              =   5400
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   15
         X1              =   5940
         X2              =   5940
         Y1              =   5880
         Y2              =   5400
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   6
         X1              =   5940
         X2              =   7620
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   5940
         X2              =   7620
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   14
         X1              =   7620
         X2              =   7620
         Y1              =   4980
         Y2              =   4380
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   12
         X1              =   5940
         X2              =   5940
         Y1              =   4980
         Y2              =   4380
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   10
         X1              =   3660
         X2              =   3120
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   3120
         X2              =   3120
         Y1              =   3660
         Y2              =   3180
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   9840
         Picture         =   "OLD_Feed_Water.frx":43A42
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Feed pump reciculation to Deaerator"
         Height          =   675
         Left            =   9480
         TabIndex        =   38
         Top             =   5220
         Width           =   1035
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   1620
         Picture         =   "OLD_Feed_Water.frx":43D4C
         Stretch         =   -1  'True
         Top             =   6540
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   1
         X1              =   6300
         X2              =   5760
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   9
         X1              =   1620
         X2              =   1740
         Y1              =   3840
         Y2              =   4320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   9
         X1              =   1680
         X2              =   1560
         Y1              =   3780
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   9
         X1              =   1440
         X2              =   4500
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Label Label11 
         Caption         =   "Main Feed System to Ecomomiser"
         Height          =   555
         Left            =   1860
         TabIndex        =   36
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label Label10 
         Caption         =   "Auxiliary Feed System to boiler (direct)"
         Height          =   615
         Left            =   1800
         TabIndex        =   35
         Top             =   7560
         Width           =   1635
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4950
         Top             =   6780
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   5160
         X2              =   5880
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   7
         X1              =   4080
         X2              =   4920
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   50
         X1              =   4920
         X2              =   4920
         Y1              =   6660
         Y2              =   7320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   51
         X1              =   5160
         X2              =   5160
         Y1              =   6660
         Y2              =   7320
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "H.P. Extraction"
         Height          =   315
         Left            =   3420
         TabIndex        =   34
         Top             =   900
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   6
         X1              =   6000
         X2              =   6000
         Y1              =   6960
         Y2              =   6360
      End
      Begin VB.Label Label1 
         Caption         =   "875 psi"
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   5160
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "De-Aerator Feed Tank"
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
         Height          =   435
         Left            =   10140
         TabIndex        =   29
         Top             =   2460
         Width           =   975
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   2
         Left            =   8220
         Top             =   5880
         Width           =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   8460
         X2              =   10620
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   10140
         Picture         =   "OLD_Feed_Water.frx":44056
         Stretch         =   -1  'True
         Top             =   7080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   10200
         X2              =   8460
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   11100
         Picture         =   "OLD_Feed_Water.frx":44360
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   5
         X1              =   9420
         X2              =   10620
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   10620
         X2              =   10620
         Y1              =   4980
         Y2              =   3780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   10920
         X2              =   11700
         Y1              =   1815
         Y2              =   1800
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   10080
         Picture         =   "OLD_Feed_Water.frx":4466A
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   360
         Left            =   10620
         Picture         =   "OLD_Feed_Water.frx":44974
         Stretch         =   -1  'True
         Top             =   1215
         Width           =   285
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   10200
         X2              =   10200
         Y1              =   1980
         Y2              =   540
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   8460
         X2              =   8460
         Y1              =   1620
         Y2              =   5520
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   8
         X1              =   5040
         X2              =   4020
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   43
         X1              =   10620
         X2              =   11520
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Cascade"
         Height          =   195
         Left            =   5580
         TabIndex        =   28
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "875 # / 150 #"
         Height          =   195
         Left            =   3360
         TabIndex        =   27
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "IP Extraction"
         Height          =   195
         Left            =   6480
         TabIndex        =   26
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Auxiliary Steam system, Exhaust from Turbo feed pumps"
         Height          =   855
         Index           =   1
         Left            =   7140
         TabIndex        =   25
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   61
         X1              =   10620
         X2              =   10620
         Y1              =   8100
         Y2              =   5220
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   4140
         Picture         =   "OLD_Feed_Water.frx":44C7E
         Stretch         =   -1  'True
         Top             =   7740
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "No 2"
         Height          =   255
         Left            =   8520
         TabIndex        =   12
         Top             =   7380
         Width           =   555
      End
      Begin VB.Label Label17 
         Caption         =   "No 1"
         Height          =   255
         Left            =   9420
         TabIndex        =   11
         Top             =   6780
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Turbo Feed Pumps"
         Height          =   735
         Index           =   0
         Left            =   8640
         TabIndex        =   10
         Top             =   3420
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   5280
         Picture         =   "OLD_Feed_Water.frx":44F88
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   15
         X1              =   7920
         X2              =   7920
         Y1              =   6480
         Y2              =   6960
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   5760
         X2              =   5760
         Y1              =   5880
         Y2              =   4980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   21
         X1              =   9960
         X2              =   9960
         Y1              =   6000
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   20
         X1              =   8820
         X2              =   8820
         Y1              =   6360
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   19
         X1              =   5400
         X2              =   6000
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   18
         X1              =   6000
         X2              =   8160
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   16
         X1              =   9660
         X2              =   10620
         Y1              =   8100
         Y2              =   8100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   14
         X1              =   6720
         X2              =   8820
         Y1              =   7860
         Y2              =   7860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   13
         X1              =   6000
         X2              =   9060
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   12
         X1              =   7920
         X2              =   7920
         Y1              =   6240
         Y2              =   6000
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   1
         X1              =   5220
         X2              =   5760
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   3300
         X2              =   6780
         Y1              =   7860
         Y2              =   7860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   9960
         X2              =   10380
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label2 
         Caption         =   "Electric Motor"
         Height          =   255
         Left            =   10320
         TabIndex        =   5
         Top             =   8400
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1395
         Index           =   1
         Left            =   9120
         Top             =   4980
         Width           =   75
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   13
         X1              =   5760
         X2              =   6300
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   10080
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   5
         Left            =   10080
         Shape           =   3  'Circle
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   10260
         Shape           =   3  'Circle
         Top             =   1575
         Width           =   735
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   7140
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   4800
         Top             =   6420
         Width           =   555
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   9
         X1              =   5040
         X2              =   6420
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Label Label14 
         Caption         =   "To Distil Air Ejector"
         Height          =   375
         Left            =   6000
         TabIndex        =   58
         Top             =   3720
         Width           =   1575
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   975
         Left            =   8700
         OleObjectBlob   =   "OLD_Feed_Water.frx":45292
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisFdPmp.vsd"
         TabIndex        =   108
         Top             =   4200
         Width           =   915
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   975
         Left            =   7800
         OleObjectBlob   =   "OLD_Feed_Water.frx":4CCAA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisFdPmp.vsd"
         TabIndex        =   109
         Top             =   5100
         Width           =   915
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2100
      TabIndex        =   17
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "OLD_Feed_Water"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command2_Click()

Text6.text = JYAscrsizeOrig(idFeed_water).objcount

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idFeed_water)

For i = 5 To 7
    OLE8(i).Left = OLE8(8).Left
    OLE8(i).Top = OLE8(8).Top
Next i

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then

    HavenStretch1.Visible = False
    
    
    Call JYAScreenResize(Me, idFeed_water)
    
    
    HavenStretch1.Visible = True
    
    
    For i = 5 To 7
        OLE8(i).Left = OLE8(8).Left
        OLE8(i).Top = OLE8(8).Top
    Next i
    
   
End If

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 3))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 3))
End If

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(5=Stand-bye,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 3))
End If

End Sub

Private Sub Timer1_Timer()

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

P_feed_water.text = Format(LSS_reals(791), "#0.0") & " bar"
'List8.AddItem "Pfeed_water " & Format(LSS_reals(791), "#####0.00")

' Note: update_pump_LSS(syxxx,index,nPump)
Call update_pump_LSS(Feed_Water, 0, 3)

' update_vvalve_LSS(syxxx,nValve,index1,index2)
Call update_vvalve_LSS(Feed_Water, 20, 0, -1)
Call update_vvalve_LSS(Feed_Water, 19, 2, -1)


End Sub

Private Sub valve_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
  

Select Case Index
  Case 0
    Call valve_pcdi_LSS(20)
  Case 2
    Call valve_pcdi_LSS(19)
    
End Select


End Sub

Private Sub VScroll2_Change()
Dim i As Integer

If VScroll2.Value = 100 Then
    
    OLE8(5).Visible = True
    OLE8(6).Visible = False
    OLE8(7).Visible = False
    OLE8(8).Visible = False
    
ElseIf VScroll2.Value > 50 Then
    
    OLE8(5).Visible = False
    OLE8(6).Visible = True
    OLE8(7).Visible = False
    OLE8(8).Visible = False

    'TESTING
    OLE_atomising(0).Copy
    OLE_atomising(2).Paste
    
ElseIf VScroll2.Value > 0 Then
    
    OLE8(5).Visible = False
    OLE8(6).Visible = False
    OLE8(7).Visible = True
    OLE8(8).Visible = False
    
Else
    
    OLE8(5).Visible = False
    OLE8(6).Visible = False
    OLE8(7).Visible = False
    OLE8(8).Visible = True
    
    
End If


For i = 5 To 7
    OLE8(i).Left = OLE8(8).Left
    OLE8(i).Top = OLE8(8).Top
Next i

End Sub

