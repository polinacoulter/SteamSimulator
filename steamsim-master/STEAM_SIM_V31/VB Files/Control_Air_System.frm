VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Control_Air_System 
   Caption         =   "Control Air System"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   12945
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9255
      Left            =   0
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   16325
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   180
         TabIndex        =   101
         Top             =   6720
         Visible         =   0   'False
         Width           =   1875
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
         Index           =   1
         Left            =   180
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1560
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
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
         Index           =   1
         Left            =   720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Index           =   1
         Left            =   720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox PCA_70 
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
         Left            =   2235
         TabIndex        =   43
         Text            =   "xx psi"
         Top             =   6240
         Width           =   735
      End
      Begin VB.TextBox PCA_20 
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
         Left            =   5440
         TabIndex        =   42
         Text            =   "xx psi"
         Top             =   3540
         Width           =   735
      End
      Begin VB.TextBox PCA_pump_out 
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
         Left            =   800
         TabIndex        =   41
         Text            =   "xx psi"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox PCA_bottle 
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
         Left            =   2400
         TabIndex        =   40
         Text            =   "xx psi"
         Top             =   1660
         Width           =   735
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
         Left            =   2100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3600
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
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
         Left            =   1500
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4260
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Left            =   1500
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4020
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   360
         Top             =   5700
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   17
         Left            =   1230
         TabIndex        =   7
         Top             =   2430
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   18
         Left            =   2580
         TabIndex        =   8
         Top             =   2430
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   0
         Left            =   5400
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   1
         Left            =   2340
         TabIndex        =   26
         Top             =   4740
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   13
         Left            =   3300
         TabIndex        =   29
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
         Index           =   14
         Left            =   3300
         TabIndex        =   30
         Top             =   5220
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
         Index           =   15
         Left            =   3300
         TabIndex        =   31
         Top             =   5760
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
         Index           =   16
         Left            =   3300
         TabIndex        =   32
         Top             =   6180
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
         Height          =   315
         Index           =   19
         Left            =   3450
         TabIndex        =   33
         Top             =   3405
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   20
         Left            =   3450
         TabIndex        =   34
         Top             =   1380
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin PumpPrj.Pump Havenpump1 
         Height          =   405
         Index           =   0
         Left            =   2220
         TabIndex        =   38
         Top             =   3900
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         CentrifugalPump =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   2
         Left            =   6420
         TabIndex        =   61
         Top             =   480
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
         Index           =   3
         Left            =   6420
         TabIndex        =   62
         Top             =   900
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
         Index           =   4
         Left            =   6420
         TabIndex        =   63
         Top             =   1380
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
         Index           =   5
         Left            =   6420
         TabIndex        =   64
         Top             =   1860
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
         Index           =   6
         Left            =   6420
         TabIndex        =   65
         Top             =   2460
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
         Index           =   7
         Left            =   6420
         TabIndex        =   66
         Top             =   2940
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
         Index           =   8
         Left            =   6420
         TabIndex        =   67
         Top             =   3480
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
         Index           =   9
         Left            =   6420
         TabIndex        =   68
         Top             =   3960
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
         Index           =   10
         Left            =   6420
         TabIndex        =   69
         Top             =   4440
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
         Index           =   11
         Left            =   6420
         TabIndex        =   70
         Top             =   4980
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
         Index           =   12
         Left            =   6420
         TabIndex        =   71
         Top             =   5490
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
         Index           =   21
         Left            =   6420
         TabIndex        =   72
         Top             =   6000
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
         Index           =   22
         Left            =   6420
         TabIndex        =   73
         Top             =   6510
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
         Index           =   23
         Left            =   6420
         TabIndex        =   74
         Top             =   7080
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
         Index           =   24
         Left            =   6420
         TabIndex        =   75
         Top             =   7560
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
         Index           =   25
         Left            =   6420
         TabIndex        =   83
         Top             =   8080
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
         Index           =   26
         Left            =   1980
         TabIndex        =   88
         Top             =   2040
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump Havenpump1 
         Height          =   405
         Index           =   1
         Left            =   180
         TabIndex        =   90
         Top             =   1080
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         CentrifugalPump =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   27
         Left            =   6420
         TabIndex        =   91
         Top             =   8520
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
         Index           =   28
         Left            =   4800
         TabIndex        =   93
         Top             =   6900
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
         Left            =   4800
         TabIndex        =   94
         Top             =   7440
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
         Left            =   4800
         TabIndex        =   95
         Top             =   7920
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
         Left            =   4800
         TabIndex        =   96
         Top             =   8460
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "To PORT Drum Level Transmitter"
         Height          =   255
         Left            =   1380
         TabIndex        =   100
         Top             =   8640
         Width           =   2925
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "To STBD Drum Level Transmitter"
         Height          =   255
         Left            =   1380
         TabIndex        =   99
         Top             =   7620
         Width           =   2925
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "To PORT Steam Flow Transmitter"
         Height          =   255
         Left            =   1380
         TabIndex        =   98
         Top             =   8100
         Width           =   2925
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "To STBD Steam Flow Transmitter"
         Height          =   255
         Left            =   1380
         TabIndex        =   97
         Top             =   7080
         Width           =   2925
      End
      Begin VB.Image Image28 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4380
         Picture         =   "Control_Air_System.frx":0000
         Stretch         =   -1  'True
         Top             =   8640
         Width           =   255
      End
      Begin VB.Image Image27 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4380
         Picture         =   "Control_Air_System.frx":030A
         Stretch         =   -1  'True
         Top             =   8100
         Width           =   255
      End
      Begin VB.Image Image26 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4380
         Picture         =   "Control_Air_System.frx":0614
         Stretch         =   -1  'True
         Top             =   7620
         Width           =   255
      End
      Begin VB.Image Image32 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4380
         Picture         =   "Control_Air_System.frx":091E
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   255
      End
      Begin VB.Line Line44 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   4440
         X2              =   5310
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   4440
         X2              =   5310
         Y1              =   8220
         Y2              =   8220
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   4440
         X2              =   5310
         Y1              =   7740
         Y2              =   7740
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   4440
         X2              =   5310
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   7200
         Y2              =   8760
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   5340
         X2              =   6210
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Label Label26 
         Caption         =   "To Fuel Oil Un-loader"
         Height          =   315
         Left            =   7380
         TabIndex        =   92
         Top             =   8700
         Width           =   2355
      End
      Begin VB.Image Image25 
         Height          =   255
         Left            =   6960
         Picture         =   "Control_Air_System.frx":0C28
         Stretch         =   -1  'True
         Top             =   8700
         Width           =   255
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7170
         Y1              =   8820
         Y2              =   8820
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":0F32
         Stretch         =   -1  'True
         Top             =   4620
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":123C
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":1546
         Stretch         =   -1  'True
         Top             =   5700
         Width           =   255
      End
      Begin VB.Line Service_line 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         Index           =   3
         X1              =   3000
         X2              =   3000
         Y1              =   2340
         Y2              =   2760
      End
      Begin VB.Label Service_lbl 
         Alignment       =   2  'Center
         Caption         =   "Service Air"
         Height          =   315
         Left            =   540
         TabIndex        =   87
         Top             =   780
         Width           =   975
      End
      Begin VB.Line Service_line 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         Index           =   2
         X1              =   1080
         X2              =   1680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Service_line 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   1680
         X2              =   1680
         Y1              =   1320
         Y2              =   2340
      End
      Begin VB.Line Service_line 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   1680
         X2              =   3000
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label25 
         Caption         =   "To 875/225 + Primary Steam + FO Heating Regulators"
         Height          =   315
         Left            =   7380
         TabIndex        =   84
         Top             =   8280
         Width           =   4275
      End
      Begin VB.Image Image24 
         Height          =   255
         Left            =   6960
         Picture         =   "Control_Air_System.frx":1850
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   255
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7170
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Label Label61 
         Caption         =   "xx.x psi"
         Height          =   315
         Left            =   9900
         TabIndex        =   82
         Top             =   7680
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label60 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   81
         Top             =   7200
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label59 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   80
         Top             =   6660
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label58 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   79
         Top             =   6120
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label57 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   78
         Top             =   5580
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label56 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   77
         Top             =   5100
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label55 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   76
         Top             =   4620
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Image Image23 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":1B5A
         Stretch         =   -1  'True
         Top             =   7740
         Width           =   255
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":1E64
         Stretch         =   -1  'True
         Top             =   7260
         Width           =   255
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7170
         Y1              =   7860
         Y2              =   7860
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7170
         Y1              =   7380
         Y2              =   7380
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7110
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7140
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7140
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7170
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":216E
         Stretch         =   -1  'True
         Top             =   660
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":2478
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":2782
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":2A8C
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":2D96
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":30A0
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":33AA
         Stretch         =   -1  'True
         Top             =   3660
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":36B4
         Stretch         =   -1  'True
         Top             =   4140
         Width           =   255
      End
      Begin VB.Image Image20 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":39BE
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   255
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   7020
         Picture         =   "Control_Air_System.frx":3CC8
         Stretch         =   -1  'True
         Top             =   6180
         Width           =   255
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7170
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   6300
         Y2              =   6300
      End
      Begin VB.Label Label46 
         Caption         =   "To Stbd Attemporator Regulator"
         Height          =   315
         Left            =   7380
         TabIndex        =   60
         Top             =   7740
         Width           =   2355
      End
      Begin VB.Label Label45 
         Caption         =   "To Port Attemporator Regulator"
         Height          =   315
         Left            =   7380
         TabIndex        =   59
         Top             =   7260
         Width           =   2355
      End
      Begin VB.Label Label44 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   58
         Top             =   4020
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label43 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   57
         Top             =   3540
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label42 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   56
         Top             =   3060
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label41 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   55
         Top             =   2580
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label40 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   54
         Top             =   2100
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label38 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   53
         Top             =   1620
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label37 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   52
         Top             =   1140
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label36 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   9900
         TabIndex        =   51
         Top             =   660
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label39 
         Caption         =   "xx.x psi"
         Height          =   255
         Left            =   60
         TabIndex        =   50
         Top             =   2160
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label35 
         Height          =   255
         Left            =   600
         TabIndex        =   49
         Top             =   60
         Width           =   675
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Diaphragm Pressure"
         Height          =   435
         Left            =   9720
         TabIndex        =   48
         Top             =   60
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   2280
         OleObjectBlob   =   "Control_Air_System.frx":3FD2
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Reducing_VV3.vsd"
         TabIndex        =   47
         Top             =   5160
         Width           =   405
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   5340
         OleObjectBlob   =   "Control_Air_System.frx":6BEA
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Reducing_VV3.vsd"
         TabIndex        =   46
         Top             =   2625
         Width           =   405
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3060
         X2              =   3060
         Y1              =   6480
         Y2              =   4980
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2160
         X2              =   2160
         Y1              =   5520
         Y2              =   4440
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2160
         X2              =   5220
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label24 
         Caption         =   "To Port Drum Level Regulator"
         Height          =   225
         Left            =   7380
         TabIndex        =   45
         Top             =   6210
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "To Stbd Drum Level Regulator"
         Height          =   315
         Left            =   7380
         TabIndex        =   44
         Top             =   6690
         Width           =   2175
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   5100
         Picture         =   "Control_Air_System.frx":9802
         Stretch         =   -1  'True
         Top             =   4140
         Width           =   255
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3540
         X2              =   3540
         Y1              =   1620
         Y2              =   1020
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3540
         X2              =   3540
         Y1              =   4020
         Y2              =   3420
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3150
         Top             =   2010
         Width           =   60
      End
      Begin VB.Line Line24 
         BorderStyle     =   2  'Dash
         X1              =   690
         X2              =   690
         Y1              =   2070
         Y2              =   3150
      End
      Begin VB.Line Line23 
         BorderStyle     =   2  'Dash
         X1              =   3240
         X2              =   690
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Label Label17 
         Caption         =   "To Port Boiler Front"
         Height          =   255
         Left            =   4140
         TabIndex        =   27
         Top             =   4950
         Width           =   1845
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   5880
         Picture         =   "Control_Air_System.frx":9B0C
         Stretch         =   -1  'True
         Top             =   2835
         Width           =   255
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   2760
         Picture         =   "Control_Air_System.frx":9E16
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   255
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   5220
         X2              =   5220
         Y1              =   4440
         Y2              =   2100
      End
      Begin VB.Line Line31 
         X1              =   1500
         X2              =   120
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Label Label22 
         Caption         =   "70 psi lines"
         Height          =   405
         Left            =   2460
         TabIndex        =   24
         Top             =   5820
         Width           =   465
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "20 psi lines"
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   5700
         TabIndex        =   23
         Top             =   3120
         Width           =   525
      End
      Begin VB.Label Label20 
         Caption         =   "To Stbd Air Damper Drive"
         Height          =   345
         Left            =   4140
         TabIndex        =   22
         Top             =   6450
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "To Port Air Damper Drive"
         Height          =   285
         Left            =   4140
         TabIndex        =   21
         Top             =   5940
         Width           =   1965
      End
      Begin VB.Label Label18 
         Caption         =   "To Stbd Boiler Front"
         Height          =   195
         Left            =   4140
         TabIndex        =   20
         Top             =   5430
         Width           =   1905
      End
      Begin VB.Label Label16 
         Caption         =   "No 1 Feed Pump Steam Reg"
         Height          =   315
         Left            =   7380
         TabIndex        =   19
         Top             =   5670
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "No 2 Feed Pump Steam Reg"
         Height          =   225
         Left            =   7380
         TabIndex        =   18
         Top             =   5190
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "To Gland Seal Regulator"
         Height          =   225
         Left            =   7380
         TabIndex        =   17
         Top             =   4680
         Width           =   2205
      End
      Begin VB.Label Label13 
         Caption         =   "To 35 psig Dump Regulator"
         Height          =   225
         Left            =   7380
         TabIndex        =   16
         Top             =   4140
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "To Condensate Recirc. Reg."
         Height          =   255
         Left            =   7380
         TabIndex        =   15
         Top             =   3660
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "To Make up Feed Regulator"
         Height          =   225
         Left            =   7380
         TabIndex        =   14
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "To Dearator Dump Regulator"
         Height          =   255
         Left            =   7380
         TabIndex        =   13
         Top             =   2640
         Width           =   2115
      End
      Begin VB.Label Label9 
         Caption         =   "To 875/35 psig Regulator"
         Height          =   345
         Left            =   7380
         TabIndex        =   12
         Top             =   2100
         Width           =   2145
      End
      Begin VB.Label Label8 
         Caption         =   "To 875/150 psig Regulator"
         Height          =   255
         Left            =   7380
         TabIndex        =   11
         Top             =   1590
         Width           =   1995
      End
      Begin VB.Label Label7 
         Caption         =   "To IP Extraction Regulator"
         Height          =   255
         Left            =   7380
         TabIndex        =   10
         Top             =   1140
         Width           =   2085
      End
      Begin VB.Label Label6 
         Caption         =   "To HP Extraction Regulator"
         Height          =   255
         Left            =   7380
         TabIndex        =   9
         Top             =   690
         Width           =   2205
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   3840
         Picture         =   "Control_Air_System.frx":A120
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   3840
         Picture         =   "Control_Air_System.frx":A42A
         Stretch         =   -1  'True
         Top             =   5910
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   3840
         Picture         =   "Control_Air_System.frx":A734
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   3840
         Picture         =   "Control_Air_System.frx":AA3E
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   6120
         Picture         =   "Control_Air_System.frx":AD48
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "De-Humidifier"
         Height          =   285
         Left            =   4140
         TabIndex        =   6
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Air Receiver"
         Height          =   465
         Left            =   3180
         TabIndex        =   5
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "After Cooler"
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   3210
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Air Compressor"
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   4860
         Width           =   1155
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3060
         X2              =   3900
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3030
         X2              =   3990
         Y1              =   6030
         Y2              =   6030
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3030
         X2              =   3930
         Y1              =   4980
         Y2              =   5010
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   6240
         Y1              =   8820
         Y2              =   780
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   1725
         Left            =   2940
         OleObjectBlob   =   "Control_Air_System.frx":B052
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Air_receiver.vsd"
         TabIndex        =   2
         Top             =   1680
         Width           =   1185
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1620
         OleObjectBlob   =   "Control_Air_System.frx":D46A
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Cooler.vsd"
         TabIndex        =   1
         Top             =   2580
         Width           =   945
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   4380
         Shape           =   4  'Rounded Rectangle
         Top             =   1980
         Width           =   615
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3900
         X2              =   4650
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2550
         X2              =   3210
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   960
         X2              =   1740
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   960
         X2              =   960
         Y1              =   3120
         Y2              =   2760
      End
      Begin VB.OLE Compressor 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   1725
         Left            =   240
         OleObjectBlob   =   "Control_Air_System.frx":F882
         SizeMode        =   1  'Stretch
         TabIndex        =   0
         Top             =   3060
         Width           =   1185
      End
      Begin VB.Line Line6 
         X1              =   1650
         X2              =   4380
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   4890
         X2              =   5220
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label23 
         Caption         =   "Pressure Unloading switch"
         Height          =   585
         Left            =   30
         TabIndex        =   28
         Top             =   2430
         Width           =   765
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   5220
         X2              =   6240
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2160
         X2              =   3960
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7200
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   6240
         X2              =   7260
         Y1              =   780
         Y2              =   780
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   39
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Control_Air_System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_Control_Air)

' 24/8/00
If Not I_AM_INSTRUCTOR And Not PTT Then


  ' === 13/03/01  REMOVE ALL THIS ===
  '  ' I MUST be an ACT Terminal (Full Mission) so
  '  Service_line(0).Visible = False
  '  Service_line(1).Visible = False
  '  Service_line(2).Visible = False
  '  Service_line(3).Visible = False
  '  Service_lbl.Visible = False
  '
  '  ' === 01/03/01 Point 10.2b NOT the Service Air Compressor ===
  '  'ONpump(1).Visible = False
  '  'OFFpump(1).Visible = False
  '
  '      valve(26).Visible = False
        
Else

    Service_line(0).Visible = True
    Service_line(1).Visible = True
    Service_line(2).Visible = True
    Service_line(3).Visible = True
    Service_lbl.Visible = True
    
    ONpump(1).Visible = True
    OFFpump(1).Visible = True
    
        Valve(26).Visible = True
        
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_Control_Air)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 33))

ElseIf Index = 1 Then       ' 24/8/00  Service Air Compressor
   Call setPCDI(PumpIO_LSS(4, 29))


End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 33))
   
ElseIf Index = 1 Then       ' 24/8/00  Service Air Compressor
   Call setPCDI(PumpIO_LSS(2, 29))
   
End If

End Sub

Private Sub PCA_20_Click()



' 3/4/00 MOVED
'  '6/1/00
'    Unload Alarm_set_point   ' Do this first
'    nAlarm_LO = 32
'
'    Alarm_set_point.Show


' === 22/01/02 ===
'   Alarm(31) = PCA_20_psi  .lt. ALARM_HILO(2,31)

    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 31
    
    Alarm_set_point.Show




End Sub

Private Sub PCA_70_Click()

' === 22/01/02 ===
'        Alarm(39) = PCA_70_psi  .lt. ALARM_HILO(2,39)

    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 39
    
    Alarm_set_point.Show


End Sub

Private Sub PCA_bottle_Click()

  '3/4/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 32
    
    Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 7/1/00

' == 22/01 ===
'Call paint_alarm(Me.PCA_20, 32, 0, 0, 0)
Call paint_alarm(Me.PCA_20, 31, 0, 0, 0)

' == 22/01 ===
Call paint_alarm(Me.PCA_70, 39, 0, 0, 0)

' 3/4/00
Call paint_alarm(Me.PCA_bottle, 32, 0, 0, 0)


'          PC_REALS(320) = PCA_pump_OUT
'          PC_REALS(321) = PCA_bottle
'          PC_REALS(322) = PCA_20
'          PC_REALS(323) = PCA_70
         
         
'         PC_REALS(816) = P_CONTROL_air_loading(1) / 14.5
'         PC_REALS(817) = P_CONTROL_air_loading(2) / 14.5
'         PC_REALS(818) = P_CONTROL_air_loading(3) / 14.5
'         PC_REALS(819) = P_CONTROL_air_loading(4) / 14.5
'
'         PC_REALS(820) = P_CONTROL_air_supply(1) / 14.5
'         PC_REALS(821) = P_CONTROL_air_supply(2) / 14.5
'         PC_REALS(822) = P_CONTROL_air_supply(3) / 14.5
'         PC_REALS(823) = P_CONTROL_air_supply(4) / 14.5


' 9/12/00
List1.clear
List1.AddItem Format(LSS_reals(820) * 14.5, "##0.0") & "  " & Format(LSS_reals(816) * 14.5, "##0.0")
List1.AddItem Format(LSS_reals(821) * 14.5, "##0.0") & "  " & Format(LSS_reals(817) * 14.5, "##0.0")
List1.AddItem Format(LSS_reals(822) * 14.5, "##0.0") & "  " & Format(LSS_reals(818) * 14.5, "##0.0")
List1.AddItem Format(LSS_reals(823) * 14.5, "##0.0") & "  " & Format(LSS_reals(819) * 14.5, "##0.0")

If SI_units Then
    PCA_pump_out.text = Format(LSS_reals(320) * Press_mult, "##0.0") & Press_chars     ' Bar
    PCA_bottle.text = Format(LSS_reals(321) * Press_mult, "##0.0") & Press_chars     ' Bar
    PCA_20.text = Format(LSS_reals(322) * Press_mult, "##0.0") & Press_chars     ' Bar
    PCA_70.text = Format(LSS_reals(323) * Press_mult, "##0.0") & Press_chars     ' Bar
  
  '  FFO_recirc(1).text = Format(LSS_reals(234) * Flow_mult, "###0.000") & Flow_chars
Else
    PCA_pump_out.text = Format(LSS_reals(320) * Press_mult, "###0") & Press_chars    ' psi
    PCA_bottle.text = Format(LSS_reals(321) * Press_mult, "###0") & Press_chars    ' psi
    PCA_20.text = Format(LSS_reals(322) * Press_mult, "###0") & Press_chars    ' psi
    PCA_70.text = Format(LSS_reals(323) * Press_mult, "###0") & Press_chars    ' psi
   
   ' FFO_recirc(1).text = Format(LSS_reals(234) * Flow_mult / 3600#, "###0.000") & " lb/s"
End If


' 11/10/99
'Sub update_pump_LSS(frm As Form, index_no As Integer, pump_no As Integer)
Call update_pump_LSS(Me, 0, 33)      ' Control Air Compressor

' 24/8/00
Call update_pump_LSS(Me, 1, 29)      ' Service Air Compressor


' update_vvalve_LSS(Formname As Form, valve_number As Integer, indx1 As Integer, indx2 As Integer)

For i = 2 To 18
   Call update_vvalve_LSS(Me, 26 + i, i, -1)
Next i

' 26/10/99
   Call update_vvalve_LSS(Me, 6, 21, -1)
   Call update_vvalve_LSS(Me, 7, 22, -1)

' 28/10/99
   Call update_vvalve_LSS(Me, 47, 23, -1)
   Call update_vvalve_LSS(Me, 48, 24, -1)
   
' 8/8/00
   Call update_vvalve_LSS(Me, 49, 25, -1)

' 24/8/00
' 24/8/00 Can't convert valve 50 so leave for now..
 ' Call update_vvalve_LSS(Me, 50, 26, -1)
 
 ' 11/9/00
  Call update_vvalve_LSS(Me, 576, 26, -1)
 
 ' 15/9/00
  Call update_vvalve_LSS(Me, 578, 27, -1)
 
 
 ' 9/12/00  NEW VALVES (see Bilge Xref V6 )
  Call update_vvalve_LSS(Me, 584, 28, -1)
  Call update_vvalve_LSS(Me, 585, 29, -1)
  Call update_vvalve_LSS(Me, 586, 30, -1)
  Call update_vvalve_LSS(Me, 587, 31, -1)
 
 ' 10/12/00  NEW VALVE (see Bilge Xref V7 )
  Call update_vvalve_LSS(Me, 588, 19, -1)
   
   
End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
If Index = 19 Then
 ' 10/12/00  ADDED to Bilge Xref V7 ( so snapshots can be re-used )
    Call valve_pcdi_LSS(588)

ElseIf Index = 21 Then
    Call valve_pcdi_LSS(6)
ElseIf Index = 22 Then
    Call valve_pcdi_LSS(7)
ElseIf Index = 23 Then          ' 28/10/99
    Call valve_pcdi_LSS(47)
ElseIf Index = 24 Then
    Call valve_pcdi_LSS(48)        ' 28/10/99
ElseIf Index = 25 Then
    Call valve_pcdi_LSS(49)        ' 8/8/00
ElseIf Index = 26 Then
   
   ' 11/9/00  ADDED to Bilge Xref V3 ( so snapshots can be re-used )
     Call valve_pcdi_LSS(576)
   
ElseIf Index = 27 Then
    ' 15/9/00  ADDED to Bilge Xref V5 ( so snapshots can be re-used )
     Call valve_pcdi_LSS(578)

ElseIf Index = 28 Then
    ' 8/12/00  ADDED to Bilge Xref V6 ( so snapshots can be re-used )
     Call valve_pcdi_LSS(584)
ElseIf Index = 29 Then
    ' 8/12/00  ADDED to Bilge Xref V6 ( so snapshots can be re-used )
     Call valve_pcdi_LSS(585)
ElseIf Index = 30 Then
    ' 8/12/00  ADDED to Bilge Xref V6 ( so snapshots can be re-used )
     Call valve_pcdi_LSS(586)
ElseIf Index = 31 Then
    ' 8/12/00  ADDED to Bilge Xref V6 ( so snapshots can be re-used )
     Call valve_pcdi_LSS(587)
  
Else
    Call valve_pcdi_LSS(Index + 26)
End If

End Sub
