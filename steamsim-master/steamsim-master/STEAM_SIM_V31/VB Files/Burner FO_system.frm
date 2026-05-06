VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{A6D9C597-48C2-11D1-8915-000021240899}#2.0#0"; "Haven3WayValve v1.41.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Burner_FO_system 
   Caption         =   "Fuel To Burners"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   11085
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6855
      Left            =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12091
      Begin VB.TextBox F_DO_emerg 
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
         Left            =   240
         TabIndex        =   160
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox P_DO_emerg_out 
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
         Left            =   1200
         TabIndex        =   159
         Top             =   4200
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
         Left            =   1620
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   5760
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
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
         Left            =   600
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   5880
         UseMaskColor    =   -1  'True
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
         Left            =   1080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox TFO_Port 
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
         Left            =   3720
         TabIndex        =   134
         Text            =   "xxx F"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox TFO_Stbd 
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
         Left            =   3780
         TabIndex        =   133
         Text            =   "xxx F"
         Top             =   5820
         Width           =   855
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   3
         Left            =   8880
         Max             =   1000
         TabIndex        =   129
         Top             =   1020
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   2
         Left            =   8880
         Max             =   1000
         TabIndex        =   128
         Top             =   1800
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   1
         Left            =   8880
         Max             =   1000
         TabIndex        =   127
         Top             =   4440
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   0
         Left            =   8880
         Max             =   1000
         TabIndex        =   126
         Top             =   5220
         Width           =   1335
      End
      Begin VB.TextBox FFO_burner 
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
         Index           =   4
         Left            =   4380
         TabIndex        =   53
         Text            =   "2.1 kg/s"
         Top             =   540
         Width           =   795
      End
      Begin VB.TextBox FFO_burner 
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
         Index           =   3
         Left            =   4860
         TabIndex        =   52
         Text            =   "2.1 kg/s"
         Top             =   2280
         Width           =   795
      End
      Begin VB.TextBox FFO_burner 
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
         Left            =   4860
         TabIndex        =   51
         Text            =   "2.1 kg/s"
         Top             =   3720
         Width           =   795
      End
      Begin VB.TextBox FFO_burner 
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
         Index           =   1
         Left            =   4800
         TabIndex        =   50
         Text            =   "2.1 kg/s"
         Top             =   5880
         Width           =   795
      End
      Begin VB.TextBox PFO_B 
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
         Left            =   2280
         TabIndex        =   49
         Text            =   "85 to 450 psi"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox PFO_A 
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
         Left            =   480
         TabIndex        =   48
         Text            =   "85 to 450 psi"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox PFO_stbd 
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
         Left            =   3780
         TabIndex        =   47
         Text            =   "85 to 450 psi"
         Top             =   5520
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   480
      End
      Begin VB.HScrollBar H_fuel_control 
         Height          =   135
         Index           =   1
         Left            =   1380
         Max             =   1000
         TabIndex        =   46
         Top             =   6540
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.HScrollBar H_fuel_control 
         Height          =   135
         Index           =   2
         Left            =   960
         Max             =   1000
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox PFO_port 
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
         Left            =   3720
         TabIndex        =   11
         Text            =   "85 to 450 psi"
         Top             =   2100
         Width           =   855
      End
      Begin VB.TextBox F_flowmeter 
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
         Left            =   720
         TabIndex        =   10
         Text            =   "2.1 kg/s"
         Top             =   3480
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Fuel Oil Service"
         Height          =   555
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1260
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   2925
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   2925
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   2
         Left            =   5400
         TabIndex        =   5
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   3
         Left            =   5400
         TabIndex        =   6
         Top             =   1560
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   4
         Left            =   4860
         TabIndex        =   7
         Top             =   780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   5
         Left            =   4860
         TabIndex        =   8
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   8
         Left            =   4560
         TabIndex        =   17
         Top             =   2640
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   11
         Left            =   4860
         TabIndex        =   18
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   12
         Left            =   4860
         TabIndex        =   19
         Top             =   5100
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   15
         Left            =   4560
         TabIndex        =   20
         Top             =   3360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   16
         Left            =   3600
         TabIndex        =   27
         Top             =   1110
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   17
         Left            =   3510
         TabIndex        =   28
         Top             =   4740
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   7
         Left            =   1020
         TabIndex        =   42
         Top             =   2220
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   6
         Left            =   6600
         TabIndex        =   130
         Top             =   6360
         Visible         =   0   'False
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   9
         Left            =   5340
         TabIndex        =   131
         Top             =   4320
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   10
         Left            =   5340
         TabIndex        =   132
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
         ValveType       =   1
      End
      Begin Haven3WayValvePrj.Haven3WayValve Three_way_VV 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   139
         Top             =   1260
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         OpenBtmTri      =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin Haven3WayValvePrj.Haven3WayValve Three_way_VV 
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   140
         Top             =   4860
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         OpenTopTri      =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   141
         Top             =   5280
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   423
         Left            =   1500
         TabIndex        =   142
         Top             =   5220
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   422
         Left            =   360
         TabIndex        =   143
         Top             =   5220
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   18
         Left            =   900
         TabIndex        =   144
         Top             =   4800
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
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   421
         Left            =   480
         TabIndex        =   148
         Top             =   4500
         Visible         =   0   'False
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   424
         Left            =   2340
         TabIndex        =   150
         Top             =   1680
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   425
         Left            =   2280
         TabIndex        =   151
         Top             =   5400
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   486
         Left            =   5340
         TabIndex        =   152
         Top             =   60
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   485
         Left            =   6000
         TabIndex        =   153
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   484
         Left            =   6000
         TabIndex        =   154
         Top             =   3240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   483
         Left            =   5340
         TabIndex        =   155
         Top             =   6180
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
      Begin VB.Image Right_arrow_image 
         Height          =   255
         Left            =   7200
         Picture         =   "Burner FO_system.frx":0000
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Left_arrow_image 
         Height          =   255
         Left            =   7200
         Picture         =   "Burner FO_system.frx":030A
         Stretch         =   -1  'True
         Top             =   2580
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image BIGN2_arrow 
         Height          =   255
         Index           =   2
         Left            =   5220
         Picture         =   "Burner FO_system.frx":0614
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   255
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   6
         X1              =   4680
         X2              =   4680
         Y1              =   5425
         Y2              =   4635
      End
      Begin VB.Label M10_label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   4860
         Width           =   255
      End
      Begin VB.Image BIGN2_arrow 
         Height          =   255
         Index           =   1
         Left            =   5220
         Picture         =   "Burner FO_system.frx":091E
         Stretch         =   -1  'True
         Top             =   4770
         Width           =   255
      End
      Begin VB.Image BIGN1_arrow 
         Height          =   255
         Index           =   2
         Left            =   5220
         Picture         =   "Burner FO_system.frx":0C28
         Stretch         =   -1  'True
         Top             =   2070
         Width           =   255
      End
      Begin VB.Image BIGN1_arrow 
         Height          =   255
         Index           =   1
         Left            =   5220
         Picture         =   "Burner FO_system.frx":0F32
         Stretch         =   -1  'True
         Top             =   5620
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   240
         Picture         =   "Burner FO_system.frx":123C
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   360
         X2              =   360
         Y1              =   1740
         Y2              =   3210
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Atomising steam"
         Height          =   255
         Left            =   3840
         TabIndex        =   158
         Top             =   6300
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Atomising steam"
         Height          =   255
         Left            =   3600
         TabIndex        =   157
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Atomising steam"
         Height          =   255
         Left            =   6540
         TabIndex        =   156
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5220
         X2              =   5820
         Y1              =   6420
         Y2              =   6420
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6600
         X2              =   5820
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6600
         X2              =   5820
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4980
         X2              =   5820
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "DO Emerg Pump"
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
         TabIndex        =   147
         Top             =   6180
         Width           =   1335
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   7
         X1              =   4680
         X2              =   3840
         Y1              =   5010
         Y2              =   5010
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   60
         Picture         =   "Burner FO_system.frx":1546
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   20
         X1              =   1380
         X2              =   1380
         Y1              =   5520
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   19
         X1              =   780
         X2              =   780
         Y1              =   5100
         Y2              =   5520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   18
         X1              =   1380
         X2              =   780
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   17
         X1              =   180
         X2              =   180
         Y1              =   5520
         Y2              =   1800
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   12
         X1              =   1980
         X2              =   1980
         Y1              =   5520
         Y2              =   3360
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   16
         X1              =   2220
         X2              =   2220
         Y1              =   1320
         Y2              =   1080
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   15
         X1              =   2220
         X2              =   1980
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   14
         X1              =   1980
         X2              =   1980
         Y1              =   2340
         Y2              =   1080
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   13
         X1              =   1980
         X2              =   1980
         Y1              =   3060
         Y2              =   2640
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   11
         X1              =   2220
         X2              =   2220
         Y1              =   5160
         Y2              =   5520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   10
         X1              =   2220
         X2              =   180
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   5
         X1              =   3540
         X2              =   3000
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label V_fuel_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   138
         Top             =   4320
         Width           =   555
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Oil Control"
         Height          =   435
         Index           =   0
         Left            =   2340
         TabIndex        =   137
         Top             =   3900
         Width           =   915
      End
      Begin VB.Label V_fuel_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   136
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Oil Control"
         Height          =   435
         Index           =   9
         Left            =   2460
         TabIndex        =   135
         Top             =   360
         Width           =   915
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   6540
         Top             =   1020
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   6540
         Top             =   1800
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6540
         Top             =   4560
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   6540
         Top             =   5340
         Width           =   195
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":1850
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   115
         Top             =   1980
         Width           =   375
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":3868
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   114
         Top             =   1200
         Width           =   375
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":5880
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   96
         Top             =   1440
         Width           =   375
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":7898
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   97
         Top             =   660
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":98B0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   90
         Top             =   5520
         Width           =   375
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":B8C8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   125
         Top             =   1980
         Width           =   375
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":D8E0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   124
         Top             =   1200
         Width           =   375
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":F8F8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   123
         Top             =   1980
         Width           =   375
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":11910
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   122
         Top             =   1140
         Width           =   375
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":13928
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   121
         Top             =   1980
         Width           =   375
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":15940
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   120
         Top             =   1200
         Width           =   375
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":17958
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   119
         Top             =   1980
         Width           =   375
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":19970
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   118
         Top             =   1200
         Width           =   375
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":1B988
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   117
         Top             =   1980
         Width           =   375
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":1D9A0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   116
         Top             =   1200
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":1F9B8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   108
         Top             =   5040
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":219D0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   113
         Top             =   5040
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":239E8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   112
         Top             =   5040
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":25A00
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   111
         Top             =   5040
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":27A18
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   110
         Top             =   4980
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":29A30
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   109
         Top             =   5040
         Width           =   375
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":2BA48
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   107
         Top             =   660
         Width           =   375
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":2DA60
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   106
         Top             =   1440
         Width           =   375
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":2FA78
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   105
         Top             =   660
         Width           =   375
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":31A90
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   104
         Top             =   1440
         Width           =   375
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":33AA8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   103
         Top             =   660
         Width           =   375
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":35AC0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   102
         Top             =   1440
         Width           =   375
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":37AD8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   101
         Top             =   660
         Width           =   375
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":39AF0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   100
         Top             =   1440
         Width           =   375
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":3BB08
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   99
         Top             =   660
         Width           =   375
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":3DB20
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   98
         Top             =   1440
         Width           =   375
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":3FB38
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   84
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":41B50
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   78
         Top             =   4740
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":43B68
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   95
         Top             =   5520
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":45B80
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   94
         Top             =   5520
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":47B98
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   93
         Top             =   5520
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":49BB0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   92
         Top             =   5520
         Width           =   375
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":4BBC8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   91
         Top             =   5520
         Width           =   375
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":4DBE0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   89
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":4FBF8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   88
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":51C10
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   87
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":53C28
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   86
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":55C40
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   85
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":57C58
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   83
         Top             =   4740
         Width           =   375
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":59C70
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   82
         Top             =   4740
         Width           =   375
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":5BC88
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   81
         Top             =   4740
         Width           =   375
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":5DCA0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   80
         Top             =   4740
         Width           =   375
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   5940
         OleObjectBlob   =   "Burner FO_system.frx":5FCB8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   79
         Top             =   4740
         Width           =   375
      End
      Begin VB.OLE Valve9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   0
         Left            =   9780
         OleObjectBlob   =   "Burner FO_system.frx":61CD0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   74
         Top             =   5520
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE VV_M 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   3
         Left            =   9180
         OleObjectBlob   =   "Burner FO_system.frx":698E8
         SizeMode        =   1  'Stretch
         TabIndex        =   73
         Top             =   5460
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE VV_M 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   2
         Left            =   2520
         OleObjectBlob   =   "Burner FO_system.frx":71500
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUMG.vsd"
         TabIndex        =   72
         Top             =   5880
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE VV_M 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   1
         Left            =   2940
         OleObjectBlob   =   "Burner FO_system.frx":79118
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUMR.vsd"
         TabIndex        =   71
         Top             =   5460
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE VV_M 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   0
         Left            =   3060
         OleObjectBlob   =   "Burner FO_system.frx":80D30
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   70
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE Valve10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   0
         Left            =   9180
         OleObjectBlob   =   "Burner FO_system.frx":88948
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   66
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   7
         Left            =   9180
         OleObjectBlob   =   "Burner FO_system.frx":90560
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   61
         Top             =   2220
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   6
         Left            =   9780
         OleObjectBlob   =   "Burner FO_system.frx":98378
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   60
         Top             =   2340
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   5
         Left            =   9180
         OleObjectBlob   =   "Burner FO_system.frx":A0590
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   59
         Top             =   2820
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   4
         Left            =   9720
         OleObjectBlob   =   "Burner FO_system.frx":A8FA8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   58
         Top             =   3240
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   5040
         TabIndex        =   44
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   2640
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   3
         X1              =   360
         X2              =   2220
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Air Registers Port"
         Height          =   255
         Index           =   4
         Left            =   8940
         TabIndex        =   41
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Air Registers Stbd"
         Height          =   255
         Index           =   5
         Left            =   8940
         TabIndex        =   40
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Shape Ign1_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   4800
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   4800
         Shape           =   2  'Oval
         Top             =   4860
         Width           =   375
      End
      Begin VB.Shape Ign1_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   4800
         Shape           =   2  'Oval
         Top             =   5700
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   4800
         Shape           =   2  'Oval
         Top             =   1380
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   7140
         Shape           =   2  'Oval
         Top             =   1260
         Width           =   375
      End
      Begin VB.Shape Ign1_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   7140
         Shape           =   2  'Oval
         Top             =   1980
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   7140
         Shape           =   2  'Oval
         Top             =   4860
         Width           =   375
      End
      Begin VB.Shape Ign1_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   7140
         Shape           =   2  'Oval
         Top             =   5580
         Width           =   375
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   780
         Width           =   135
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   5100
         Width           =   135
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   4080
         Picture         =   "Burner FO_system.frx":B15C0
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   8
         X1              =   4110
         X2              =   4680
         Y1              =   3225
         Y2              =   3225
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   7
         X1              =   4680
         X2              =   4680
         Y1              =   4640
         Y2              =   1860
      End
      Begin VB.Label Label28 
         Height          =   135
         Left            =   6345
         TabIndex        =   39
         Top             =   5550
         Width           =   195
      End
      Begin VB.Label Label27 
         Height          =   135
         Left            =   6345
         TabIndex        =   38
         Top             =   5190
         Width           =   195
      End
      Begin VB.Label Label26 
         Height          =   135
         Left            =   6345
         TabIndex        =   37
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label Label25 
         Height          =   135
         Left            =   6345
         TabIndex        =   36
         Top             =   1980
         Width           =   195
      End
      Begin VB.Label Label24 
         Height          =   135
         Left            =   6345
         TabIndex        =   35
         Top             =   1620
         Width           =   195
      End
      Begin VB.Label Label23 
         Height          =   135
         Left            =   6345
         TabIndex        =   34
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label22 
         Height          =   135
         Left            =   6345
         TabIndex        =   33
         Top             =   4410
         Width           =   195
      End
      Begin VB.Label Label21 
         Height          =   135
         Left            =   6345
         TabIndex        =   32
         Top             =   4770
         Width           =   195
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   3
         X1              =   4680
         X2              =   3960
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   2
         X1              =   4680
         X2              =   4680
         Y1              =   1860
         Y2              =   1080
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   5940
         X2              =   4680
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FO re-circ"
         Height          =   465
         Left            =   3420
         TabIndex        =   31
         Top             =   3060
         Width           =   645
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   3540
         TabIndex        =   30
         Top             =   4380
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   3630
         TabIndex        =   29
         Top             =   870
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   4590
         Width           =   795
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   4590
         Y2              =   4470
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   4710
         Y2              =   4830
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   5940
         X2              =   8940
         Y1              =   6090
         Y2              =   6090
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   5370
         Width           =   795
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   5370
         Y2              =   5250
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   5490
         Y2              =   5610
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   6630
         X2              =   6630
         Y1              =   6090
         Y2              =   5880
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   8460
         X2              =   8460
         Y1              =   5850
         Y2              =   6090
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Boiler Furnace"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   3870
         Width           =   2235
      End
      Begin VB.Line Line12 
         X1              =   4200
         X2              =   4200
         Y1              =   5040
         Y2              =   5520
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Air"
         Height          =   195
         Left            =   6060
         TabIndex        =   25
         Top             =   3990
         Width           =   435
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner No 2"
         Height          =   195
         Left            =   6960
         TabIndex        =   24
         Top             =   4530
         Width           =   915
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner No 1"
         Height          =   195
         Left            =   6960
         TabIndex        =   23
         Top             =   5310
         Width           =   915
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   5940
         X2              =   5940
         Y1              =   4050
         Y2              =   6090
      End
      Begin VB.Label M9_label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   4020
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   540
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5940
         X2              =   5940
         Y1              =   480
         Y2              =   2520
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner No 1"
         Height          =   195
         Left            =   6960
         TabIndex        =   14
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner No 2"
         Height          =   195
         Left            =   6960
         TabIndex        =   13
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Air"
         Height          =   195
         Left            =   6060
         TabIndex        =   12
         Top             =   420
         Width           =   435
      End
      Begin VB.Line Line10 
         X1              =   4200
         X2              =   4200
         Y1              =   1440
         Y2              =   2160
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Port Boiler Furnace"
         Height          =   255
         Left            =   6720
         TabIndex        =   9
         Top             =   300
         Width           =   2235
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   5940
         X2              =   4680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   8460
         X2              =   8460
         Y1              =   2280
         Y2              =   2520
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   6600
         X2              =   6600
         Y1              =   2520
         Y2              =   2310
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   1920
         Y2              =   2040
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   1800
         Y2              =   1680
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   1800
         Width           =   795
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   5940
         X2              =   8940
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   1140
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   6720
         X2              =   6840
         Y1              =   1020
         Y2              =   900
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   1020
         Width           =   795
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   2
         X1              =   2220
         X2              =   2220
         Y1              =   1440
         Y2              =   5010
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flowmeter"
         Height          =   195
         Left            =   900
         TabIndex        =   3
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1080
         TabIndex        =   2
         Top             =   3120
         Width           =   105
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1695
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2415
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   1695
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   4170
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1020
         Shape           =   3  'Circle
         Top             =   3030
         Width           =   285
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   360
         X2              =   2220
         Y1              =   3210
         Y2              =   3210
      End
      Begin VB.OLE Valve9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   3
         Left            =   8100
         OleObjectBlob   =   "Burner FO_system.frx":B18CA
         SizeMode        =   1  'Stretch
         TabIndex        =   77
         Top             =   2580
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE Valve9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   2
         Left            =   8040
         OleObjectBlob   =   "Burner FO_system.frx":B94E2
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUMG.vsd"
         TabIndex        =   76
         Top             =   3180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE Valve9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   1
         Left            =   8700
         OleObjectBlob   =   "Burner FO_system.frx":C10FA
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUMR.vsd"
         TabIndex        =   75
         Top             =   3060
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   4
         X1              =   5940
         X2              =   4680
         Y1              =   4635
         Y2              =   4635
      End
      Begin VB.OLE Valve10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   1
         Left            =   8580
         OleObjectBlob   =   "Burner FO_system.frx":C8D12
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUMR.vsd"
         TabIndex        =   69
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE Valve10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   2
         Left            =   7920
         OleObjectBlob   =   "Burner FO_system.frx":D092A
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUMG.vsd"
         TabIndex        =   68
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE Valve10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   585
         Index           =   3
         Left            =   7260
         OleObjectBlob   =   "Burner FO_system.frx":D8542
         SizeMode        =   1  'Stretch
         TabIndex        =   67
         Top             =   6180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   8
         X1              =   5940
         X2              =   5700
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Burner_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   5
         X1              =   5400
         X2              =   4680
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   0
         Left            =   2580
         OleObjectBlob   =   "Burner FO_system.frx":E015A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   65
         Top             =   1080
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   1
         Left            =   2580
         OleObjectBlob   =   "Burner FO_system.frx":E8772
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   64
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   2580
         OleObjectBlob   =   "Burner FO_system.frx":F118A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   63
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   0
         Left            =   2520
         OleObjectBlob   =   "Burner FO_system.frx":F93A2
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   54
         Top             =   4620
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   2520
         OleObjectBlob   =   "Burner FO_system.frx":1019BA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   56
         Top             =   4620
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   1
         Left            =   2520
         OleObjectBlob   =   "Burner FO_system.frx":109BD2
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   55
         Top             =   4620
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   3
         Left            =   2520
         OleObjectBlob   =   "Burner FO_system.frx":1125EA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   57
         Top             =   4620
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   3
         Left            =   2580
         OleObjectBlob   =   "Burner FO_system.frx":11A402
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   6
         X1              =   2640
         X2              =   2220
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   9
         X1              =   2520
         X2              =   2220
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   4
         X1              =   3660
         X2              =   3060
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5820
         X2              =   5820
         Y1              =   3000
         Y2              =   1860
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5820
         X2              =   5820
         Y1              =   1080
         Y2              =   300
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5820
         X2              =   5820
         Y1              =   4620
         Y2              =   3480
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5820
         X2              =   5820
         Y1              =   6420
         Y2              =   5400
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1020
      TabIndex        =   161
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Burner_FO_system"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub BIGN1_arrow_Click(Index As Integer)


' === 21/01/03 ===
'        BIGN1_DI(1) = 387
'        BIGN2_DI(1) = 388
'        BIGN1_DI(2) = 404
'        BIGN2_DI(2) = 405
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If Index = 1 Then
   Call setPCDI(387)
ElseIf Index = 2 Then
   Call setPCDI(404)
End If


End Sub

Private Sub BIGN2_arrow_Click(Index As Integer)

' === 21/01/03 ===
'        BIGN1_DI(1) = 387
'        BIGN2_DI(1) = 388
'        BIGN1_DI(2) = 404
'        BIGN2_DI(2) = 405
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If Index = 1 Then
   Call setPCDI(388)
ElseIf Index = 2 Then
   Call setPCDI(405)
End If

End Sub

Private Sub Command1_Click()

FO_Service.Show

End Sub


Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idFO_system)




' 26/11/99
HScroll1(0).Visible = PTT
HScroll1(1).Visible = PTT
HScroll1(2).Visible = PTT
HScroll1(3).Visible = PTT

Reg_label(4).Visible = PTT
Reg_label(5).Visible = PTT


' 21/9/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
     V_fuel_pos(1).Visible = False
     V_fuel_pos(2).Visible = False
Else
     V_fuel_pos(1).Visible = True
     V_fuel_pos(2).Visible = True
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idFO_system)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub H_fuel_control_Change(Index As Integer)

' 30/7/99
If Index = 1 Then
    Call set_VSCROLL(5, H_fuel_control(1).Value) ' Stbd
ElseIf Index = 2 Then
    Call set_VSCROLL(6, H_fuel_control(2).Value) ' Port
End If

End Sub

Private Sub HScroll1_Change(Index As Integer)
Dim VV_position As Single


    Call set_VSCROLL(Index + 8, HScroll1(Index).Value)   ' Air register position

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub M10_label_Click()

' 21/9/99
'   Call valve_pcdi_LSS(219)
    Call valve_pcdi_LSS(210)

End Sub

Private Sub M9_label_Click()

' 21/9/99
'   Call valve_pcdi_LSS(217)
    Call valve_pcdi_LSS(209)

End Sub


Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

' === 12/02/03 SG DEFINATELY wants to be able to operate from ACT terminals... ===
' === 12/02/03 SO REMOVED... ===
' === 20/01/03 PUT IN:===
'If Not I_AM_INSTRUCTOR And Not PTT Then
'      ' I MUST be an ACT Terminal (Full Mission) so
'      Exit Sub
'End If
  
' === 19/08/02 ===

'   Note:  PumpIO_LSS(4=OFF,nPump)

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 41))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

' === 12/02/03 SG DEFINATELY wants to be able to operate from ACT terminals... ===
' === 12/02/03 SO REMOVED... ===
'' === 20/01/03 PUT IN:===
'If Not I_AM_INSTRUCTOR And Not PTT Then
'      ' I MUST be an ACT Terminal (Full Mission) so
'      Exit Sub
'End If
  
  
' === 19/08/02 ===

'   Note:  PumpIO_LSS(2=ON,nPump)

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 41))
End If


End Sub

Private Sub PFO_B_Click()

  '6/1/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 25
    nAlarm_LO = 33
    
    Alarm_set_point.Show

End Sub

Private Sub PFO_port_Click()
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 56
    nAlarm_LO = 64
    Alarm_set_point.Show

End Sub

Private Sub PFO_stbd_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 1
    nAlarm_LO = 9
    Alarm_set_point.Show
End Sub


Private Sub TFO_Port_Click()
Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 54
    nAlarm_LO = 62
    Alarm_set_point.Show
End Sub

Private Sub TFO_Stbd_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 3
    nAlarm_LO = 11
    Alarm_set_point.Show
    
End Sub

Private Sub Three_way_VV_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub

Select Case Index
    Case 1
        Call valve_pcdi_LSS(425)
    Case 2
        Call valve_pcdi_LSS(424)
End Select

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim VV_pos As Integer


' === 26/11/02 ===
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' === 21/01/03 ADDED: ===
If PTT Then
    If VB_steam_logicals(21) Then
    ' STBD Boiler is in CONSOLE mode...
        BIGN1_arrow(1).Visible = False
        BIGN2_arrow(1).Visible = False
    Else
    ' STBD Boiler is in MANUAL or BOILER FRONT mode...
        BIGN1_arrow(1).Visible = True
        BIGN2_arrow(1).Visible = True
    End If
    If VB_steam_logicals(24) Then
    ' PORT Boiler is in CONSOLE mode...
        BIGN1_arrow(2).Visible = False
        BIGN2_arrow(2).Visible = False
    Else
    ' PORT Boiler is in MANUAL or BOILER FRONT mode...
        BIGN1_arrow(2).Visible = True
        BIGN2_arrow(2).Visible = True
    End If
    
Else
        BIGN1_arrow(1).Visible = False
        BIGN1_arrow(2).Visible = False
        BIGN2_arrow(1).Visible = False
        BIGN2_arrow(2).Visible = False
End If

' === 21/01/03 end of edit ===


' 7/1/00
Call paint_alarm(Me.PFO_stbd, 1, 9, 0, 0)
Call paint_alarm(Me.PFO_port, 56, 64, 0, 0)

Call paint_alarm(Me.TFO_Stbd, 3, 11, 0, 0)
Call paint_alarm(Me.TFO_Port, 54, 62, 0, 0)

Call paint_alarm(Me.PFO_B, 25, 33, 0, 0)


' 12/5/00 Do not need twice
'16/9/99
'Call paint_alarm(Me.PFO_stbd, 1, 9, 0, 0)
'Call paint_alarm(Me.PFO_port, 56, 64, 0, 0)
'
'Call paint_alarm(Me.TFO_Stbd, 3, 11, 0, 0)
'Call paint_alarm(Me.TFO_Port, 54, 62, 0, 0)



' 26/11/99   HO/DO Nozzles
'equivalence(STEAM_logicals(15),HO_nozzle(1))

For i = 0 To 3

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


' 26/11/99      Burner FLAME
'        equivalence(STEAM_logicals(1),Burner1_flame(1))
'        equivalence(STEAM_logicals(3),Burner2_flame(1))
If VB_steam_logicals(1) Then
        Burner_shape(1).FillColor = &HFF&              ' RED
Else
        Burner_shape(1).FillColor = &HC0C0C0           ' GREY
End If
If VB_steam_logicals(3) Then
        Burner_shape(2).FillColor = &HFF&              ' RED
Else
        Burner_shape(2).FillColor = &HC0C0C0           ' GREY
End If

If VB_steam_logicals(2) Then
        Burner_shape(3).FillColor = &HFF&              ' RED
Else
        Burner_shape(3).FillColor = &HC0C0C0           ' GREY
End If
If VB_steam_logicals(4) Then
        Burner_shape(4).FillColor = &HFF&              ' RED
Else
        Burner_shape(4).FillColor = &HC0C0C0           ' GREY
End If


'               72,73           DI( BIGN1_DI(1-2) )
'               74,75           DI( BIGN2_DI(1-2) )

' 26/11/99   Igniter shapes
For i = 1 To 2

    If VB_steam_logicals(71 + i) Then
        Ign1_shape_in(i).Visible = True
        Ign1_shape_out(i).Visible = False
        
        ' === 21/01/03 ===
        BIGN1_arrow(i).Picture = Left_arrow_image.Picture

    Else
        Ign1_shape_in(i).Visible = False
        Ign1_shape_out(i).Visible = True
        
        ' === 21/01/03 ===
        BIGN1_arrow(i).Picture = Right_arrow_image.Picture
    End If
    
    If VB_steam_logicals(73 + i) Then
        Ign2_shape_in(i).Visible = True
        Ign2_shape_out(i).Visible = False
        
        ' === 21/01/03 ===
        BIGN2_arrow(i).Picture = Left_arrow_image.Picture
    Else
        Ign2_shape_in(i).Visible = False
        Ign2_shape_out(i).Visible = True
        
        ' === 21/01/03 ===
        BIGN2_arrow(i).Picture = Right_arrow_image.Picture
    End If
    
Next i


        
  


'Fuel_pos_stbd.text = Format(LSS_reals(193), "###0") & " %"
V_fuel_pos(1).Caption = Format(LSS_reals(193), "###0") & " %"
If LSS_reals(193) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(193) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(193) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Fuel_VV_stbd(i).Visible = True
    Else
        Fuel_VV_stbd(i).Visible = False
    End If
Next i

' 11/11/99
'Fuel_pos_port.text = Format(LSS_reals(194), "###0") & " %"
'V_fuel_pos(1).Caption = Format(LSS_reals(194), "###0") & " %"
V_fuel_pos(2).Caption = Format(LSS_reals(194), "###0") & " %"
If LSS_reals(194) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(194) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(194) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Fuel_VV_port(i).Visible = True
    Else
        Fuel_VV_port(i).Visible = False
    End If
Next i


TFO_Stbd.text = Format(LSS_reals(215) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TFO_Port.text = Format(LSS_reals(216) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


' === 21/08/02 ===
' PC_REALS(853) = P_DO_emerg_out
' PC_REALS(854) = F_DO_emerg




' === 22/08/02 ===

If SI_units Then

    PFO_A.text = Format(LSS_reals(230) * Press_mult, "##0.0") & Press_chars     ' Bar
    PFO_B.text = Format(LSS_reals(231) * Press_mult, "##0.0") & Press_chars     ' Bar
    PFO_port.text = Format(LSS_reals(228) * Press_mult, "##0.0") & Press_chars     ' Bar
    PFO_stbd.text = Format(LSS_reals(229) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    ' === 22/08/02 ===
    P_DO_emerg_out.text = Format(LSS_reals(853) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    F_DO_emerg.text = Format(LSS_reals(854) * Flow_mult, "###0.000") & Flow_chars
    
    F_flowmeter.text = Format(LSS_reals(226) * Flow_mult, "###0.000") & Flow_chars
    
    FFO_burner(1).text = Format(LSS_reals(232) * Flow_mult, "###0.000") & Flow_chars
    FFO_burner(2).text = Format(LSS_reals(233) * Flow_mult, "###0.000") & Flow_chars
    FFO_burner(3).text = Format(LSS_reals(190) * Flow_mult, "###0.000") & Flow_chars
    FFO_burner(4).text = Format(LSS_reals(191) * Flow_mult, "###0.000") & Flow_chars
    
   ' FFO_recirc(1).Text = Format(LSS_reals(234) * Flow_mult, "###0.000") & Flow_chars
   ' FFO_recirc(2).Text = Format(LSS_reals(192) * Flow_mult, "###0.000") & Flow_chars

Else
    PFO_A.text = Format(LSS_reals(230) * Press_mult, "##0") & Press_chars     ' Bar
    PFO_B.text = Format(LSS_reals(231) * Press_mult, "##0") & Press_chars     ' Bar
    PFO_port.text = Format(LSS_reals(228) * Press_mult, "##0") & Press_chars     ' Bar
    PFO_stbd.text = Format(LSS_reals(229) * Press_mult, "##0") & Press_chars     ' Bar
    
    ' === 22/08/02 ===
    P_DO_emerg_out.text = Format(LSS_reals(853) * Press_mult, "##0") & Press_chars     ' Bar
    F_DO_emerg.text = Format(LSS_reals(854) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
    
    F_flowmeter.text = Format(LSS_reals(226) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
    FFO_burner(1).text = Format(LSS_reals(232) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_burner(2).text = Format(LSS_reals(233) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_burner(3).text = Format(LSS_reals(190) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_burner(4).text = Format(LSS_reals(191) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
   ' FFO_recirc(1).Text = Format(LSS_reals(234) * Flow_mult / 3600#, "###0.000") & " lb/s"
   ' FFO_recirc(2).Text = Format(LSS_reals(192) * Flow_mult / 3600#, "###0.000") & " lb/s"

End If

' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)
'   21/9/99 OLD LSS Valves...
'   Call update_vvalve_LSS(Burner_FO_system, 206, 0, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 207, 1, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 211, 2, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 213, 3, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 210, 4, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 212, 5, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 208, 7, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 220, 8, -1)
'
'   Call update_vvalve_LSS(Burner_FO_system, 216, 11, -1)
'   Call update_vvalve_LSS(Burner_FO_system, 218, 12, -1)
'
'   Call update_vvalve_LSS(Burner_FO_system, 214, 15, -1)
'
'   Call update_vvalve_LSS(Burner_FO_system, 209, 16, -1)
'
'   Call update_vvalve_LSS(Burner_FO_system, 215, 17, -1)
   
For i = 0 To 5
   Call update_vvalve_LSS(Me, 200 + i, i, -1)
Next i
   Call update_vvalve_LSS(Me, 207, 7, -1)
   Call update_vvalve_LSS(Me, 208, 8, -1)

' 4/10/99
'For i = 11 To 12
For i = 9 To 12
   Call update_vvalve_LSS(Me, 200 + i, i, -1)
Next i
For i = 15 To 17
   Call update_vvalve_LSS(Me, 200 + i, i, -1)
Next i

' === 27/11/02 19.111 ===
If VB_More_integers(8) = 1 Then
    valve(18).OpenBtmRight = True
    valve(18).OpenTopLeft = True
Else
    valve(18).OpenBtmRight = False
    valve(18).OpenTopLeft = False
End If

   
   
' 21/9/99
'   Call update_motorised_VV(Valve9(0), Valve9(1), Valve9(2), Valve9(3), 217)
'   Call update_motorised_VV(Valve10(0), Valve10(1), Valve10(2), Valve10(3), 219)
    Call update_motorised_VV(Valve9(0), Valve9(1), Valve9(2), Valve9(3), 209)
    Call update_motorised_VV(Valve10(0), Valve10(1), Valve10(2), Valve10(3), 210)
    
    
' Air Registers   ( VV_pos 0 to 5 )  ,  LSS_reals(   ) 0 to 100.0

' Register1_stbd
Call Air_VV_position(219, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        Register1_stbd_top(i).Visible = True
        Register1_stbd_bot(i).Visible = True
    Else
        Register1_stbd_top(i).Visible = False
        Register1_stbd_bot(i).Visible = False
    End If
Next i

' Register2_stbd
Call Air_VV_position(439, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        register2_stbd_top(i).Visible = True
        Register2_stbd_bot(i).Visible = True
    Else
        register2_stbd_top(i).Visible = False
        Register2_stbd_bot(i).Visible = False
    End If
Next i

' Register1_port
Call Air_VV_position(440, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        Register1_port_top(i).Visible = True
        Register1_port_bot(i).Visible = True
    Else
        Register1_port_top(i).Visible = False
        Register1_port_bot(i).Visible = False
    End If
Next i

' Register2_port
Call Air_VV_position(441, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        Register2_port_top(i).Visible = True
        Register2_port_bot(i).Visible = True
    Else
        Register2_port_top(i).Visible = False
        Register2_port_bot(i).Visible = False
    End If
Next i


' === 19/08/02 ===
' Note: update_pump_LSS(syxxx,index,nPump)
Call update_pump_LSS(Me, 0, 41)      ' FO Supply Pump # 1

' === 19/08/02 See Bilge X-ref V10.xls...===
Call update_vvalve_LSS(Me, 421, 421, -1)
Call update_vvalve_LSS(Me, 422, 422, -1)
Call update_vvalve_LSS(Me, 423, 423, -1)

For i = 483 To 486
    Call update_vvalve_LSS(Me, i, i, -1)
Next i

' Port Boiler HO/DO
Call update_vvalve_LSS(Me, 424, 424, -1)
If vv_status_LSS(424) = 0 Then
    ' HO
    Three_way_VV(2).OpenBtmTri = False
    Three_way_VV(2).OpenTopTri = True
ElseIf vv_status_LSS(424) = 3 Then
    ' DO
    Three_way_VV(2).OpenBtmTri = True
    Three_way_VV(2).OpenTopTri = False
Else
    ' DO AND HO
    Three_way_VV(2).OpenBtmTri = True
    Three_way_VV(2).OpenTopTri = True
End If

' Stbd Boiler HO/DO
Call update_vvalve_LSS(Me, 425, 425, -1)
If vv_status_LSS(425) = 0 Then
    ' HO
    Three_way_VV(1).OpenBtmTri = True
    Three_way_VV(1).OpenTopTri = False
ElseIf vv_status_LSS(425) = 3 Then
    ' DO
    Three_way_VV(1).OpenBtmTri = False
    Three_way_VV(1).OpenTopTri = True
Else
    ' DO AND HO
    Three_way_VV(1).OpenBtmTri = True
    Three_way_VV(1).OpenTopTri = True
End If


End Sub

Private Sub valve_Click(Index As Integer)
   
   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
   
' 30/11/99  If Full Mission then do NOT allow clicks on MOTV and BOTV
If (Not PTT) Then
    If Index = 2 Or Index = 3 Or Index = 9 Or Index = 10 Or Index = 16 Or Index = 17 Then
        Exit Sub
    End If
End If

' See Fuel_Oil_Service_System_Valve Xref.xls

' === 19/08/02 ===
If (Index > 420) Then

    Call valve_pcdi_LSS(Index)
    
Else
    ' === 27/11/02 ===
    If Index = 18 Then
        ' Do now't
    Else
        Call valve_pcdi_LSS(Index + 200)
    End If
    
End If
   


End Sub
Sub update_motorised_VV(VV0 As OLE, VV1 As OLE, VV2 As OLE, VV3 As OLE, valve_number As Integer)

Dim single_char As String * 1

If valve_number < 251 Then

    single_char = Mid$(PC_INT_STRING, valve_number, 1)

    If single_char = "0" Then
        vv_status_LSS(valve_number) = 0
    ElseIf single_char = "1" Then
        vv_status_LSS(valve_number) = 1
    ElseIf single_char = "2" Then
        vv_status_LSS(valve_number) = 2
    Else
        vv_status_LSS(valve_number) = 3
    End If
    
Else
        ' 17/12/98  EXTRA valves added from 251 to 300
        vv_status_LSS(valve_number) = VB_VV_extras(valve_number - 250)

End If

If vv_status_LSS(valve_number) <> OLD_vv_status_LSS(valve_number) Then

    OLD_vv_status_LSS(valve_number) = vv_status_LSS(valve_number)
    
    If vv_status_LSS(valve_number) = 1 Or vv_status_LSS(valve_number) = 2 Then
        OLD_vv_status_LSS(valve_number) = -1     ' ALWAYS update when flashing
    End If
    
    If vv_status_LSS(valve_number) = 3 Then
        VV3.ZOrder
    ElseIf vv_status_LSS(valve_number) = 0 Then
       VV0.ZOrder
    Else
        If flash_flag Then
            VV1.ZOrder
        Else
            VV2.ZOrder
        End If
    End If
    
End If

End Sub

Sub Air_VV_position(VV_number As Integer, VV_posn As Integer)

If LSS_reals(VV_number) > 95# Then
    VV_posn = 5
ElseIf LSS_reals(VV_number) > 75# Then
    VV_posn = 4
ElseIf LSS_reals(VV_number) > 50# Then
    VV_posn = 3
ElseIf LSS_reals(VV_number) > 25# Then
    VV_posn = 2
ElseIf LSS_reals(VV_number) > 5# Then
    VV_posn = 1
Else
    VV_posn = 0
End If

End Sub

