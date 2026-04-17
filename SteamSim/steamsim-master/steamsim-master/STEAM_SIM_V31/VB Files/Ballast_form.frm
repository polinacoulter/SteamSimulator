VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Ballast_form 
   Caption         =   "Ballast"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   15165
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6525
      Left            =   0
      Top             =   0
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   11509
      Begin VB.CommandButton DIText_cmd 
         Caption         =   "DIText"
         Height          =   615
         Left            =   13620
         TabIndex        =   86
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "x10 Time"
         Height          =   555
         Left            =   13290
         TabIndex        =   81
         Top             =   3060
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Normal Time"
         Height          =   555
         Left            =   13290
         TabIndex        =   80
         Top             =   2520
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   6330
         TabIndex        =   79
         Top             =   1860
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   12780
         Top             =   4350
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   12750
         Top             =   3870
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
         Left            =   1740
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
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
         Left            =   1230
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   0
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
         Index           =   1
         Left            =   1200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   4140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   4410
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
         Left            =   1170
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1290
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1170
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   225
         Left            =   4320
         TabIndex        =   62
         Top             =   2790
         Width           =   255
      End
      Begin PumpPrj.Pump Pump1 
         Height          =   465
         Index           =   0
         Left            =   1140
         TabIndex        =   5
         Top             =   1905
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   820
         OrientLeft      =   0   'False
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   30
         Left            =   1830
         TabIndex        =   1
         Top             =   1410
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   1
         Left            =   1830
         TabIndex        =   2
         Top             =   1710
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   192
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   2
         Left            =   1830
         TabIndex        =   3
         Top             =   2010
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   3
         Left            =   1830
         TabIndex        =   4
         Top             =   2310
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   192
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   4
         Left            =   660
         TabIndex        =   7
         Top             =   1350
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   5
         Left            =   660
         TabIndex        =   8
         Top             =   1650
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   6
         Left            =   660
         TabIndex        =   9
         Top             =   1950
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   192
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   7
         Left            =   1830
         TabIndex        =   10
         Top             =   3120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   192
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   8
         Left            =   1830
         TabIndex        =   11
         Top             =   3420
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   9
         Left            =   1830
         TabIndex        =   12
         Top             =   3720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   192
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   10
         Left            =   1830
         TabIndex        =   13
         Top             =   4020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin PumpPrj.Pump Pump1 
         Height          =   465
         Index           =   1
         Left            =   1140
         TabIndex        =   14
         Top             =   3615
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   820
         OrientLeft      =   0   'False
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   11
         Left            =   660
         TabIndex        =   15
         Top             =   3060
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   192
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   12
         Left            =   660
         TabIndex        =   16
         Top             =   3360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   13
         Left            =   660
         TabIndex        =   17
         Top             =   3660
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   14
         Left            =   5430
         TabIndex        =   27
         Top             =   1155
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   15
         Left            =   7590
         TabIndex        =   28
         Top             =   1155
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   16
         Left            =   9840
         TabIndex        =   29
         Top             =   1155
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   19
         Left            =   7590
         TabIndex        =   33
         Top             =   4425
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   20
         Left            =   9840
         TabIndex        =   34
         Top             =   4425
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   18
         Left            =   5430
         TabIndex        =   39
         Top             =   4425
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   17
         Left            =   11200
         TabIndex        =   40
         Top             =   1890
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Index           =   21
         Left            =   11200
         TabIndex        =   41
         Top             =   3570
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   285
         Index           =   34
         Left            =   12810
         TabIndex        =   42
         Top             =   2880
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   192
         BottomColour    =   192
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   22
         Left            =   10740
         TabIndex        =   45
         Top             =   3300
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   23
         Left            =   2700
         TabIndex        =   46
         Top             =   2340
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   285
         Index           =   24
         Left            =   2520
         TabIndex        =   47
         Top             =   2820
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   31
         Left            =   1350
         TabIndex        =   48
         Top             =   2700
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   192
         BottomColour    =   192
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   25
         Left            =   2700
         TabIndex        =   49
         Top             =   3180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Index           =   32
         Left            =   750
         TabIndex        =   51
         Top             =   2460
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   192
         BottomColour    =   192
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   192
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   285
         Index           =   26
         Left            =   360
         TabIndex        =   52
         Top             =   5130
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
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
         Index           =   27
         Left            =   2700
         TabIndex        =   82
         Top             =   1410
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Index           =   28
         Left            =   2700
         TabIndex        =   83
         Top             =   4020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Index           =   29
         Left            =   3120
         TabIndex        =   84
         Top             =   2340
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
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
         Height          =   375
         Index           =   33
         Left            =   3120
         TabIndex        =   85
         Top             =   3180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1395
         Index           =   2
         Left            =   1770
         Top             =   3090
         Width           =   375
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1395
         Index           =   0
         Left            =   1770
         Top             =   1380
         Width           =   375
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   23
         X1              =   1230
         X2              =   2910
         Y1              =   5070
         Y2              =   5070
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   32
         X1              =   3480
         X2              =   3480
         Y1              =   4260
         Y2              =   3450
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   31
         X1              =   3480
         X2              =   3480
         Y1              =   2610
         Y2              =   1680
      End
      Begin VB.Label Z_wing_stbd 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   11730
         TabIndex        =   78
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label Z_wing_port 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   11730
         TabIndex        =   77
         Top             =   2220
         Width           =   645
      End
      Begin VB.Label Z_wing_stbd 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   8730
         TabIndex        =   76
         Top             =   5280
         Width           =   645
      End
      Begin VB.Label Z_wing_stbd 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   10770
         TabIndex        =   75
         Top             =   5280
         Width           =   645
      End
      Begin VB.Label Z_wing_stbd 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   6690
         TabIndex        =   74
         Top             =   5280
         Width           =   645
      End
      Begin VB.Label Z_wing_port 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   8730
         TabIndex        =   73
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Z_wing_port 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   10710
         TabIndex        =   72
         Top             =   900
         Width           =   645
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   1530
         X2              =   12840
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   1
         X1              =   10830
         X2              =   11550
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label20 
         Caption         =   "Ballast #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         TabIndex        =   70
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Ballast #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         TabIndex        =   69
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1440
         Picture         =   "Ballast_form.frx":0000
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   320
         Picture         =   "Ballast_form.frx":0442
         Stretch         =   -1  'True
         Top             =   4290
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   320
         Picture         =   "Ballast_form.frx":0884
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   2190
         Picture         =   "Ballast_form.frx":0CC6
         Stretch         =   -1  'True
         Top             =   4150
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   2190
         Picture         =   "Ballast_form.frx":1108
         Stretch         =   -1  'True
         Top             =   3850
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   2190
         Picture         =   "Ballast_form.frx":154A
         Stretch         =   -1  'True
         Top             =   1870
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   2190
         Picture         =   "Ballast_form.frx":198C
         Stretch         =   -1  'True
         Top             =   1550
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   1550
         Picture         =   "Ballast_form.frx":1DCE
         Stretch         =   -1  'True
         Top             =   2670
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   1550
         Picture         =   "Ballast_form.frx":2210
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   950
         Picture         =   "Ballast_form.frx":2652
         Stretch         =   -1  'True
         Top             =   2370
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   950
         Picture         =   "Ballast_form.frx":2A94
         Stretch         =   -1  'True
         Top             =   2790
         Width           =   255
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   900
         Picture         =   "Ballast_form.frx":2ED6
         Stretch         =   -1  'True
         Top             =   770
         Width           =   255
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   13140
         Picture         =   "Ballast_form.frx":3318
         Stretch         =   -1  'True
         Top             =   1710
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   13140
         Picture         =   "Ballast_form.frx":375A
         Stretch         =   -1  'True
         Top             =   1290
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   13170
         Picture         =   "Ballast_form.frx":3B9C
         Stretch         =   -1  'True
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   13170
         Picture         =   "Ballast_form.frx":3FDE
         Stretch         =   -1  'True
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Z_wing_port 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "100.0 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   61
         Top             =   900
         Width           =   645
      End
      Begin VB.Label ZHO_deep 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "100 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4470
         TabIndex        =   60
         Top             =   5280
         Width           =   525
      End
      Begin VB.Label ZDO_double_bottom 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   4170
         TabIndex        =   59
         Top             =   3780
         Width           =   555
      End
      Begin VB.Label ZDO_double_bottom 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4170
         TabIndex        =   58
         Top             =   2190
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ring Main"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   7950
         TabIndex        =   57
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pipe Tunnel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   5850
         TabIndex        =   56
         Top             =   3690
         Width           =   1215
      End
      Begin VB.Line Line12 
         BorderStyle     =   3  'Dot
         X1              =   10710
         X2              =   10020
         Y1              =   3480
         Y2              =   3750
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Isolating Valve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9150
         TabIndex        =   55
         Top             =   3720
         Width           =   1425
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Overboard Discharge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   54
         Top             =   5550
         Width           =   1035
      End
      Begin VB.Line Line11 
         BorderStyle     =   3  'Dot
         X1              =   2400
         X2              =   2970
         Y1              =   3060
         Y2              =   5700
      End
      Begin VB.Label Label12 
         Caption         =   "Sea Suction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2820
         TabIndex        =   53
         Top             =   5700
         Width           =   1455
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   30
         X1              =   2130
         X2              =   2490
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   29
         X1              =   2130
         X2              =   2490
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line10 
         X1              =   2280
         X2              =   2400
         Y1              =   2820
         Y2              =   3000
      End
      Begin VB.Line Line9 
         X1              =   2250
         X2              =   2400
         Y1              =   3000
         Y2              =   2820
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   6
         X1              =   2610
         X2              =   2610
         Y1              =   4290
         Y2              =   1680
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   28
         X1              =   2340
         X2              =   2340
         Y1              =   2760
         Y2              =   2400
      End
      Begin VB.Shape Shape32 
         Height          =   435
         Index           =   1
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2700
         Width           =   345
      End
      Begin VB.Shape Shape32 
         Height          =   255
         Index           =   0
         Left            =   2220
         Shape           =   3  'Circle
         Top             =   2790
         Width           =   225
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   14
         X1              =   1230
         X2              =   1290
         Y1              =   3090
         Y2              =   2970
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   27
         X1              =   1110
         X2              =   1230
         Y1              =   4950
         Y2              =   5070
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   26
         X1              =   1110
         X2              =   1110
         Y1              =   4080
         Y2              =   4950
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   25
         X1              =   1050
         X2              =   1110
         Y1              =   3930
         Y2              =   4080
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   24
         X1              =   960
         X2              =   1050
         Y1              =   3930
         Y2              =   3930
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   13
         X1              =   2910
         X2              =   3030
         Y1              =   5070
         Y2              =   4950
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   22
         X1              =   3030
         X2              =   3030
         Y1              =   4950
         Y2              =   4410
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   12
         X1              =   630
         X2              =   690
         Y1              =   2850
         Y2              =   2730
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   21
         X1              =   450
         X2              =   450
         Y1              =   5280
         Y2              =   2010
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   11
         X1              =   450
         X2              =   600
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   10
         X1              =   540
         X2              =   600
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   9
         X1              =   450
         X2              =   540
         Y1              =   2010
         Y2              =   1890
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   8
         X1              =   510
         X2              =   600
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   7
         X1              =   420
         X2              =   510
         Y1              =   1530
         Y2              =   1620
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   6
         X1              =   420
         X2              =   510
         Y1              =   1020
         Y2              =   900
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   20
         X1              =   420
         X2              =   420
         Y1              =   1530
         Y2              =   1020
      End
      Begin VB.Label ZHO_deep 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "100 %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4470
         TabIndex        =   50
         Top             =   660
         Width           =   525
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   5
         X1              =   2250
         X2              =   2340
         Y1              =   3690
         Y2              =   3570
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   4
         X1              =   2250
         X2              =   2340
         Y1              =   2310
         Y2              =   2400
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   19
         X1              =   2250
         X2              =   2130
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   18
         X1              =   2130
         X2              =   2250
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   15
         X1              =   1080
         X2              =   1080
         Y1              =   3330
         Y2              =   2250
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   17
         X1              =   960
         X2              =   1080
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   16
         X1              =   960
         X2              =   1050
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   14
         X1              =   690
         X2              =   1050
         Y1              =   2730
         Y2              =   2700
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   13
         X1              =   2340
         X2              =   2340
         Y1              =   3570
         Y2              =   3090
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   12
         X1              =   3480
         X2              =   2610
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   11
         X1              =   3480
         X2              =   2610
         Y1              =   2610
         Y2              =   2580
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   5
         X1              =   3030
         X2              =   3030
         Y1              =   4200
         Y2              =   1830
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   10
         X1              =   1680
         X2              =   1290
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   9
         X1              =   1800
         X2              =   1680
         Y1              =   3390
         Y2              =   3390
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   8
         X1              =   1680
         X2              =   1800
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   7
         X1              =   1680
         X2              =   1680
         Y1              =   3390
         Y2              =   2580
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   3
         X1              =   2940
         X2              =   3030
         Y1              =   900
         Y2              =   1020
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   4
         X1              =   3030
         X2              =   3030
         Y1              =   1560
         Y2              =   1020
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   3
         X1              =   510
         X2              =   2940
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Shape Shape31 
         Height          =   915
         Left            =   3570
         Top             =   2490
         Width           =   7095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pipe Tunnel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   5850
         TabIndex        =   44
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fore Peak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   6
         Left            =   12630
         TabIndex        =   43
         Top             =   3210
         Width           =   735
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         Index           =   2
         X1              =   12900
         X2              =   12900
         Y1              =   2760
         Y2              =   3150
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   2
         X1              =   12810
         X2              =   12900
         Y1              =   2640
         Y2              =   2760
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   2
         X1              =   10830
         X2              =   12810
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         Index           =   1
         X1              =   11610
         X2              =   11610
         Y1              =   3960
         Y2              =   4350
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   1
         X1              =   11550
         X2              =   11610
         Y1              =   3840
         Y2              =   3960
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         Index           =   0
         X1              =   11610
         X2              =   11610
         Y1              =   1650
         Y2              =   2040
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         Index           =   0
         X1              =   11550
         X2              =   11610
         Y1              =   2160
         Y2              =   2040
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         Index           =   0
         X1              =   10830
         X2              =   11520
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Shape Shape24 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   60
         Top             =   4980
         Width           =   2295
      End
      Begin VB.Shape Shape23 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1080
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   2
         Height          =   555
         Left            =   240
         Shape           =   2  'Oval
         Top             =   4995
         Width           =   2115
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   1500
         X2              =   12840
         Y1              =   5520
         Y2              =   5535
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   5
         X1              =   5520
         X2              =   5520
         Y1              =   4290
         Y2              =   4650
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   13
         X1              =   11520
         X2              =   12540
         Y1              =   2910
         Y2              =   5490
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   12
         X1              =   11490
         X2              =   12540
         Y1              =   5520
         Y2              =   2880
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 1 Stbd DB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   5
         Left            =   11700
         TabIndex        =   38
         Top             =   2970
         Width           =   615
      End
      Begin VB.Shape Shape21 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   2655
         Left            =   11460
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 3 Stbd Wing Ballast Tank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5820
         TabIndex        =   37
         Top             =   4740
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 2 Stbd Wing Ballast Tank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   7770
         TabIndex        =   36
         Top             =   4740
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 1 Stbd Wing Ballast Tank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   9810
         TabIndex        =   35
         Top             =   4740
         Width           =   1275
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   4
         X1              =   7680
         X2              =   7680
         Y1              =   4290
         Y2              =   4650
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   3
         X1              =   9930
         X2              =   9930
         Y1              =   4290
         Y2              =   4650
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   11
         X1              =   11520
         X2              =   12540
         Y1              =   300
         Y2              =   2910
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   10
         X1              =   11520
         X2              =   12540
         Y1              =   2880
         Y2              =   300
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 1 Port DB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   11730
         TabIndex        =   32
         Top             =   390
         Width           =   615
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   2625
         Left            =   11460
         Top             =   300
         Width           =   1095
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   8880
         X2              =   10650
         Y1              =   3390
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   8
         X1              =   10650
         X2              =   8880
         Y1              =   3390
         Y2              =   2490
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Void"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   9270
         TabIndex        =   31
         Top             =   2580
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   7110
         X2              =   8880
         Y1              =   3420
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   8880
         X2              =   7110
         Y1              =   3420
         Y2              =   2490
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Void"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7470
         TabIndex        =   30
         Top             =   2580
         Width           =   975
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   2
         X1              =   9930
         X2              =   9930
         Y1              =   1200
         Y2              =   1680
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   1
         X1              =   7680
         X2              =   7680
         Y1              =   1200
         Y2              =   1680
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   0
         X1              =   5520
         X2              =   5520
         Y1              =   1230
         Y2              =   1680
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   5340
         X2              =   3600
         Y1              =   5550
         Y2              =   2970
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   3570
         X2              =   5340
         Y1              =   300
         Y2              =   2970
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   7110
         X2              =   5340
         Y1              =   3420
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   5340
         X2              =   7110
         Y1              =   3420
         Y2              =   2520
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Void"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5760
         TabIndex        =   26
         Top             =   2580
         Width           =   975
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   915
         Index           =   0
         Left            =   5340
         Top             =   2490
         Width           =   1785
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 1 Port Wing Ballast Tank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   9810
         TabIndex        =   25
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 2 Port Wing Ballast Tank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   7830
         TabIndex        =   24
         Top             =   390
         Width           =   1275
      End
      Begin VB.Shape Shape15 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   9720
         Top             =   300
         Width           =   1755
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7530
         Top             =   300
         Width           =   2205
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   3570
         X2              =   5310
         Y1              =   5520
         Y2              =   2970
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   3570
         X2              =   5340
         Y1              =   2970
         Y2              =   300
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "HO Tank Stbd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3780
         TabIndex        =   23
         Top             =   5010
         Width           =   945
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3570
         Top             =   4950
         Width           =   1785
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFE&
         Caption         =   "Coffer Dam"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         TabIndex        =   22
         Top             =   4710
         Width           =   1335
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H00FFFFFE&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   3570
         Top             =   4680
         Width           =   1785
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DO DB  Stbd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   3540
         Width           =   1425
      End
      Begin VB.Shape Shape11 
         FillColor       =   &H000080FF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   1785
         Left            =   3570
         Top             =   2910
         Width           =   1785
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DO DB  Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   20
         Top             =   1950
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No 3 Port Wing Ballast Tank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5790
         TabIndex        =   19
         Top             =   330
         Width           =   1275
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   5340
         Top             =   300
         Width           =   2205
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   2130
         X2              =   3510
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         X1              =   2130
         X2              =   3540
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFE&
         Caption         =   "Coffer Dam"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         TabIndex        =   18
         Top             =   930
         Width           =   1335
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFE&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3570
         Top             =   900
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   5
         X1              =   960
         X2              =   1170
         Y1              =   3630
         Y2              =   3630
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   3
         Left            =   600
         Top             =   3030
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   4
         X1              =   1260
         X2              =   1770
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   960
         X2              =   1170
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   1
         Left            =   600
         Top             =   1320
         Width           =   375
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1080
         Top             =   315
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   120
         Top             =   540
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   555
         Left            =   240
         Shape           =   2  'Oval
         Top             =   285
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "HO Tank Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3780
         TabIndex        =   6
         Top             =   390
         Width           =   945
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3570
         Top             =   300
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   1260
         X2              =   1770
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.OLE OLE1 
         Class           =   "Visio.Drawing.5"
         Height          =   5025
         Left            =   4320
         OleObjectBlob   =   "Ballast_form.frx":4420
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Images\Ballast_System.vsd"
         TabIndex        =   0
         Top             =   6270
         Visible         =   0   'False
         Width           =   12495
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   3375
         Left            =   2910
         Top             =   1650
         Width           =   495
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H000080FF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   1785
         Left            =   3570
         Top             =   1140
         Width           =   1785
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   915
         Index           =   1
         Left            =   7110
         Top             =   2490
         Width           =   1785
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   915
         Index           =   2
         Left            =   8880
         Top             =   2490
         Width           =   1785
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   2625
         Left            =   3030
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   7815
      End
      Begin VB.Shape Shape18 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   9720
         Top             =   4710
         Width           =   1755
      End
      Begin VB.Shape Shape19 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7530
         Top             =   4710
         Width           =   2205
      End
      Begin VB.Shape Shape20 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   5340
         Top             =   4710
         Width           =   2205
      End
      Begin VB.Shape Shape26 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   11610
         Top             =   315
         Width           =   1335
      End
      Begin VB.Shape Shape27 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   12030
         Top             =   540
         Width           =   2295
      End
      Begin VB.Shape Shape25 
         BorderWidth     =   2
         Height          =   555
         Left            =   11670
         Shape           =   2  'Oval
         Top             =   285
         Width           =   2115
      End
      Begin VB.Shape Shape29 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   11610
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Shape Shape30 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   12030
         Top             =   4980
         Width           =   2295
      End
      Begin VB.Shape Shape28 
         BorderWidth     =   2
         Height          =   555
         Left            =   11670
         Shape           =   2  'Oval
         Top             =   4995
         Width           =   2115
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4980
      TabIndex        =   71
      Top             =   660
      Width           =   435
   End
End
Attribute VB_Name = "Ballast_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()

' === 27/05/02 ===
Bunkers.Show


End Sub

Private Sub Command2_Click()

' === 29/05/02 ===
Timer1.Interval = 100


End Sub

Private Sub Command3_Click()

' === 29/05/02 ===
Timer1.Interval = 10

End Sub

Private Sub Command4_Click()


End Sub

Private Sub DIText_cmd_Click()

Call IO_stuff

End Sub

Private Sub Form_Load()

Dim i As Integer

' === 28/05/02 REMOVE for TESTING ===
'On Error Resume Next
'Call JYAScreenInit(Me, idBallast)
'

' ADD for TESTING
Me.Width = HavenStretch1.Width
Me.Height = HavenStretch1.Height


set_dummy_focus = True
'
'HavenStretch1.Back_color = Bcolour
'Call copy_havenstretch_colours(Me, Bcolour)

Call Ballast_Init


End Sub

Private Sub Form_Resize()

Dim i As Integer

' === 28/05/02 REMOVE for TESTING ===
'On Error Resume Next
'If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
'    HavenStretch1.Visible = False
'    Call JYAScreenResize(Me, idBallast)
'    HavenStretch1.Visible = True
'End If



End Sub

Private Sub OFFpump_Click(Index As Integer)

ONpump(Index).BackColor = &HC0C0C0      ' GREY
OFFpump(Index).BackColor = vbWhite

B_pump_ON(Index + 1) = False

Pump1(Index).PumpOn = False

set_dummy_focus = True

End Sub

Private Sub ONpump_Click(Index As Integer)

ONpump(Index).BackColor = vbGreen
OFFpump(Index).BackColor = &HC0C0C0     ' GREY

B_pump_ON(Index + 1) = True

Pump1(Index).PumpOn = True

set_dummy_focus = True

End Sub

Private Sub Timer1_Timer()


' === 28/05/02 REMOVE for TESTING ===
'
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If
'
'
'' HO Deep tanks
'ZHO_deep(2).Caption = Format(LSS_reals(236), "##0.0") & " %"
'ZHO_deep(1).Caption = Format(LSS_reals(235), "##0.0") & " %"
'
'' DO Double Bottom Tanks
'ZDO_double_bottom(2).Caption = Format(LSS_reals(241), "##0.0") & " %"
'ZDO_double_bottom(1).Caption = Format(LSS_reals(240), "##0.0") & " %"

Call Ballast



End Sub

Private Sub Timer2_Timer()
Dim i As Integer
Dim Display_value As Single


' Ballast Valves

For i = 1 To 33

    If Open_ballast_valve(i) Then
        If T_ballast_valve(i) < 5# Then
            T_ballast_valve(i) = T_ballast_valve(i) + 0.5
                Valve(i).OpenBtmRight = Not Valve(i).OpenBtmRight
                Valve(i).OpenTopLeft = Valve(i).OpenBtmRight
       Else
                Valve(i).OpenBtmRight = True
                Valve(i).OpenTopLeft = True
       End If
    Else
        If T_ballast_valve(i) > 0# Then
            T_ballast_valve(i) = T_ballast_valve(i) - 0.5
                Valve(i).OpenBtmRight = Not Valve(i).OpenBtmRight
                Valve(i).OpenTopLeft = Valve(i).OpenBtmRight
        Else
                Valve(i).OpenBtmRight = False
                Valve(i).OpenTopLeft = False
        End If
    End If

    If T_ballast_valve(i) > 5# Then
        T_ballast_valve(i) = 5#
    ElseIf T_ballast_valve(i) < 0# Then
        T_ballast_valve(i) = 5#
    End If
 
Next i

    List1.Clear
    List1.AddItem "Pump # 1   " & Format(P_ballast(37), "##0.0")
    List1.AddItem "Pump # 2   " & Format(P_ballast(38), "##0.0")
    List1.AddItem "P_ballast(1)   " & Format(P_ballast(1), "##0.000")
    List1.AddItem "P_ballast(2)   " & Format(P_ballast(2), "##0.000")
    List1.AddItem "P_ballast(3)   " & Format(P_ballast(3), "##0.000")
    List1.AddItem "P_ballast(4)   " & Format(P_ballast(4), "##0.000")
    List1.AddItem "P_ballast(17)   " & Format(P_ballast(17), "##0.000")
    List1.AddItem " "
    List1.AddItem "F_ballast(2)   " & Format(F_ballast(2), "##0.000")
    List1.AddItem "F_ballast(30)   " & Format(F_ballast(30), "##0.000")
    List1.AddItem "F_ballast(39)   " & Format(F_ballast(39), "##0.000")


' Displays
For i = 1 To 4
    If ZWP(i) > 1# Then    ' WE are filling the Expansion pipe
        Display_value = 100#
    Else
        Display_value = ZWP(i) * 100#
    End If
        Z_wing_port(i).Caption = Format(Display_value, "##0.0") & " %"
    If ZWS(i) > 1# Then    ' WE are filling the Expansion pipe
        Display_value = 100#
    Else
        Display_value = ZWS(i) * 100#
    End If
        Z_wing_stbd(i).Caption = Format(Display_value, "##0.0") & " %"
Next i


End Sub

Private Sub Valve_Click(Index As Integer)
Dim skip_valve As Boolean

skip_valve = False

Select Case Index
    Case 1
        skip_valve = True
    Case 3
        skip_valve = True
    Case 7
        skip_valve = True
    Case 9
        skip_valve = True
    Case 6
        skip_valve = True
    Case 11
        skip_valve = True

    Case 31
        skip_valve = True
    Case 32
        skip_valve = True
    Case 34
        skip_valve = True
    
End Select


If Not skip_valve Then
        Open_ballast_valve(Index) = Not Open_ballast_valve(Index)
End If

End Sub

Sub IO_stuff()
Dim fp As Integer
Dim fp_out As Integer

Dim InfoString As String * 100
Dim info As String

Dim i As Integer

On Error GoTo Errhandler

fp = FreeFile(0)

'If IOType = "AI" Then
'   Open "C:\Steam_Sim\Text\AIText.dat" For Random Access Read As #fp Len = 100
'ElseIf IOType = "AO" Then
'   Open "C:\Steam_Sim\Text\AOText.dat" For Random Access Read As #fp Len = 100
'ElseIf IOType = "DI" Then
   
   Open "C:\Steam_Sim\Text\DIText.dat" For Random Access Read As #fp Len = 100

'ElseIf IOType = "DO" Then
'   Open "C:\Steam_Sim\Text\DOText.dat" For Random Access Read As #fp Len = 100
'Else
'    GoTo Errhandler
'End If


fp_out = FreeFile(0)

Open "C:\Steam_Sim\Text\DIText_new.dat" For Output As #fp_out


i = 0

Do While Not EOF(fp)


    i = i + 1
    
   'Get #fp, IONum, InfoString
    Get #fp, i, InfoString
    IOText = Mid$(InfoString, 1, 100)
    
    Print #fp_out, i, InfoString
     
Loop

Close #fp
Close #fp_out

Exit Sub

Errhandler:
IOText = "Error Reading Text for " & IOType & "(" & IONum & ")"

'12/12/00 'can't reproduce it
Close #fp


End Sub


