VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{A6D9C597-48C2-11D1-8915-000021240899}#2.0#0"; "Haven3WayValve v1.41.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Distil 
   Caption         =   "Distiller"
   ClientHeight    =   8700
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11205
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8235
      Left            =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   14526
      Begin VB.CommandButton Reset_command 
         Caption         =   "Reset"
         Height          =   315
         Left            =   7860
         TabIndex        =   73
         Top             =   3600
         Width           =   1035
      End
      Begin VB.TextBox FWG_vac 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   420
         MousePointer    =   1  'Arrow
         TabIndex        =   71
         Top             =   2460
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salinometer"
         Height          =   375
         Left            =   7860
         TabIndex        =   70
         Top             =   3180
         Width           =   1035
      End
      Begin VB.TextBox F_FWG_to_BILGE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Top             =   5280
         Width           =   855
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   21
         Left            =   9120
         TabIndex        =   40
         Top             =   960
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
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
         Left            =   5460
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
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
         Left            =   7380
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   7740
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox F_shore_tank_out 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6900
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox F_shore_tank_in 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox F_FWG_DWT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6540
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox F_shore_tank_FWG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4740
         MousePointer    =   1  'Arrow
         TabIndex        =   62
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox Z_shore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Top             =   1200
         Width           =   675
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
         Left            =   4860
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   4320
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
         Left            =   6720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   7740
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
         Left            =   4380
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   4320
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
         Left            =   6240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   7740
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox FWG_salinity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6780
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   3180
         Width           =   855
      End
      Begin Haven3WayValvePrj.Haven3WayValve Haven3WayValve1 
         Height          =   495
         Left            =   6000
         TabIndex        =   11
         Top             =   3540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OpenBtmTri      =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         Direction       =   3
      End
      Begin VB.CommandButton Command3D2 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   2640
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4920
         Top             =   7200
      End
      Begin VB.TextBox Z_DWT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10140
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Top             =   2700
         Width           =   675
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   0
         Left            =   7440
         TabIndex        =   12
         Top             =   6795
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   1
         Left            =   8505
         TabIndex        =   13
         Top             =   6795
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   2
         Left            =   6240
         TabIndex        =   14
         Top             =   6540
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   3
         Left            =   3480
         TabIndex        =   15
         Top             =   4320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   4
         Left            =   3480
         TabIndex        =   16
         Top             =   4800
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   5
         Left            =   4320
         TabIndex        =   17
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   6
         Left            =   4320
         TabIndex        =   18
         Top             =   2160
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   7
         Left            =   4200
         TabIndex        =   19
         Top             =   3660
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   8
         Left            =   7440
         TabIndex        =   20
         Top             =   2040
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
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   4680
         TabIndex        =   21
         Top             =   3780
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   16711680
         CentrLineWidth  =   3
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   6780
         TabIndex        =   22
         Top             =   6870
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   32768
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 DWT_tank 
         Height          =   735
         Left            =   7860
         TabIndex        =   23
         Top             =   2280
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1296
         TopColour       =   8388608
         BottomColour    =   8388608
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   8388608
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   11
         Left            =   1620
         TabIndex        =   25
         Top             =   -100
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Shore_tank 
         Height          =   855
         Left            =   1140
         TabIndex        =   27
         Top             =   660
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1508
         TopColour       =   8388608
         BottomColour    =   8388608
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   8388608
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   9
         Left            =   5220
         TabIndex        =   30
         Top             =   3420
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   10
         Left            =   6540
         TabIndex        =   31
         Top             =   3900
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   12
         Left            =   3720
         TabIndex        =   33
         Top             =   540
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
         Height          =   495
         Index           =   13
         Left            =   3720
         TabIndex        =   36
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
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
         Index           =   19
         Left            =   10260
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   49152
         BackColour      =   12632256
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   18
         Left            =   9780
         TabIndex        =   38
         Top             =   660
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
         Index           =   20
         Left            =   9780
         TabIndex        =   39
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
         Height          =   375
         Index           =   15
         Left            =   9780
         TabIndex        =   46
         Top             =   3540
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
         Index           =   16
         Left            =   9780
         TabIndex        =   47
         Top             =   4200
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
         Index           =   17
         Left            =   9120
         TabIndex        =   48
         Top             =   3780
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   105
         Left            =   420
         TabIndex        =   72
         Top             =   1440
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
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1620
         Picture         =   "Distiller.frx":0000
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   255
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3300
         X2              =   2880
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3300
         X2              =   3300
         Y1              =   1740
         Y2              =   1980
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1620
         Picture         =   "Distiller.frx":030A
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   255
      End
      Begin VB.Line Line48 
         X1              =   1200
         X2              =   1440
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   9900
         X2              =   9900
         Y1              =   2940
         Y2              =   3600
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   9900
         X2              =   9900
         Y1              =   1680
         Y2              =   2280
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   9900
         X2              =   9900
         Y1              =   420
         Y2              =   1020
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   30
         X1              =   9240
         X2              =   9240
         Y1              =   600
         Y2              =   1860
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   600
         Picture         =   "Distiller.frx":0614
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5040
         Picture         =   "Distiller.frx":091E
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   255
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00000000&
         X1              =   720
         X2              =   1440
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dist. Pump"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4440
         MousePointer    =   1  'Arrow
         TabIndex        =   57
         Top             =   3480
         Width           =   765
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6840
         Picture         =   "Distiller.frx":0C28
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   255
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   7860
         X2              =   6240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label DWT_dump_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10140
         TabIndex        =   55
         Top             =   3780
         Width           =   555
      End
      Begin VB.Label DFT_dump_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10140
         TabIndex        =   54
         Top             =   900
         Width           =   555
      End
      Begin VB.Label DWT_dump_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9240
         TabIndex        =   53
         Top             =   3540
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   9240
         X2              =   9240
         Y1              =   3420
         Y2              =   4680
      End
      Begin VB.Label DFT_dump_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9240
         TabIndex        =   45
         Top             =   720
         Width           =   555
      End
      Begin VB.Image Image12 
         Height          =   360
         Left            =   8880
         Picture         =   "Distiller.frx":0F32
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   285
      End
      Begin VB.Image Image18 
         Height          =   360
         Left            =   2640
         Picture         =   "Distiller.frx":123C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   285
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6360
         X2              =   8040
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   600
         X2              =   2280
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Image Image21 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6120
         Picture         =   "Distiller.frx":1546
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "DFT Dump Regulator"
         Height          =   375
         Left            =   9000
         TabIndex        =   35
         Top             =   120
         Width           =   795
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9900
         X2              =   9240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9900
         X2              =   9240
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Heating Steam"
         Height          =   255
         Left            =   4140
         TabIndex        =   34
         Top             =   4800
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Make up Feed Water to ADT"
         Height          =   735
         Left            =   9300
         TabIndex        =   32
         Top             =   5040
         Width           =   1155
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9900
         X2              =   9240
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9900
         X2              =   9240
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6240
         X2              =   6240
         Y1              =   900
         Y2              =   3540
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Goggle plate"
         Height          =   315
         Left            =   4920
         TabIndex        =   29
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Drain to bilge"
         Height          =   255
         Left            =   6300
         TabIndex        =   28
         Tag             =   "Drain to bilges"
         Top             =   5940
         Width           =   1035
      End
      Begin VB.Image Image30 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   840
         Picture         =   "Distiller.frx":1850
         Stretch         =   -1  'True
         Top             =   6780
         Width           =   255
      End
      Begin VB.Line Line32 
         BorderWidth     =   2
         X1              =   6240
         X2              =   6240
         Y1              =   6060
         Y2              =   5760
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   6240
         X2              =   6360
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   6120
         X2              =   6240
         Y1              =   5640
         Y2              =   5760
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   2280
         X2              =   2280
         Y1              =   660
         Y2              =   240
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   4740
         X2              =   4740
         Y1              =   1260
         Y2              =   1860
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   4800
         X2              =   4800
         Y1              =   1260
         Y2              =   1500
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   4680
         X2              =   4680
         Y1              =   1260
         Y2              =   1500
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " S "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5640
         TabIndex        =   6
         Top             =   3720
         Width           =   240
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   5100
         X2              =   6180
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shore Supply"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         MousePointer    =   1  'Arrow
         TabIndex        =   26
         Top             =   360
         Width           =   945
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   8040
         X2              =   8040
         Y1              =   1320
         Y2              =   2280
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6120
         X2              =   1560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8040
         Picture         =   "Distiller.frx":1B5A
         Stretch         =   -1  'True
         Top             =   7005
         Width           =   255
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   7020
         X2              =   8640
         Y1              =   7125
         Y2              =   7125
      End
      Begin VB.Image Image9 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1560
         Picture         =   "Distiller.frx":1E64
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   4200
         Y2              =   6780
      End
      Begin VB.Image Image11 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2160
         Picture         =   "Distiller.frx":216E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image28 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4800
         Picture         =   "Distiller.frx":2478
         Stretch         =   -1  'True
         Top             =   6780
         Width           =   255
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1320
         Picture         =   "Distiller.frx":2782
         Stretch         =   -1  'True
         Top             =   6690
         Width           =   495
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   720
         X2              =   6960
         Y1              =   6900
         Y2              =   6900
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00000000&
         Visible         =   0   'False
         X1              =   9600
         X2              =   10320
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3540
         X2              =   6240
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8880
         Top             =   6900
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEMISTER PLATE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1800
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4200
         X2              =   1860
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4200
         X2              =   1860
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4200
         X2              =   1860
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4200
         X2              =   1860
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1860
         Top             =   2340
         Width           =   495
      End
      Begin VB.Image Image27 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5160
         Picture         =   "Distiller.frx":2A8C
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   255
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4800
         Picture         =   "Distiller.frx":2D96
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image Image32 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4800
         Picture         =   "Distiller.frx":30A0
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   255
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5160
         Picture         =   "Distiller.frx":33AA
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   255
      End
      Begin VB.Image Image23 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6120
         Picture         =   "Distiller.frx":36B4
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   255
      End
      Begin VB.Image Image29 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3960
         Picture         =   "Distiller.frx":39BE
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   255
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   3480
         X2              =   4080
         Y1              =   5580
         Y2              =   5580
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4680
         X2              =   5280
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4680
         X2              =   5280
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5820
         X2              =   3840
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5760
         X2              =   3600
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Distilled Water tank"
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   8160
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   1680
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SALINOMETER"
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
         Left            =   6600
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   3600
         Width           =   1020
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   3240
         Top             =   4200
         Width           =   255
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   4080
         X2              =   4080
         Y1              =   5580
         Y2              =   6900
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   4200
         Top             =   2280
         Width           =   135
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00008000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1440
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00000000&
         X1              =   3840
         X2              =   3840
         Y1              =   3960
         Y2              =   4200
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00000000&
         X1              =   1440
         X2              =   1440
         Y1              =   3960
         Y2              =   4200
      End
      Begin VB.Line Line47 
         BorderColor     =   &H00000000&
         X1              =   3840
         X2              =   3480
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Image Image26 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3000
         Picture         =   "Distiller.frx":3CC8
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   255
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         Picture         =   "Distiller.frx":3FD2
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   255
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2520
         Picture         =   "Distiller.frx":42DC
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   255
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3720
         Top             =   3240
         Width           =   615
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3720
         Top             =   3120
         Width           =   615
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   4080
         X2              =   4080
         Y1              =   4020
         Y2              =   3600
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   5100
         X2              =   4080
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6240
         X2              =   6240
         Y1              =   5580
         Y2              =   3960
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00000000&
         X1              =   9120
         X2              =   9120
         Y1              =   6660
         Y2              =   7620
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   6720
         X2              =   5760
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   5760
         X2              =   5760
         Y1              =   3360
         Y2              =   3720
      End
      Begin VB.Shape Shape21 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   6600
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   7140
         X2              =   6120
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   7140
         X2              =   7140
         Y1              =   3600
         Y2              =   3840
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Orifice"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   5880
         Width           =   450
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00000000&
         X1              =   3780
         X2              =   3780
         Y1              =   5385
         Y2              =   5550
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00000000&
         X1              =   3780
         X2              =   3780
         Y1              =   5640
         Y2              =   5790
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   3855
         X2              =   3855
         Y1              =   5400
         Y2              =   5775
      End
      Begin VB.Line Line44 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   3720
         X2              =   3720
         Y1              =   5400
         Y2              =   5775
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SW Feed Pump"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6660
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   7500
         Width           =   1365
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O/B"
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   6540
         Width           =   255
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Brine Ejector"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1860
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   6600
         Width           =   900
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00000000&
         X1              =   1440
         X2              =   1800
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Air Ejector"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   2040
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   1800
         Top             =   4200
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   2160
         Top             =   4200
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   2520
         Top             =   4200
         Width           =   255
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   2880
         Top             =   4200
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   1440
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   375
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00008000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   1800
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1800
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1560
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   615
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1440
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2820
         X2              =   2820
         Y1              =   2140
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3300
         X2              =   300
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   9900
         X2              =   9900
         Y1              =   4500
         Y2              =   4920
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":45E6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   41
         Top             =   1020
         Width           =   405
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":C5FE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   44
         Top             =   1020
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":14E16
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   43
         Top             =   1020
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":1CE2E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   42
         Top             =   1020
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":24C46
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   49
         Top             =   3900
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":2CC5E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   52
         Top             =   3900
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":34C76
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   51
         Top             =   3900
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   9780
         OleObjectBlob   =   "Distiller.frx":3D48E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   50
         Top             =   3900
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image22 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2460
         Picture         =   "Distiller.frx":452A6
         Stretch         =   -1  'True
         Top             =   1780
         Width           =   495
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2520
         X2              =   1260
         Y1              =   1995
         Y2              =   1995
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      TabIndex        =   68
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Distil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer

Private Sub Command1_Click()

' 14/9/00
SALINOMETER.ZOrder


End Sub

Private Sub Command3D2_Click()

' 4/4/00
'sy104_LSS.Show
Salt_Water.Show

End Sub

Private Sub Command3D5_Click()

' 4/4/00
'sy080_LSS.Show

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idsy087)


If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        DFT_dump_reg.Visible = False
        DFT_dump_bypass.Visible = False
        DWT_dump_reg.Visible = False
        DWT_dump_bypass.Visible = False
        
        ' === 13/12/02 17.004 ===
        Command1.Visible = False
Else
        DFT_dump_reg.Visible = True
        DFT_dump_bypass.Visible = True
        DWT_dump_reg.Visible = True
        DWT_dump_bypass.Visible = True
        
        ' === 13/12/02 17.004 ===
        Command1.Visible = True

End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy087)
    HavenStretch1.Visible = True
End If
End Sub

Private Sub Timer2_Timer()

End Sub



Private Sub OFFpump_Click(Index As Integer)


set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 3/8/00
If Index = 0 Then
   ' FWG SW pump  OFF
   Call setPCDI(PumpIO_LSS(4, 38))
ElseIf Index = 1 Then
   ' FWG FW pump  OFF
   Call setPCDI(PumpIO_LSS(4, 39))
End If



End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 3/8/00
If Index = 0 Then
   ' FWG SW pump  ON
   Call setPCDI(PumpIO_LSS(2, 38))
ElseIf Index = 1 Then
   ' FWG FW pump  ON
   Call setPCDI(PumpIO_LSS(2, 39))
End If



End Sub

Private Sub Reset_command_Click()

set_dummy_focus = True

' === 22/11/02 17.033 ===
'equivalence(STEAM_logicals(119),FWG_to_BILGE)
nlog = 119
status_flag = False
' === 22/11/02 ===
'Call set_STEAM_logicals(nlog, status_flag)
Call setSTEAMlogicals(nlog, status_flag)


End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim single_char As String * 1


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


'   Steam_Sim Valves:
' update_vvalve_LSS(syxxx,nValve,index1,index2)
 For i = 0 To 13
   Call update_vvalve_LSS(Me, i + 51, i, -1)
 Next i
 
   Call update_vvalve_LSS(Me, 18, 18, -1)
   Call update_vvalve_LSS(Me, 20, 20, -1)
   Call update_vvalve_LSS(Me, 15, 15, -1)
   Call update_vvalve_LSS(Me, 16, 16, -1)

' === 21/11/02 17.030 ===
   Call update_vvalve_LSS(Me, 105, 105, -1)


' === 2/8/00 ========

Call paint_alarm(Me.Z_DWT, 80, 0, 0, 0)
Z_DWT.text = Format(LSS_reals(452), "###0.0") & " %"
DWT_tank.LineWidth = LSS_reals(452)


' 3/8/00
'   PC_REALS(561) = Z_shore_tank * 100#
Z_shore.text = Format(LSS_reals(561), "###0.0") & " %"
Shore_tank.LineWidth = LSS_reals(561)


'          PC_REALS(552) = F_FWG_DWT
'          PC_REALS(553) = F_shore_tank_FWG
'          PC_REALS(554) = FWG_salinity
'          PC_REALS(555) = F_shore_tank_in
'          PC_REALS(556) = F_shore_tank_out



' === 22/01/02 15.021 ===

'    F_FWG_DWT.text = Format(LSS_reals(552), "###0.0") & " kg/s"
'    F_shore_tank_FWG.text = Format(LSS_reals(553), "###0.0") & " kg/s"
'    FWG_salinity.text = Format(LSS_reals(554), "###0.0") & " ppm"
'    F_shore_tank_in.text = Format(LSS_reals(555), "###0.0") & " kg/s"
'    F_shore_tank_out.text = Format(LSS_reals(556), "###0.0") & " kg/s"
'    F_FWG_to_BILGE.text = Format(LSS_reals(562), "###0.0") & " kg/s"


' === 25/09/02 17.031 ===
' === 25/09/02 17.031 ===
'   FWG_salinity.text = Format(LSS_reals(554) * Flow_mult, "###0.0") & " ppm"
    FWG_salinity.text = Format(LSS_reals(554), "###0.0") & " ppm"



'c       === 21/11/02 ===
'c equivalence(Bunker_Vars(88), FWG_vac)
'c       real*4                      FWG_vac
'        PC_REALS(885) = FWG_vac
FWG_vac.text = Format(LSS_reals(885), "###0.0") & " in Hg"


If SI_units Then
    F_FWG_DWT.text = Format(LSS_reals(552) * Flow_mult, "###0.0") & Flow_chars
    F_shore_tank_FWG.text = Format(LSS_reals(553) * Flow_mult, "###0.0") & Flow_chars
    
    ' === 25/09/02 17.031 ===
    'FWG_salinity.text = Format(LSS_reals(554) * Flow_mult, "###0.0") & Flow_chars
    
    F_shore_tank_in.text = Format(LSS_reals(555) * Flow_mult, "###0.0") & Flow_chars
    F_shore_tank_out.text = Format(LSS_reals(556) * Flow_mult, "###0.0") & Flow_chars
    F_FWG_to_BILGE.text = Format(LSS_reals(562) * Flow_mult, "###0.0") & Flow_chars
Else
    F_FWG_DWT.text = Format(LSS_reals(552) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_shore_tank_FWG.text = Format(LSS_reals(553) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    ' === 25/09/02 17.031 ===
    'FWG_salinity.text = Format(LSS_reals(554) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_shore_tank_in.text = Format(LSS_reals(555) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_shore_tank_out.text = Format(LSS_reals(556) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_FWG_to_BILGE.text = Format(LSS_reals(562) * Flow_mult / 3600#, "###0.0") & " lb/s"
End If


'================================================================
' SAMPLE from Lub_oil...
'TLO_bearings.text = Format(LSS_reals(211) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'If SI_units Then
'    PLO_throttle.text = Format(LSS_reals(217) * Press_mult, "##0.0") & Press_chars     ' Bar
'    FLO_sight_glass.text = Format(LSS_reals(202) * Flow_mult, "###0.0") & Flow_chars
'Else
'    PLO_throttle.text = Format(LSS_reals(217) * Press_mult, "###0") & Press_chars     ' psi
'    FLO_sight_glass.text = Format(LSS_reals(202) * Flow_mult / 3600#, "###0.0") & " lb/s"
'End If
'================================================================




' Drain to Bilge if Salinity is too high

' === 22/11/02 17.033 ===
'equivalence(STEAM_logicals(119),FWG_to_BILGE)
'If LSS_reals(554) > 4.5 Then
If (VB_steam_logicals(119)) Then
  Haven3WayValve1.OpenTopTri = False
  Haven3WayValve1.OpenBtmTri = True
  FWG_salinity.BackColor = &HFF&         'red
Else
  Haven3WayValve1.OpenTopTri = True
  Haven3WayValve1.OpenBtmTri = False
  FWG_salinity.BackColor = &HFF00&      'green
End If


' DFT Dump Regulator
DFT_dump_reg.Caption = Format(LSS_reals(454), "###0") & " %"
' 6/11/99
Call update_regulator(454, DFT_dump(0), DFT_dump(1), DFT_dump(2), DFT_dump(3))

' DFT Dump Bypass
DFT_dump_bypass.Caption = Format(LSS_reals(455), "###0") & " %"

Call update_bypass(455, Valve(21))


' DWT Dump Regulator
DWT_dump_reg.Caption = Format(LSS_reals(456), "###0") & " %"
' 6/11/99
Call update_regulator(456, DWT_dump(0), DWT_dump(1), DWT_dump(2), DWT_dump(3))

' DWT Dump Bypass
DWT_dump_bypass.Caption = Format(LSS_reals(457), "###0") & " %"

Call update_bypass(457, Valve(17))

' === 2/8/00 ========


' === 3/8/00 ===
Call update_pump_LSS(Me, 0, 38)      ' FWG SW pump
Call update_pump_LSS(Me, 1, 39)      ' FWG FW pump



'single_char = Mid$(PC_INT_STRING, 476, 1)
'If single_char = "1" Then
'    HavenPump1(0).PumpOn = True
'Else
'    HavenPump1(0).PumpOn = False
'End If


End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
' 16/10/99
If Index < 14 Then

    Call valve_pcdi_LSS(Index + 51)
    
Else

    Call valve_pcdi_LSS(Index)

End If


End Sub

Private Sub zht90_Click()

' 14/7/98
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 153
nAlarm_LO = 154
Alarm_set_point.Show

End Sub

Sub update_regulator(nReals As Integer, obj_0 As Object, obj_1 As Object, obj_2 As Object, obj_3 As Object)


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
Sub update_bypass(nReals As Integer, obj As Object)
Dim green_colour As Integer

If LSS_reals(nReals) > 99# Then
    green_colour = 255
ElseIf LSS_reals(nReals) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(nReals) * 2.55
End If

obj.ColourClosed = RGB(0, green_colour, 0)

End Sub

Private Sub Z_DWT_Click()

' 2/8/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 80
Alarm_set_point.Show

End Sub
