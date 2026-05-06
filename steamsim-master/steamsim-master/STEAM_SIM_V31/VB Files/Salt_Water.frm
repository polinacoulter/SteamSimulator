VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Salt_Water 
   Caption         =   "Salt Water Services"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   11715
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   14208
      Begin VB.CommandButton Command1 
         Caption         =   "SSDG FW"
         Height          =   615
         Left            =   9180
         TabIndex        =   81
         Top             =   1620
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TLO_main_feed 
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
         Index           =   0
         Left            =   2280
         TabIndex        =   80
         Text            =   "xx F"
         Top             =   300
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TLO_main_feed 
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
         Left            =   2280
         TabIndex        =   79
         Text            =   "xx F"
         Top             =   780
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TSW_main_feed 
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
         Left            =   1020
         TabIndex        =   78
         Text            =   "xx F"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TSW_main_feed 
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
         Index           =   0
         Left            =   1020
         TabIndex        =   77
         Text            =   "xx F"
         Top             =   300
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox PSW_2 
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
         Left            =   1920
         TabIndex        =   73
         Text            =   "0 psi"
         ToolTipText     =   "Main Line Salt Water Pressure"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox PSW_1 
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
         Left            =   1920
         TabIndex        =   72
         Text            =   "0 psi"
         ToolTipText     =   "Main Line Salt Water Pressure"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox FSW_LO_cooler 
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
         Left            =   8520
         TabIndex        =   68
         Text            =   "0 kg/s"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TSW_LO_cooler 
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
         Left            =   4260
         TabIndex        =   67
         Text            =   "xx F"
         Top             =   420
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   120
         Top             =   7320
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
         Left            =   2520
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   4320
         UseMaskColor    =   -1  'True
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
         Left            =   2460
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2940
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
         Index           =   1
         Left            =   2040
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   59
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
         Left            =   1980
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2940
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
         Left            =   1560
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   57
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
         Left            =   1500
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2940
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox PSW_out 
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
         Left            =   3360
         TabIndex        =   48
         Text            =   "0 psi"
         ToolTipText     =   "Main Line Salt Water Pressure"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Tsea 
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
         Left            =   3960
         TabIndex        =   47
         Text            =   "54 F"
         ToolTipText     =   "Sea Temperature"
         Top             =   7680
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   11
         Left            =   1560
         TabIndex        =   0
         Top             =   2340
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
         Index           =   12
         Left            =   2610
         TabIndex        =   1
         Top             =   2100
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
         Index           =   13
         Left            =   1560
         TabIndex        =   2
         Top             =   3720
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
         Index           =   14
         Left            =   2580
         TabIndex        =   3
         Top             =   3480
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
         Index           =   0
         Left            =   1410
         TabIndex        =   8
         Top             =   5760
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
         Index           =   1
         Left            =   660
         TabIndex        =   9
         Top             =   5760
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
         Index           =   2
         Left            =   2940
         TabIndex        =   11
         Top             =   6075
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   3
         Left            =   2940
         TabIndex        =   12
         Top             =   6810
         Width           =   495
         _ExtentX        =   873
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
         Index           =   4
         Left            =   7350
         TabIndex        =   13
         Top             =   5760
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
         Index           =   5
         Left            =   8115
         TabIndex        =   14
         Top             =   5760
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
         Index           =   6
         Left            =   5220
         TabIndex        =   15
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
         Index           =   7
         Left            =   7110
         TabIndex        =   22
         Top             =   1980
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
         Index           =   8
         Left            =   5220
         TabIndex        =   24
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
         Index           =   9
         Left            =   7110
         TabIndex        =   25
         Top             =   2460
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
         Index           =   15
         Left            =   5160
         TabIndex        =   26
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   18
         Left            =   5160
         TabIndex        =   29
         Top             =   4260
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
         Index           =   19
         Left            =   7050
         TabIndex        =   31
         Top             =   4260
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
         Index           =   20
         Left            =   5160
         TabIndex        =   32
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   21
         Left            =   7050
         TabIndex        =   34
         Top             =   4740
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
         Index           =   22
         Left            =   5160
         TabIndex        =   35
         Top             =   3300
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
         Index           =   23
         Left            =   7140
         TabIndex        =   37
         Top             =   3300
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
         Index           =   24
         Left            =   8430
         TabIndex        =   38
         Top             =   3420
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
         Index           =   25
         Left            =   5220
         TabIndex        =   39
         Top             =   1020
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
         Index           =   26
         Left            =   7110
         TabIndex        =   41
         Top             =   1020
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
         Index           =   27
         Left            =   5220
         TabIndex        =   42
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   28
         Left            =   7110
         TabIndex        =   44
         Top             =   1560
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
         Index           =   29
         Left            =   5220
         TabIndex        =   49
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
         Index           =   30
         Left            =   7110
         TabIndex        =   51
         Top             =   2880
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
         Index           =   31
         Left            =   5160
         TabIndex        =   52
         Top             =   5220
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
         Left            =   7050
         TabIndex        =   54
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   17
         Left            =   7050
         TabIndex        =   28
         Top             =   3780
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
         Index           =   10
         Left            =   5220
         TabIndex        =   63
         Top             =   600
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
         Index           =   16
         Left            =   7110
         TabIndex        =   64
         Top             =   600
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   65
         Top             =   3705
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   32768
         CentrLineWidth  =   3
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   2040
         TabIndex        =   66
         Top             =   2385
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   32768
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   33
         Left            =   4320
         TabIndex        =   70
         Top             =   1800
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
         Index           =   34
         Left            =   4260
         TabIndex        =   71
         Top             =   4500
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
         Height          =   495
         Index           =   35
         Left            =   5220
         TabIndex        =   75
         Top             =   -60
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
         Index           =   36
         Left            =   7080
         TabIndex        =   76
         Top             =   -60
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
      Begin VB.Line Line11 
         X1              =   7680
         X2              =   7680
         Y1              =   660
         Y2              =   900
      End
      Begin VB.Line Line10 
         X1              =   8640
         X2              =   8100
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line9 
         X1              =   7680
         X2              =   7860
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Image Arrow_left 
         Height          =   255
         Left            =   6720
         Picture         =   "Salt_Water.frx":0000
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   1920
         Picture         =   "Salt_Water.frx":030A
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   3540
         Picture         =   "Salt_Water.frx":0614
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "FWG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5760
         TabIndex        =   74
         ToolTipText     =   "Turbo Alternator Number 1 Lube Oil Cooler"
         Top             =   180
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5700
         Top             =   120
         Width           =   1275
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   42
         X1              =   4080
         X2              =   7980
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   900
         Picture         =   "Salt_Water.frx":091E
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   3600
         Picture         =   "Salt_Water.frx":0C28
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   41
         X1              =   1020
         X2              =   3720
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   40
         X1              =   1020
         X2              =   1020
         Y1              =   2640
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   39
         X1              =   4080
         X2              =   4860
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   3960
         Picture         =   "Salt_Water.frx":0F32
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   38
         X1              =   4860
         X2              =   4860
         Y1              =   1590
         Y2              =   2520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   32
         X1              =   4080
         X2              =   5040
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   31
         X1              =   7620
         X2              =   7980
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   30
         X1              =   5040
         X2              =   5040
         Y1              =   3180
         Y2              =   3600
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   27
         X1              =   7620
         X2              =   7620
         Y1              =   3180
         Y2              =   3600
      End
      Begin VB.Line Line2 
         X1              =   4620
         X2              =   6780
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   6780
         X2              =   6780
         Y1              =   540
         Y2              =   840
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "MAIN LO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5760
         TabIndex        =   62
         ToolTipText     =   "Turbo Alternator Number 1 Lube Oil Cooler"
         Top             =   780
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5700
         Top             =   720
         Width           =   1275
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   0
         Left            =   7860
         Picture         =   "Salt_Water.frx":123C
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   2
         Left            =   3960
         Picture         =   "Salt_Water.frx":1546
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   7860
         Picture         =   "Salt_Water.frx":1850
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   29
         X1              =   7680
         X2              =   7980
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   26
         X1              =   7650
         X2              =   7980
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   24
         X1              =   7620
         X2              =   7620
         Y1              =   2280
         Y2              =   2760
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   23
         X1              =   7650
         X2              =   7950
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   11
         X1              =   7620
         X2              =   7620
         Y1              =   1320
         Y2              =   1860
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   10
         X1              =   4080
         X2              =   4980
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   8
         X1              =   4860
         X2              =   5040
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   7
         X1              =   4860
         X2              =   5100
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   6
         X1              =   7650
         X2              =   7650
         Y1              =   4080
         Y2              =   5520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   5
         X1              =   5010
         X2              =   5010
         Y1              =   4080
         Y2              =   5520
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "#1 Feed p/p LO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5760
         TabIndex        =   50
         ToolTipText     =   "Feed Pump # 1  Lube Oil Cooler"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Line Line8 
         BorderStyle     =   3  'Dot
         X1              =   7830
         X2              =   8070
         Y1              =   6180
         Y2              =   5880
      End
      Begin VB.Line Line7 
         BorderStyle     =   3  'Dot
         X1              =   1140
         X2              =   1380
         Y1              =   6180
         Y2              =   5880
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   9
         X1              =   2475
         X2              =   3265
         Y1              =   3735
         Y2              =   3735
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   0
         X1              =   2400
         X2              =   3270
         Y1              =   2415
         Y2              =   2400
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "DG Alt Air"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5670
         TabIndex        =   53
         ToolTipText     =   "Diesel Generator Alternator Air Cooler"
         Top             =   5460
         Width           =   1035
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   37
         X1              =   5040
         X2              =   7620
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Shape Shape17 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5610
         Top             =   5400
         Width           =   1275
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   36
         X1              =   5040
         X2              =   7620
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   35
         X1              =   3270
         X2              =   4080
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   34
         X1              =   5070
         X2              =   5070
         Y1              =   2280
         Y2              =   2760
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   33
         X1              =   3270
         X2              =   3270
         Y1              =   2400
         Y2              =   3740
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   8160
         Picture         =   "Salt_Water.frx":1B5A
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   2
         X1              =   4080
         X2              =   4080
         Y1              =   300
         Y2              =   4800
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "#2SSTG Alt Air"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5730
         TabIndex        =   46
         ToolTipText     =   "Turbo Alternator Number 2 Alternator Air Cooler"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "#2 SSTG  LO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5730
         TabIndex        =   45
         ToolTipText     =   "Turbo Alternator Number 2 Lube Oil Cooler"
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   28
         X1              =   5100
         X2              =   5100
         Y1              =   1320
         Y2              =   1860
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   360
         X2              =   360
         Y1              =   5760
         Y2              =   6540
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "#1SSTG Alt Air"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5730
         TabIndex        =   43
         ToolTipText     =   "Turbo Alternator Number 1 Alternator Air Cooler"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   22
         X1              =   5100
         X2              =   7620
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Shape Shape15 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5670
         Top             =   1740
         Width           =   1275
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   " #1 SSTG  LO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5730
         TabIndex        =   40
         ToolTipText     =   "Turbo Alternator Number 1 Lube Oil Cooler"
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   21
         X1              =   5100
         X2              =   7620
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5670
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   20
         X1              =   8400
         X2              =   7980
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   19
         X1              =   7980
         X2              =   7980
         Y1              =   300
         Y2              =   4860
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "#2 Feed p/p LO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5700
         TabIndex        =   36
         ToolTipText     =   "Feed Pump # 2  Lube Oil Cooler"
         Top             =   3540
         Width           =   1245
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   18
         X1              =   5040
         X2              =   7620
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5640
         Top             =   3480
         Width           =   1305
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "DG LO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5670
         TabIndex        =   33
         ToolTipText     =   "Diesel Generator Lube Oil Cooler"
         Top             =   4980
         Width           =   1035
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   17
         X1              =   5040
         X2              =   7620
         Y1              =   5040
         Y2              =   5070
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5610
         Top             =   4920
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "DG JCW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5670
         TabIndex        =   30
         ToolTipText     =   "Diesel Generator Jacket Cooling Water Cooler"
         Top             =   4500
         Width           =   1035
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   16
         X1              =   5010
         X2              =   7650
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Shape Shape11 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5580
         Top             =   4440
         Width           =   1305
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "DG Charge Air"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5670
         TabIndex        =   27
         ToolTipText     =   "Diesel Generator Charge Air Cooling"
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   15
         X1              =   5040
         X2              =   7650
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5610
         Top             =   3960
         Width           =   1275
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   14
         X1              =   5100
         X2              =   7620
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5670
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Hull"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6360
         TabIndex        =   23
         ToolTipText     =   "A fishing town in the north of England"
         Top             =   7530
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   13
         X1              =   5070
         X2              =   7620
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   2760
         X2              =   3480
         Y1              =   7500
         Y2              =   7500
      End
      Begin VB.Label Label11 
         Height          =   195
         Left            =   2775
         TabIndex        =   21
         Top             =   7380
         Width           =   615
      End
      Begin VB.Label Label10 
         Height          =   615
         Left            =   300
         TabIndex        =   20
         Top             =   5775
         Width           =   195
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   8820
         X2              =   8820
         Y1              =   5700
         Y2              =   6420
      End
      Begin VB.Label Label9 
         Height          =   615
         Left            =   8640
         TabIndex        =   19
         Top             =   5775
         Width           =   315
      End
      Begin VB.Label Label8 
         Caption         =   "Upper Port Suction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   6480
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Upper Stbd Suction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6900
         TabIndex        =   17
         Top             =   6480
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Lower Port Suction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   7200
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5670
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   2940
         X2              =   3180
         Y1              =   6780
         Y2              =   6480
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2880
         Top             =   6450
         Width           =   375
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   675
         Left            =   8500
         Top             =   5760
         Width           =   315
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   315
         Left            =   2760
         Top             =   7200
         Width           =   675
      End
      Begin VB.Label Label1 
         Height          =   915
         Left            =   300
         TabIndex        =   10
         Top             =   870
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   675
         Left            =   360
         Top             =   5760
         Width           =   305
      End
      Begin VB.Label Label2 
         Height          =   795
         Left            =   8160
         TabIndex        =   6
         Top             =   900
         Width           =   855
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   1
         X1              =   3720
         X2              =   3720
         Y1              =   5340
         Y2              =   6060
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "No 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "No 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   4
         Top             =   2940
         Width           =   435
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   3
         X1              =   1020
         X2              =   2100
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   12
         X1              =   1020
         X2              =   2160
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Ovbd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   8340
         TabIndex        =   55
         ToolTipText     =   "Overboard Valve"
         Top             =   3240
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape18 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1035
         Top             =   5880
         Width           =   375
      End
      Begin VB.Shape Shape19 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   7740
         Top             =   5880
         Width           =   375
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   25
         X1              =   660
         X2              =   8160
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   5700
         Top             =   3060
         Width           =   1275
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   4
         X1              =   4080
         X2              =   7980
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   8235
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   6615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   900
         Width           =   8475
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      TabIndex        =   69
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Salt_Water"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()

' === 13/01/05 For JBF===
SSDG_FW.ZOrder

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idSalt_Water)


' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    ONpump(0).Visible = False
    OFFpump(0).Visible = False
    sbyPUMP(0).Visible = False
    ONpump(1).Visible = False
    OFFpump(1).Visible = False
    sbyPUMP(1).Visible = False
Else
    ONpump(0).Visible = True
    OFFpump(0).Visible = True
    sbyPUMP(0).Visible = True
    ONpump(1).Visible = True
    OFFpump(1).Visible = True
    sbyPUMP(1).Visible = True
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idSalt_Water)
    HavenStretch1.Visible = True
End If
   
End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)

' 3/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 1))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 13))
End If


End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)


' 3/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 1))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 13))
End If

End Sub

Private Sub PSW_out_Click()


Unload Alarm_set_point   ' Do this first
nAlarm_LO = 98
'nAlarm_LOLO = 12
Alarm_set_point.Show

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
' 3/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 1))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(5, 13))
End If


End Sub

Private Sub Timer1_Timer()
Dim i As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

'c 24 / 11 / 99
'        PC_REALS(395) = PSW_1
'        PC_REALS(396) = PSW_2


' 10/8/00 LSS_reals(304) is  T_hotwell NOT Tsea
'Tsea.text = Format(LSS_reals(304) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsea.text = Format(LSS_reals(59) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char



TSW_LO_cooler.text = Format(LSS_reals(214) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' 10/8/00
'          PC_REALS(313) = TSW_main_feed(1)
'          PC_REALS(314) = TSW_main_feed(2)
TSW_main_feed(0).text = Format(LSS_reals(313) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TSW_main_feed(1).text = Format(LSS_reals(314) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

'          PC_REALS(155) = TLO_main_feed(1)
'          PC_REALS(156) = TLO_main_feed(2)
TLO_main_feed(0).text = Format(LSS_reals(155) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TLO_main_feed(1).text = Format(LSS_reals(156) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char



If SI_units Then

    ' 24/11/99
    PSW_1.text = Format(LSS_reals(395) * Press_mult, "##0.0") & Press_chars     ' Bar
    PSW_2.text = Format(LSS_reals(396) * Press_mult, "##0.0") & Press_chars     ' Bar


    PSW_out.text = Format(LSS_reals(212) * Press_mult, "##0.0") & Press_chars     ' Bar
    FSW_LO_cooler.text = Format(LSS_reals(213) * Flow_mult, "###0.000") & Flow_chars
Else

    ' 24/11/99
    PSW_1.text = Format(LSS_reals(395) * Press_mult, "###0") & Press_chars     ' psi
    PSW_2.text = Format(LSS_reals(396) * Press_mult, "###0") & Press_chars     ' psi

    PSW_out.text = Format(LSS_reals(212) * Press_mult, "###0") & Press_chars     ' psi
    FSW_LO_cooler.text = Format(LSS_reals(213) * Flow_mult / 3600#, "###0.000") & " lb/s"

End If


Call paint_alarm(Me.PSW_out, 98, 0, 0, 0)


' Note: update_pump_LSS(syxxx,index,nPump)
'   NEW Steam_Sim Pump numbers
'Call update_pump_LSS(Salt_Water, 0, 11)      ' SW Supply Pump # 1
'Call update_pump_LSS(Salt_Water, 1, 12)      ' SW Supply Pump # 2
Call update_pump_LSS(Salt_Water, 0, 1)      ' SW Supply Pump # 1
Call update_pump_LSS(Salt_Water, 1, 13)      ' SW Supply Pump # 2


' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)

'  16/9/99
For i = 0 To 32
   Call update_vvalve_LSS(Salt_Water, i + 330, i, -1)
Next i

' *** 19/11/99 ***
' SSDG SW supply
   Call update_vvalve_LSS(Salt_Water, 405, 34, -1)
' SSTG SW supply
   Call update_vvalve_LSS(Salt_Water, 389, 33, -1)

' 3/8/00 FWG SW supply IN
Call update_vvalve_LSS(Salt_Water, 56, 35, -1)
' 3/8/00 FWG SW supply OUT
Call update_vvalve_LSS(Salt_Water, 57, 36, -1)
    
    
End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
   
 ' 3/4/00  Do NOT allow clicks on MECD valves
If Index > 10 And Index < 15 Then
  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
End If

  
'  *** 19/11/99 ***
If Index = 34 Then
    Call valve_pcdi_LSS(405)        ' SSDG SW supply
ElseIf Index = 33 Then
    Call valve_pcdi_LSS(389)        ' SSTG SW supply

ElseIf Index = 35 Then
    Call valve_pcdi_LSS(56)        ' 3/8/00 FWG SW supply IN
ElseIf Index = 36 Then
    Call valve_pcdi_LSS(57)        ' 3/8/00 FWG SW supply OUT


Else
    'See Salt_Water_Services_System Valve X-Ref.xls
    Call valve_pcdi_LSS(Index + 330)
End If


End Sub
