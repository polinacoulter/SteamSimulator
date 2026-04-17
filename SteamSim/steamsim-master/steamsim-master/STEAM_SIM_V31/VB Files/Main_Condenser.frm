VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Main_Condenser 
   Caption         =   "Main Condenser"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10425
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   11033
      Begin VB.TextBox T_hotwell 
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
         Left            =   8280
         TabIndex        =   75
         Text            =   "27 C"
         Top             =   1860
         Width           =   1155
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
         Index           =   3
         Left            =   6600
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   5160
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
         Index           =   2
         Left            =   6720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   3900
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
         Index           =   3
         Left            =   6120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   5160
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
         Index           =   2
         Left            =   6240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3900
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
         Index           =   3
         Left            =   5640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   5160
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
         Index           =   2
         Left            =   5760
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   3900
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox Tsteam17 
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
         Left            =   6720
         TabIndex        =   57
         Text            =   "89 F"
         Top             =   360
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox Pactual3_Hg 
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
         Left            =   7140
         TabIndex        =   56
         Text            =   "-27.5 inHg"
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox Pactual3 
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
         Left            =   6240
         TabIndex        =   55
         Text            =   "0.047 Bar"
         Top             =   900
         Width           =   795
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   7320
         TabIndex        =   54
         Top             =   3420
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox F_scoop 
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
         Left            =   2400
         TabIndex        =   53
         Text            =   "1200 lbs/hr"
         Top             =   5040
         Width           =   1155
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1635
         Left            =   5100
         Max             =   1000
         TabIndex        =   52
         Top             =   4260
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OPEN / CLOSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3900
         TabIndex        =   51
         Top             =   4860
         Width           =   675
      End
      Begin VB.TextBox Psteam14_Hg 
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
         Left            =   6360
         TabIndex        =   48
         Text            =   "-27.5 inHg"
         Top             =   5460
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   3660
         Top             =   0
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
         Left            =   1320
         TabIndex        =   46
         Text            =   "17 C"
         Top             =   6000
         Width           =   795
      End
      Begin VB.TextBox Vship 
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
         TabIndex        =   45
         Text            =   "6.3 Knotts"
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox Fwater3 
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
         Left            =   8280
         TabIndex        =   44
         Text            =   "1200 lbs/hr"
         Top             =   1560
         Width           =   1155
      End
      Begin VB.TextBox Twater3 
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
         Left            =   8280
         TabIndex        =   43
         Text            =   "27 C"
         Top             =   2220
         Visible         =   0   'False
         Width           =   1155
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
         Left            =   3060
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3300
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
         Left            =   2580
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3300
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
         Left            =   2100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3300
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
         Left            =   3060
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1020
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
         Left            =   2580
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1020
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
         Left            =   2100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1020
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox Pcirc_water 
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
         Left            =   4020
         TabIndex        =   34
         Text            =   "5.0 Bar"
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Condensate System"
         Height          =   615
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox Z_hotwell 
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
         Left            =   7800
         TabIndex        =   20
         Text            =   "20 %"
         Top             =   2820
         Width           =   615
      End
      Begin VB.TextBox Psteam14 
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
         Left            =   5520
         TabIndex        =   17
         Text            =   "0.047 Bar"
         Top             =   5460
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox Tsteam14 
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
         Left            =   5040
         TabIndex        =   16
         Text            =   "89 F"
         Top             =   360
         Width           =   795
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   1140
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
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   1140
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
         Index           =   3
         Left            =   480
         TabIndex        =   5
         Top             =   3315
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
         Left            =   1200
         TabIndex        =   7
         Top             =   3320
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
         Index           =   4
         Left            =   1920
         TabIndex        =   8
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   5
         Left            =   3000
         TabIndex        =   9
         Top             =   1320
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
         Index           =   6
         Left            =   1920
         TabIndex        =   10
         Top             =   2640
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
         Index           =   7
         Left            =   3000
         TabIndex        =   11
         Top             =   2400
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Hotwell_tank 
         Height          =   675
         Left            =   7500
         TabIndex        =   19
         Top             =   2460
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   1191
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   20
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   10
         Left            =   5460
         TabIndex        =   22
         Top             =   3060
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
         Left            =   5460
         TabIndex        =   24
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
         Index           =   13
         Left            =   6900
         TabIndex        =   25
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
      Begin PumpPrj.Pump havenPump1 
         Height          =   495
         Index           =   2
         Left            =   5940
         TabIndex        =   26
         Top             =   3345
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump havenPump1 
         Height          =   495
         Index           =   3
         Left            =   5880
         TabIndex        =   30
         Top             =   4545
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump havenPump1 
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   36
         Top             =   1620
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   49152
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   9
         Left            =   3960
         TabIndex        =   66
         Top             =   4380
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   11
         Left            =   6900
         TabIndex        =   23
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
      Begin PumpPrj.Pump havenPump1 
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   67
         Top             =   2700
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   49152
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   14
         Left            =   6780
         TabIndex        =   68
         Top             =   4200
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
         Height          =   315
         Index           =   15
         Left            =   6420
         TabIndex        =   69
         Top             =   3060
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
         Height          =   435
         Index           =   16
         Left            =   9150
         TabIndex        =   70
         Top             =   1060
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
         Index           =   17
         Left            =   1140
         TabIndex        =   71
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
         Height          =   255
         Index           =   18
         Left            =   5280
         TabIndex        =   74
         Top             =   2820
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Shape Shape26 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   7020
         Shape           =   3  'Circle
         Top             =   2700
         Width           =   135
      End
      Begin VB.Line R_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   5400
         X2              =   5400
         Y1              =   2760
         Y2              =   3420
      End
      Begin VB.Line R_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5400
         X2              =   7080
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line5 
         Visible         =   0   'False
         X1              =   3480
         X2              =   3840
         Y1              =   1620
         Y2              =   1500
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "See Form Load"
         Height          =   555
         Left            =   3840
         TabIndex        =   77
         Top             =   1140
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   23
         X1              =   9540
         X2              =   9540
         Y1              =   840
         Y2              =   1800
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "#2"
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
         Left            =   1800
         TabIndex        =   76
         Top             =   3300
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Main Condenser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   5760
         TabIndex        =   1
         Top             =   1585
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5580
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   45
         X1              =   1500
         X2              =   1740
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Index           =   44
         X1              =   960
         X2              =   960
         Y1              =   4380
         Y2              =   4560
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   660
         X2              =   1260
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   480
         X2              =   1800
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   3  'Dot
         Height          =   255
         Left            =   660
         Top             =   4560
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Index           =   25
         X1              =   960
         X2              =   1260
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   29
         X1              =   7440
         X2              =   7440
         Y1              =   4800
         Y2              =   3120
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   7380
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   255
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   5460
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   22
         X1              =   2880
         X2              =   3600
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   21
         X1              =   2820
         X2              =   3600
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   7260
         Picture         =   "Main_Condenser.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LP Turbine"
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
         Left            =   7620
         TabIndex        =   13
         Top             =   60
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         Index           =   43
         X1              =   7860
         X2              =   8280
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line Line1 
         Index           =   42
         Visible         =   0   'False
         X1              =   7080
         X2              =   7080
         Y1              =   1380
         Y2              =   540
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6060
         X2              =   6060
         Y1              =   480
         Y2              =   1380
      End
      Begin VB.Line Line1 
         Index           =   41
         X1              =   7500
         X2              =   7500
         Y1              =   1380
         Y2              =   1140
      End
      Begin VB.Shape Scoop 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   1680
         Top             =   5520
         Width           =   1275
      End
      Begin VB.Shape Scoop_closed 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   1080
         Top             =   5520
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Shape Scoop_open 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2340
         Top             =   5520
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   34
         X1              =   1140
         X2              =   2400
         Y1              =   5700
         Y2              =   5700
      End
      Begin VB.Line Line1 
         Index           =   31
         X1              =   6660
         X2              =   6660
         Y1              =   1380
         Y2              =   1200
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Abs Press"
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
         Left            =   6240
         TabIndex        =   50
         Top             =   705
         Width           =   795
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gauge Press"
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
         Left            =   7080
         TabIndex        =   49
         Top             =   705
         Width           =   1155
      End
      Begin VB.Label Sea_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sea Speed"
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
         Left            =   3900
         TabIndex        =   47
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   28
         X1              =   6480
         X2              =   6840
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   26
         X1              =   6840
         X2              =   6840
         Y1              =   2960
         Y2              =   4800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         Height          =   195
         Left            =   3900
         TabIndex        =   35
         Top             =   5520
         Width           =   195
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   255
         Left            =   3840
         Top             =   5505
         Width           =   315
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   6480
         X2              =   6480
         Y1              =   1800
         Y2              =   3540
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   27
         X1              =   7500
         X2              =   7500
         Y1              =   2460
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   35
         X1              =   5400
         X2              =   5400
         Y1              =   3360
         Y2              =   4560
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   3900
         Picture         =   "Main_Condenser.frx":030A
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sea Temp"
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
         Left            =   1260
         TabIndex        =   33
         Top             =   5760
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Hotwell"
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
         Left            =   6780
         TabIndex        =   32
         Top             =   2520
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   39
         X1              =   3600
         X2              =   5040
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   38
         X1              =   5400
         X2              =   6120
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   37
         X1              =   5400
         X2              =   6120
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Condensate Pumps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7560
         TabIndex        =   29
         Top             =   3540
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   5040
         Picture         =   "Main_Condenser.frx":0614
         Stretch         =   -1  'True
         Top             =   3780
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   36
         X1              =   5160
         X2              =   5400
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "#2"
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
         Left            =   5400
         TabIndex        =   28
         Top             =   5160
         Width           =   195
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "#1"
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
         Left            =   5460
         TabIndex        =   27
         Top             =   3900
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   33
         X1              =   6360
         X2              =   7440
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   32
         X1              =   6360
         X2              =   7440
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "#1"
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
         Left            =   1800
         TabIndex        =   21
         Top             =   1020
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   5400
         Top             =   1390
         Width           =   2295
      End
      Begin VB.Line Line1 
         Index           =   30
         X1              =   5460
         X2              =   6060
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   24
         X1              =   7500
         X2              =   7500
         Y1              =   2880
         Y2              =   3120
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Circ Pumps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2220
         TabIndex        =   18
         Top             =   2160
         Width           =   795
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   12
         X1              =   120
         X2              =   480
         Y1              =   3840
         Y2              =   3240
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   120
         Top             =   3240
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   8580
         Picture         =   "Main_Condenser.frx":091E
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   20
         X1              =   7800
         X2              =   9540
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Overboard"
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
         Left            =   8640
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   5400
         Picture         =   "Main_Condenser.frx":0C28
         Stretch         =   -1  'True
         Top             =   60
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Astern Turbine"
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
         Left            =   4200
         TabIndex        =   14
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Air Ejector"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Width           =   1155
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   5100
         Picture         =   "Main_Condenser.frx":0F32
         Stretch         =   -1  'True
         Top             =   720
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   19
         X1              =   5220
         X2              =   5580
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   18
         X1              =   5580
         X2              =   5580
         Y1              =   840
         Y2              =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   17
         X1              =   3600
         X2              =   3600
         Y1              =   1620
         Y2              =   2700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   16
         X1              =   1740
         X2              =   2460
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   15
         X1              =   1740
         X2              =   2520
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   14
         X1              =   1740
         X2              =   1740
         Y1              =   1440
         Y2              =   4380
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Suction"
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
         TabIndex        =   6
         Top             =   3000
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   13
         X1              =   120
         X2              =   120
         Y1              =   3000
         Y2              =   4080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   11
         X1              =   120
         X2              =   1740
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "High Suction"
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
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   10
         X1              =   120
         X2              =   120
         Y1              =   840
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   120
         X2              =   480
         Y1              =   1680
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         Index           =   6
         X1              =   4380
         X2              =   5160
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   0
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   5
         X1              =   1560
         X2              =   5160
         Y1              =   5700
         Y2              =   2400
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scoop"
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
         Left            =   1200
         TabIndex        =   0
         Top             =   5220
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   4
         X1              =   2700
         X2              =   3540
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         X1              =   2340
         X2              =   2700
         Y1              =   5700
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   960
         X2              =   1140
         Y1              =   5820
         Y2              =   5700
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   480
         X2              =   960
         Y1              =   6000
         Y2              =   5820
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   240
         X2              =   480
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   7155
         Shape           =   3  'Circle
         Top             =   1330
         Width           =   1095
      End
      Begin VB.Shape Screw 
         BorderWidth     =   2
         FillStyle       =   5  'Downward Diagonal
         Height          =   75
         Left            =   2340
         Top             =   5580
         Width           =   1515
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   885
         Left            =   6300
         OleObjectBlob   =   "Main_Condenser.frx":123C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hotwell.vsd"
         TabIndex        =   64
         Top             =   2400
         Width           =   1275
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   1545
         Left            =   5700
         OleObjectBlob   =   "Main_Condenser.frx":8A54
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   65
         Top             =   -840
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilge Injection"
         Height          =   315
         Left            =   0
         TabIndex        =   72
         Top             =   4140
         Width           =   1035
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   120
         Top             =   1080
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Index           =   8
         X1              =   120
         X2              =   1740
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      TabIndex        =   73
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Main_Condenser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()
    
set_dummy_focus = True


' 23/9/99
'status_flag = True
'Call set_STEAM_logicals(6, status_flag)

Call valve_pcdi_LSS(269)


End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command5_Click()

Condensate.Show

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idMain_Condenser)

set_dummy_focus = True


' === 25/09/02 17.042 ===
If (TARGET_this_sim = CALIFORNIA) Or (TARGET_this_sim = NOWHERE) Then
'   NOT stop check valves...
    Valve(5).EnabledBtmRight = True
    Valve(7).EnabledBtmRight = True
Else
'   Stop check valves...
    Valve(5).EnabledBtmRight = False
    Valve(7).EnabledBtmRight = False
End If



' 4/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' === 16/12/02  ===
    Sea_label.Visible = False
    Vship.Visible = False
    
    ' === 14/03/01 =======
     Command1.Visible = False
     VScroll4.Visible = False
    ' === 14/03/01 end ===
    
    
    ONpump(0).Visible = False
    OFFpump(0).Visible = False
    sbyPUMP(0).Visible = False
    
    ONpump(1).Visible = False
    OFFpump(1).Visible = False
    sbyPUMP(1).Visible = False
    
    ONpump(2).Visible = False
    OFFpump(2).Visible = False
    sbyPUMP(2).Visible = False
    
    ONpump(3).Visible = False
    OFFpump(3).Visible = False
    sbyPUMP(3).Visible = False
    
Else
    ONpump(0).Visible = True
    OFFpump(0).Visible = True
    sbyPUMP(0).Visible = True
    
    ONpump(1).Visible = True
    OFFpump(1).Visible = True
    sbyPUMP(1).Visible = True
    
    ONpump(2).Visible = True
    OFFpump(2).Visible = True
    sbyPUMP(2).Visible = True
    
    ONpump(3).Visible = True
    OFFpump(3).Visible = True
    sbyPUMP(3).Visible = True
    
End If

' === 16/12/02 19.126 ===
Valve(18).Visible = False
R_line(1).Visible = False
R_line(2).Visible = False
Shape26.Visible = False



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idMain_Condenser)
    HavenStretch1.Visible = True
End If
   
End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(3=OFF,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(3, 4))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(3, 5))
'End If

'   Note:  PumpIO(4=OFF,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 2))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 14))
ElseIf Index = 2 Then
   Call setPCDI(PumpIO_LSS(4, 23))
ElseIf Index = 3 Then
   Call setPCDI(PumpIO_LSS(4, 24))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(1=ON,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(1, 4))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(1, 5))
'End If

'   Note:  PumpIO_LSS(2=ON,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 2))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 14))
ElseIf Index = 2 Then
   Call setPCDI(PumpIO_LSS(2, 23))
ElseIf Index = 3 Then
   Call setPCDI(PumpIO_LSS(2, 24))
End If


End Sub

Private Sub Pactual3_Hg_Click()

' === 26/11/01 ===
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 91
Alarm_set_point.Show

End Sub

Private Sub Pcirc_water_Click()

' === 26/11/01 ===
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 97
Alarm_set_point.Show

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(5=Stand-bye,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(5, 4))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(5, 5))
'End If

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 2))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(5, 14))
ElseIf Index = 2 Then
   Call setPCDI(PumpIO_LSS(5, 23))
ElseIf Index = 3 Then
   Call setPCDI(PumpIO_LSS(5, 24))
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim Hotwell_level As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' === 15/03/01 ===
Fwater3.Visible = Advanced_mode
F_scoop.Visible = Advanced_mode

' === 26/11/01 ===
Call paint_alarm(Me.Pactual3_Hg, 91, 0, 0, 0)
Call paint_alarm(Me.Pcirc_water, 97, 0, 0, 0)

' 12/5/00
Call paint_alarm(Me.Z_Hotwell, 92, 0, 0, 0)

List1.clear

' 8/5/00
List1.AddItem "Qavail        = " & Format(LSS_reals(263), "#####0.0") & " MW"

List1.AddItem "Qactual       = " & Format(LSS_reals(264), "#####0.0") & " kW"

' === 28/01/02 ===
List1.AddItem "Q_hotwell_sea = " & Format(LSS_reals(751), "#####0.0") & " kW"
' === 28/01/02 ===
List1.AddItem "T_hotwell = " & Format(LSS_reals(304), "#####0.0000") & Chr(176) & "C"
List1.AddItem "Twater(3) = " & Format(LSS_reals(326), "#####0.0000") & Chr(176) & "C"
List1.AddItem "Fwater(2) = " & Format(LSS_reals(750), "#####0.0") & "kg/s"


List1.AddItem "F_stm_cond   = " & Format(LSS_reals(265), "#####0.0")
List1.AddItem "F_stm_uncond = " & Format(LSS_reals(266), "#####0.0")
List1.AddItem "F_cond       = " & Format(LSS_reals(267), "#####0.0")
List1.AddItem " "
List1.AddItem "M_stm      = " & Format(LSS_reals(268), "#####0.0")
List1.AddItem "P_stm      = " & Format(LSS_reals(269), "#####0.0")



'List1.AddItem "Mair = " & Format(LSS_reals(51), "###0.00")
'List1.AddItem "Pair = " & Format(LSS_reals(50), "###0.00")
'List1.AddItem "T_air = " & Format(LSS_reals(52), "###0.00")
'List1.AddItem "AIR EJ Eff = " & Format(LSS_reals(5), "###0.00")
'List1.AddItem "Tsteam(17) = " & Format(LSS_reals(4), "###0.00")


' Scoop Screw...
Scoop_pos = LSS_reals(6)
If OLD_Scoop_pos <> Scoop_pos Then
    OLD_Scoop_pos = Scoop_pos
    If Screw.FillStyle = 4 Then
        Screw.FillStyle = 5 ' Downward diagonal
    Else
        Screw.FillStyle = 4 ' Upward diagonal
    End If
End If

Scoop.Top = Scoop_closed.Top
Scoop.Left = Scoop_closed.Left + (Scoop_open.Left - Scoop_closed.Left) * LSS_reals(6) * 0.01

Tsteam17.text = Format(LSS_reals(4) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


Tsteam14.text = Format(LSS_reals(74) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' Absolute Pressure...
If SI_units Then
    Psteam14.text = Format(LSS_reals(64) * Press_mult, "##0.000") & Press_chars     ' Bar
    Pactual3.text = Format(LSS_reals(3) * Press_mult, "##0.000") & Press_chars      ' Bar
    
    Z_Hotwell.text = Format(LSS_reals(465) * 25.4, "###0.0") & " mm"
Else
    Psteam14.text = Format(LSS_reals(64) * Press_mult, "##0.0") & Press_chars     ' psi
    Pactual3.text = Format(LSS_reals(3) * Press_mult, "##0.0") & Press_chars     ' psi
    Z_Hotwell.text = Format(LSS_reals(465), "###0.0") & " in"
End If

' Gauge Pressure...
' ONE BAR = 29.527 inHg....
Psteam14_Hg.text = Format(LSS_reals(64) * 29.527 - 29.527, "##0.0") & " inHg"        ' inHg
Pactual3_Hg.text = Format(LSS_reals(3) * 29.527 - 29.527, "##0.0") & " inHg"        ' inHg



Pcirc_water.text = Format(LSS_reals(792) * Press_mult, "##0.0") & Press_chars

Twater3.text = Format(LSS_reals(326) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' 10/5/00
T_hotwell.text = Format(LSS_reals(304) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

If SI_units Then
    Fwater3.text = Format(LSS_reals(327) * Flow_mult, "###0") & Flow_chars
    F_scoop.text = Format(LSS_reals(7) * Flow_mult, "###0") & Flow_chars
Else
    Fwater3.text = Format(LSS_reals(327) * Flow_mult / 3600#, "###0") & " lb/s"
    F_scoop.text = Format(LSS_reals(7) * Flow_mult / 3600#, "###0") & " lb/s"
End If

' 10/5/00 LSS_reals(304) is now T_hotwell
'Tsea.text = Format(LSS_reals(304) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsea.text = Format(LSS_reals(59) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


Vship.text = Format(LSS_reals(54), "##0.0") & " Knotts"

'  Hotwell Tank  ( 0 inches to +8.0 inches )
Hotwell_level = LSS_reals(465) * 100# / 8#
If Hotwell_level > 99 Then
  Hotwell_tank.LineWidth = 100
ElseIf Hotwell_level < 1 Then
  Hotwell_tank.LineWidth = 0
Else
  Hotwell_tank.LineWidth = Hotwell_level
End If

' Note: update_pump_LSS(syxxx,index,nPump)
'Call update_pump_LSS(Main_Condenser, 0, 4)      ' Main Circ Pump # 1
'Call update_pump_LSS(Main_Condenser, 1, 5)      ' Main Circ Pump # 2
Call update_pump_LSS(Main_Condenser, 0, 2)       ' Main Circ Pump # 1
Call update_pump_LSS(Main_Condenser, 1, 14)      ' Main Circ Pump # 2

Call update_pump_LSS(Main_Condenser, 2, 23)      ' Condensate Pump # 1
Call update_pump_LSS(Main_Condenser, 3, 24)      ' Condensate Pump # 2


' update_vvalve_LSS(syxxx,nValve,index1,index2)
For i = 0 To 7
   Call update_vvalve_LSS(Me, 260 + i, i, -1)
Next i
' 12/10/99
'For i = 9 To 13
For i = 9 To 17
   Call update_vvalve_LSS(Me, 260 + i, i, -1)
Next i

   Call update_vvalve_LSS(Me, 13, 18, -1)



End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
   
If Index <> 8 Then

'   16/10/99
    If (Index = 18) Then
      Call valve_pcdi_LSS(13)
    Else
      Call valve_pcdi_LSS(Index + 260)
    End If
    
End If
   
   

End Sub

Private Sub VScroll4_Change()

' 15/7/99  SHIP SPEED (TEST )
Call set_VSCROLL(4, VScroll4.Value)

End Sub

Private Sub Z_Hotwell_Click()

' 12/5/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 92
Alarm_set_point.Show

End Sub
