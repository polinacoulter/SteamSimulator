VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form SSDG_FW 
   Caption         =   "SSDG"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   10050
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8655
      Left            =   0
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   15266
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   1
         Left            =   5460
         MousePointer    =   1  'Arrow
         TabIndex        =   70
         Top             =   2280
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   7080
         TabIndex        =   69
         Top             =   960
         Visible         =   0   'False
         Width           =   2115
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
         Height          =   195
         Index           =   1
         Left            =   2280
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   5220
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
         Index           =   1
         Left            =   2520
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4500
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
         Index           =   1
         Left            =   2160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   4500
         UseMaskColor    =   -1  'True
         Width           =   435
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
         Height          =   195
         Index           =   0
         Left            =   1800
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1800
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox Psea 
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
         Left            =   120
         TabIndex        =   60
         Text            =   "500 psi"
         Top             =   7920
         Width           =   735
      End
      Begin VB.TextBox Psea_DA1 
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
         Left            =   240
         TabIndex        =   59
         Text            =   "500 psi"
         Top             =   6000
         Width           =   735
      End
      Begin VB.TextBox Tjack 
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
         Index           =   1
         Left            =   6360
         TabIndex        =   58
         Text            =   "500 F"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox Pjack 
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
         Index           =   1
         Left            =   5160
         TabIndex        =   57
         Text            =   "500 psi"
         Top             =   5280
         Width           =   735
      End
      Begin VB.TextBox Zjcw 
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
         Left            =   6840
         TabIndex        =   56
         Text            =   "10.0 %"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Texhaust 
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
         Index           =   1
         Left            =   1560
         TabIndex        =   55
         Text            =   "500 F"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Tlub 
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
         Index           =   1
         Left            =   2040
         TabIndex        =   54
         Text            =   "500 F"
         Top             =   6720
         Width           =   735
      End
      Begin VB.TextBox Plub 
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
         Index           =   1
         Left            =   2040
         TabIndex        =   53
         Text            =   "500 psi"
         Top             =   6360
         Width           =   735
      End
      Begin VB.TextBox Zsump 
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
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   52
         Text            =   "10 %"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox Pdalt 
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
         Index           =   1
         Left            =   3360
         TabIndex        =   51
         Text            =   "500 psi"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox PCA30 
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
         Left            =   1560
         TabIndex        =   50
         Text            =   "500 psi"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   2
         Left            =   7200
         TabIndex        =   48
         Text            =   "*"
         Top             =   5280
         Width           =   210
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   4920
         TabIndex        =   47
         Text            =   "*"
         Top             =   5280
         Width           =   210
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   0
         Left            =   2880
         TabIndex        =   46
         Text            =   "*"
         Top             =   6360
         Width           =   210
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   765
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   45
         Text            =   "SSDG_FW.frx":0000
         Top             =   7920
         Width           =   3870
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
         Left            =   1320
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1800
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
         Left            =   1320
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.TextBox Tair 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Text            =   "TEMP"
         Top             =   3120
         Width           =   800
      End
      Begin VB.TextBox Tscav 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4560
         TabIndex        =   2
         Text            =   "TEMP"
         Top             =   1920
         Width           =   800
      End
      Begin VB.CommandButton Command4 
         Caption         =   "SW PUMP"
         Height          =   495
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6450
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DO TANK"
         Height          =   555
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2280
         Width           =   645
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   5160
      End
      Begin PumpPrj.Pump SSDG_JCW_Pump 
         Height          =   285
         Left            =   5250
         TabIndex        =   4
         Top             =   5580
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         Colour3         =   49152
         PumpOpen        =   -1  'True
         State           =   4
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin PumpPrj.Pump SSDG_LO_Pump 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   5460
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Colour3         =   49152
         PumpOpen        =   -1  'True
         State           =   4
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   19
         Left            =   1650
         TabIndex        =   6
         Top             =   3630
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Top             =   300
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
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   1860
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
         Index           =   24
         Left            =   1620
         TabIndex        =   9
         Top             =   3225
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
         Index           =   4
         Left            =   990
         TabIndex        =   10
         Top             =   810
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   26
         Left            =   5910
         TabIndex        =   11
         Top             =   5760
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   714
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   27
         Left            =   5190
         TabIndex        =   12
         Top             =   6240
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
         Index           =   28
         Left            =   3360
         TabIndex        =   13
         Top             =   6240
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
         Index           =   12
         Left            =   6990
         TabIndex        =   14
         Top             =   0
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
         Height          =   285
         Index           =   30
         Left            =   2500
         TabIndex        =   15
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   31
         Left            =   2500
         TabIndex        =   16
         Top             =   1400
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   32
         Left            =   3360
         TabIndex        =   17
         Top             =   5820
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   33
         Left            =   4110
         TabIndex        =   18
         Top             =   5820
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
         Height          =   345
         Index           =   34
         Left            =   5250
         TabIndex        =   19
         Top             =   5850
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   18
         Left            =   6180
         TabIndex        =   20
         Top             =   810
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   65280
         BottomColour    =   65280
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   36
         Left            =   3450
         TabIndex        =   21
         Top             =   5280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   285
         Index           =   1
         Left            =   2490
         TabIndex        =   22
         Top             =   4830
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         CentrifugalPump =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   37
         Left            =   5940
         TabIndex        =   23
         Top             =   6240
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
         Index           =   38
         Left            =   4110
         TabIndex        =   24
         Top             =   6240
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
         Left            =   8520
         TabIndex        =   25
         Top             =   7440
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
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Top             =   7080
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
         Height          =   315
         Index           =   42
         Left            =   2100
         TabIndex        =   27
         Top             =   4740
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   20
         Left            =   2040
         TabIndex        =   28
         Top             =   300
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Z_JCW 
         Height          =   615
         Left            =   5880
         TabIndex        =   29
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
         TopColour       =   16711680
         LineWidth       =   20
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   5
         Left            =   7680
         TabIndex        =   49
         Top             =   2180
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   16711680
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   0
         Left            =   7860
         TabIndex        =   64
         Top             =   5940
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   3
         Left            =   8220
         TabIndex        =   65
         Top             =   6600
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
      Begin VB.Line Alt_rpm_line 
         BorderStyle     =   2  'Dash
         X1              =   5880
         X2              =   5880
         Y1              =   2460
         Y2              =   4320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2600
         X2              =   2600
         Y1              =   1440
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   8340
         X2              =   8340
         Y1              =   7740
         Y2              =   4050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         FillStyle       =   7  'Diagonal Cross
         Height          =   315
         Left            =   690
         Top             =   1410
         Width           =   135
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   930
         Picture         =   "SSDG_FW.frx":006F
         Stretch         =   -1  'True
         Top             =   6810
         Width           =   255
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   1380
         Picture         =   "SSDG_FW.frx":0379
         Stretch         =   -1  'True
         Top             =   480
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   3090
         Picture         =   "SSDG_FW.frx":0683
         Stretch         =   -1  'True
         Top             =   2430
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   5310
         Picture         =   "SSDG_FW.frx":098D
         Stretch         =   -1  'True
         Top             =   2670
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   6810
         Picture         =   "SSDG_FW.frx":0C97
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   7650
         Picture         =   "SSDG_FW.frx":0FA1
         Stretch         =   -1  'True
         Top             =   3930
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   7590
         Picture         =   "SSDG_FW.frx":12AB
         Stretch         =   -1  'True
         Top             =   7650
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   6540
         Picture         =   "SSDG_FW.frx":15B5
         Stretch         =   -1  'True
         Top             =   6810
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   4920
         Picture         =   "SSDG_FW.frx":18BF
         Stretch         =   -1  'True
         Top             =   6810
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   6
         Left            =   4680
         Picture         =   "SSDG_FW.frx":1BC9
         Stretch         =   -1  'True
         Top             =   6810
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   3090
         Picture         =   "SSDG_FW.frx":1ED3
         Stretch         =   -1  'True
         Top             =   6810
         Width           =   255
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   690
         Picture         =   "SSDG_FW.frx":21DD
         Stretch         =   -1  'True
         Top             =   7260
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   7
         Left            =   1320
         Picture         =   "SSDG_FW.frx":24E7
         Stretch         =   -1  'True
         Top             =   5250
         Width           =   255
      End
      Begin VB.Line Line126 
         BorderColor     =   &H00FF0000&
         X1              =   6270
         X2              =   6270
         Y1              =   600
         Y2              =   3180
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "ALTERNATOR"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6480
         TabIndex        =   42
         Top             =   4500
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Line Line124 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   1080
         X2              =   1080
         Y1              =   7380
         Y2              =   3570
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "SUMP "
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3330
         TabIndex        =   41
         Top             =   4620
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "AIR COOLER"
         Height          =   255
         Left            =   7680
         TabIndex        =   40
         Top             =   3660
         Width           =   1035
      End
      Begin VB.Label Label18 
         Caption         =   "AIR COMPRESSOR"
         Height          =   315
         Left            =   570
         TabIndex        =   39
         Top             =   2130
         Width           =   1515
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   810
         OleObjectBlob   =   "SSDG_FW.frx":27F1
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Turbineup.vsd"
         TabIndex        =   38
         Top             =   1830
         Width           =   405
      End
      Begin VB.Line Line123 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   9180
         X2              =   9180
         Y1              =   5340
         Y2              =   8190
      End
      Begin VB.Label Label17 
         Caption         =   "OVERBOARD"
         Height          =   225
         Left            =   7500
         TabIndex        =   37
         Top             =   8280
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "Fuel Oil"
         Height          =   345
         Left            =   7080
         TabIndex        =   36
         Top             =   2640
         Width           =   585
      End
      Begin VB.Line Line122 
         BorderStyle     =   2  'Dash
         X1              =   5730
         X2              =   6000
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line121 
         BorderStyle     =   2  'Dash
         X1              =   6300
         X2              =   6510
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line120 
         BorderStyle     =   2  'Dash
         X1              =   4980
         X2              =   4980
         Y1              =   2190
         Y2              =   3210
      End
      Begin VB.Line Line119 
         BorderStyle     =   2  'Dash
         X1              =   1890
         X2              =   2370
         Y1              =   3060
         Y2              =   3330
      End
      Begin VB.Line Line118 
         BorderStyle     =   2  'Dash
         X1              =   1890
         X2              =   1890
         Y1              =   2820
         Y2              =   3060
      End
      Begin VB.Line Line117 
         BorderStyle     =   2  'Dash
         X1              =   2490
         X2              =   2490
         Y1              =   6120
         Y2              =   6840
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "JCW"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   5520
         TabIndex        =   35
         Top             =   5880
         Width           =   435
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LO"
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   5880
         Width           =   405
      End
      Begin VB.Label Label13 
         Caption         =   "JCW EXPANSION TANK"
         Height          =   645
         Left            =   4860
         TabIndex        =   33
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "FROM  SW SYSTEM"
         Height          =   615
         Left            =   840
         TabIndex        =   32
         Top             =   8220
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Line Line116 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   6270
         X2              =   6270
         Y1              =   2790
         Y2              =   6060
      End
      Begin VB.Label Label11 
         Caption         =   "AIR RECEIVER"
         Height          =   465
         Left            =   1500
         TabIndex        =   31
         Top             =   690
         Width           =   825
      End
      Begin VB.Line Line115 
         X1              =   2580
         X2              =   2220
         Y1              =   3030
         Y2              =   3270
      End
      Begin VB.Line Line114 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   1920
         X2              =   2970
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Line Line113 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   6720
         X2              =   8310
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line112 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5670
         X2              =   5670
         Y1              =   2790
         Y2              =   3180
      End
      Begin VB.Line Line111 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5220
         X2              =   5220
         Y1              =   2790
         Y2              =   3210
      End
      Begin VB.Line Line110 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4770
         X2              =   4770
         Y1              =   2790
         Y2              =   3180
      End
      Begin VB.Line Line109 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4380
         X2              =   4380
         Y1              =   2790
         Y2              =   3150
      End
      Begin VB.Line Line108 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3960
         X2              =   3960
         Y1              =   2790
         Y2              =   3210
      End
      Begin VB.Line Line107 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3540
         X2              =   3540
         Y1              =   2790
         Y2              =   3150
      End
      Begin VB.Line Line106 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   7620
         X2              =   7950
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line105 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   7560
         X2              =   8340
         Y1              =   4050
         Y2              =   4050
      End
      Begin VB.Line Line103 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   7980
         X2              =   7980
         Y1              =   7380
         Y2              =   4440
      End
      Begin VB.Line Line102 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   1470
         X2              =   1470
         Y1              =   7500
         Y2              =   7770
      End
      Begin VB.Line Line101 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2820
         X2              =   2820
         Y1              =   4980
         Y2              =   5640
      End
      Begin VB.Line Line100 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   4110
         X2              =   4560
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line99 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   6030
         X2              =   5760
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Line Line98 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   6030
         X2              =   6030
         Y1              =   4020
         Y2              =   5700
      End
      Begin VB.Line Line97 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   6540
         Y2              =   7140
      End
      Begin VB.Line Line96 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3540
         X2              =   6270
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Line Line95 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   4110
         X2              =   4800
         Y1              =   6540
         Y2              =   6540
      End
      Begin VB.Line Line94 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   3180
         X2              =   3780
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   3750
         Top             =   6240
         Width           =   345
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   3750
         Top             =   5820
         Width           =   345
      End
      Begin VB.Line Line93 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   5070
         X2              =   5640
         Y1              =   6570
         Y2              =   6570
      End
      Begin VB.Line Line92 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   5880
         X2              =   6660
         Y1              =   6570
         Y2              =   6570
      End
      Begin VB.Line Line91 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   1920
         X2              =   1920
         Y1              =   6090
         Y2              =   4200
      End
      Begin VB.Line Line90 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   4590
         X2              =   4590
         Y1              =   5610
         Y2              =   6120
      End
      Begin VB.Line Line89 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   1920
         X2              =   3480
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line88 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   9000
         X2              =   1470
         Y1              =   7770
         Y2              =   7770
      End
      Begin VB.Line Line87 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   5070
         X2              =   5040
         Y1              =   6570
         Y2              =   7380
      End
      Begin VB.Line Line86 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   6660
         X2              =   6660
         Y1              =   7500
         Y2              =   7740
      End
      Begin VB.Line Line85 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   7560
         Y2              =   7770
      End
      Begin VB.Line Line84 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   6660
         X2              =   6660
         Y1              =   6540
         Y2              =   7170
      End
      Begin VB.Line Line83 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   3210
         X2              =   3210
         Y1              =   6600
         Y2              =   7350
      End
      Begin VB.Line Line82 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2820
         X2              =   4590
         Y1              =   5640
         Y2              =   5610
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5160
         X2              =   5820
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line81 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5160
         X2              =   5160
         Y1              =   5730
         Y2              =   6060
      End
      Begin VB.Line Line80 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5160
         X2              =   6030
         Y1              =   5730
         Y2              =   5730
      End
      Begin VB.Line Line79 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5880
         X2              =   6270
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   5580
         Top             =   5820
         Width           =   345
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   5580
         Top             =   6240
         Width           =   345
      End
      Begin VB.Line Line78 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   7950
         X2              =   0
         Y1              =   7380
         Y2              =   7380
      End
      Begin VB.Line Line77 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   1470
         X2              =   1470
         Y1              =   7350
         Y2              =   3960
      End
      Begin VB.Line Line76 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   1080
         X2              =   2760
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line Line75 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   1470
         X2              =   2790
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line74 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   3210
         X2              =   3210
         Y1              =   600
         Y2              =   3210
      End
      Begin VB.Line Line73 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2790
         X2              =   3210
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   2400
         Top             =   510
         Width           =   375
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   930
         Top             =   1440
         Width           =   195
      End
      Begin VB.Line Line72 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6600
         X2              =   8160
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line71 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   6630
         X2              =   6630
         Y1              =   3660
         Y2              =   2520
      End
      Begin VB.Line Line70 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6390
         X2              =   6630
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line69 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   5670
         X2              =   6150
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line68 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   1920
         X2              =   3060
         Y1              =   4230
         Y2              =   4230
      End
      Begin VB.Line Line44 
         BorderStyle     =   2  'Dash
         X1              =   7140
         X2              =   7140
         Y1              =   3330
         Y2              =   4350
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   1080
         X2              =   1080
         Y1              =   1410
         Y2              =   600
      End
      Begin VB.OLE OLE1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Height          =   2745
         Left            =   2160
         OleObjectBlob   =   "SSDG_FW.frx":4409
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\SSDGVIS.vsd"
         TabIndex        =   30
         Top             =   2520
         Width           =   5775
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2910
         X2              =   3090
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   300
         Width           =   375
      End
      Begin VB.Shape Shape18 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   1020
         Width           =   375
      End
      Begin VB.Line Line125 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         X1              =   2520
         X2              =   1080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line127 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         X1              =   960
         X2              =   690
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1200
      TabIndex        =   63
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "SSDG_FW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()


End Sub

Private Sub Command3_Click()

FO_Service.Show

End Sub

Private Sub Command4_Click()

Salt_Water.Show


End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_SSDG)

' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    ONpump(1).Visible = False
    OFFpump(1).Visible = False
    
    ' === 28/03/03 ===
    ALT_rpm(1).Visible = False
    Alt_rpm_line.Visible = False
    
Else
    ONpump(1).Visible = True
    OFFpump(1).Visible = True
    
    ' === 28/03/03 ===
    ALT_rpm(1).Visible = True
    Alt_rpm_line.Visible = True
    
End If


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_SSDG)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub SSDG_Updated(Code As Integer)

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)
If Index = 0 Then
    Call setPCDI(PumpIO_LSS(4, 47))
ElseIf Index = 1 Then
         Call setPCDI(PumpIO_LSS(4, 8))    ' STOP SSDG Pre-lub pump
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 3/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


   
'   Note:  PumpIO(2=ON,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 47))
ElseIf Index = 1 Then
         Call setPCDI(PumpIO_LSS(2, 8))   ' START SSDG Pre-lub pump
End If

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Pdalt_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 372
    Alarm_set_point.Show

End Sub

Private Sub Pjack_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 375
    Alarm_set_point.Show

End Sub

Private Sub Plub_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 368
    Alarm_set_point.Show

End Sub

Private Sub psea_Click()

' 10/7/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 98
Alarm_set_point.Show

End Sub

Private Sub Psea_DA1_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 378
    Alarm_set_point.Show

End Sub

Private Sub Tair_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 382
    Alarm_set_point.Show

End Sub

Private Sub Texhaust_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 376
    Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim green_colour As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' === 26/03/03 ===
ALT_rpm(1).text = Format(LSS_reals(44 + 1), "##0") & " rpm"



' 4/4/00
Call paint_alarm(Me.Tlub(1), 374, 0, 0, 0)

' 10/7/00 achieves the same thing as above ....
'Call paint_elec_alarm(24, Tlub(1))

' ==== 10/7/00 =======

Call paint_elec_alarm(24, Tlub(1))
Call paint_elec_alarm(18, Plub(1))
Call paint_elec_alarm(30, Zsump(1))

Call paint_elec_alarm(32, Tair)

Call paint_elec_alarm(19, Tjack(1))
Call paint_elec_alarm(25, Pjack(1))

Call paint_elec_alarm(31, Zjcw)

Call paint_elec_alarm(26, Texhaust(1))

Call paint_elec_alarm(22, Pdalt(1))

Call paint_elec_alarm(28, Psea_DA1)

' 10/7/00
Call paint_elec_alarm(29, Tscav)


' 10/7/00
'Call paint_elec_alarm(98, psea)
Call paint_alarm(Psea, 98, 0, 0, 0)


' ==== 10/7/00 end ====




List1.clear
List1.AddItem "alarm(374) = " & alarm(374)
List1.AddItem "elec_alarm(24) = " & elec_alarm(24)



'          PC_REALS(45) = TD(1) * 60# / 6.284

' 23/11/99
' SSDG  LO and JCW Mechanical Pumps
If LSS_reals(45) > 720# Then
    green_colour = 255
    SSDG_LO_Pump.Colour3 = RGB(0, green_colour, 0)
ElseIf LSS_reals(45) < 180# Then
    'green_colour = 0
    SSDG_LO_Pump.Colour3 = &HFFFFFF        ' WHITE
Else
    green_colour = LSS_reals(45) * 255# / 720#
    SSDG_LO_Pump.Colour3 = RGB(0, green_colour, 0)
End If

'   23/11/99
    SSDG_JCW_Pump.Colour3 = SSDG_LO_Pump.Colour3


 '          PC_REALS(534) = tscav(1)

If SI_units Then
    PCA30.text = Format(LSS_reals(325) * Press_mult, "##0.0") & Press_chars     ' Bar
    Psea_DA1.text = Format(LSS_reals(733) * Press_mult, "##0.0") & Press_chars     ' Bar
    Psea.text = Format(LSS_reals(504) * Press_mult, "##0.0") & Press_chars     ' Bar
    Plub(1).text = Format(LSS_reals(508) * Press_mult, "##0.0") & Press_chars     ' Bar
    Pjack(1).text = Format(LSS_reals(704) * Press_mult, "##0.0") & Press_chars     ' Bar
    Pdalt(1).text = Format(LSS_reals(708) * Press_mult, "##0.0") & Press_chars     ' Bar

  '  FFO_recirc(1).text = Format(LSS_reals(234) * Flow_mult, "###0.000") & Flow_chars
Else
    PCA30.text = Format(LSS_reals(325) * Press_mult, "###0") & Press_chars    ' psi
    Psea_DA1.text = Format(LSS_reals(733) * Press_mult, "###0") & Press_chars    ' psi
    Psea.text = Format(LSS_reals(504) * Press_mult, "###0") & Press_chars    ' psi
    Plub(1).text = Format(LSS_reals(508) * Press_mult, "###0") & Press_chars    ' psi
    Pjack(1).text = Format(LSS_reals(704) * Press_mult, "###0") & Press_chars    ' psi
    Pdalt(1).text = Format(LSS_reals(708) * Press_mult, "###0") & Press_chars    ' psi
   
   ' FFO_recirc(1).text = Format(LSS_reals(234) * Flow_mult / 3600#, "###0.000") & " lb/s"
End If

'Call paint_alarm(sy201_LSS.pca30, 48, 0, 0, 0)


' 6/9/00  Point 8.40
'Zjcw.text = Format(LSS_reals(505), "##0") & " %"
Zjcw.text = Format(LSS_reals(505), "##0.0") & " %"
Z_JCW.LineWidth = LSS_reals(505)


' 21/9/99
Tair.text = Format(LSS_reals(531) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tscav.text = Format(LSS_reals(534) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

Tjack(1).text = Format(LSS_reals(512) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Texhaust(1).text = Format(LSS_reals(506) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tlub(1).text = Format(LSS_reals(706) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


Zsump(1).text = Format(LSS_reals(710), "##0") & " %"
  
Call update_pump_LSS(Me, 0, 47)      '  EMERGENCY COMPRESSOR


' 23/11/99
Call update_pump_LSS(Me, 1, 8)        ' SSDG Pre-Lub Pump


' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)
   
   Call update_vvalve_LSS(Me, 361, 0, -1)    ' 19/11/99

   Call update_vvalve_LSS(Me, 400, 1, -1)
   Call update_vvalve_LSS(Me, 401, 2, -1)
   
   Call update_vvalve_LSS(Me, 362, 3, -1)    ' 19/11/99
   
   Call update_vvalve_LSS(Me, 402, 4, -1)
   Call update_vvalve_LSS(Me, 403, 5, -1)
   
   Call update_vvalve_LSS(Me, 354, 6, -1)   ' 19/11/99
   
   Call update_vvalve_LSS(Me, 405, 10, -1)
   Call update_vvalve_LSS(Me, 406, 12, -1)
   
   Call update_vvalve_LSS(Me, 347, 19, -1)   ' 19/11/99
      
   Call update_vvalve_LSS(Me, 408, 20, -1)
   
   Call update_vvalve_LSS(Me, 345, 24, -1)   ' 19/11/99
   Call update_vvalve_LSS(Me, 348, 27, -1)   ' 19/11/99
   Call update_vvalve_LSS(Me, 350, 28, -1)   ' 19/11/99
   

   Call update_vvalve_LSS(Me, 349, 37, -1)   ' 19/11/99
   Call update_vvalve_LSS(Me, 351, 38, -1)   ' 19/11/99
  
 ' 10/12/00  NEW VALVE (see Bilge Xref V7 )
  Call update_vvalve_LSS(Me, 589, 31, -1)
  
   
End Sub

Private Sub Timer2_Timer()

End Sub

Private Sub Tjack_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 369
    Alarm_set_point.Show

End Sub

Private Sub Tlub_Click(Index As Integer)


' 4/4/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 374
Alarm_set_point.Show

End Sub

Private Sub Tscav_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 379
    Alarm_set_point.Show

End Sub

Private Sub valve_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 20/9/99  See Diesel Gen Valve X-Ref.xls




Select Case Index
    Case 0
        Call valve_pcdi_LSS(361) ' 19/11/99
    Case 1
        Call valve_pcdi_LSS(400)
    Case 2
        Call valve_pcdi_LSS(401)
    Case 3
        Call valve_pcdi_LSS(362)  ' 19/11/99
    Case 4
        Call valve_pcdi_LSS(402)
    Case 5
        Call valve_pcdi_LSS(403)
    Case 6
        Call valve_pcdi_LSS(354)  ' SW OVERBOARD Valve 19/11/99
    Case 10
        Call valve_pcdi_LSS(405)
    Case 12
        Call valve_pcdi_LSS(406)
  '  Case 18
  '      Call valve_pcdi_LSS(407)
    Case 19
        Call valve_pcdi_LSS(347) ' 19/11/99
    Case 20
        Call valve_pcdi_LSS(408)
    Case 24
        Call valve_pcdi_LSS(345) ' 19/11/99
    Case 27
        Call valve_pcdi_LSS(348) ' 19/11/99
    Case 28
        Call valve_pcdi_LSS(350) ' 19/11/99
        
    Case 31
        Call valve_pcdi_LSS(589) ' 10/12/00
        
        
    Case 37
        Call valve_pcdi_LSS(349) ' 19/11/99
    Case 38
        Call valve_pcdi_LSS(351) ' 19/11/99
        
End Select


End Sub

Private Sub Zjcw_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 381
    Alarm_set_point.Show

End Sub

Private Sub Zsump_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 380
    Alarm_set_point.Show

End Sub
Private Sub paint_elec_alarm(nalarm As Integer, obj As Object)
Dim alarmno As Integer

' SUB Copied from SSTG.frm
alarmno = nalarm + 350

'If elec_alarm(nalarm) Then
If alarm(alarmno) Then
    obj.BackColor = &HFF&            ' RED
    obj.ForeColor = &HFFFFFF         ' WHITE   ( 14/7/98 )
Else
    obj.BackColor = &HFF00&          ' GREEN
    obj.ForeColor = &H0&             ' BLACK
End If


End Sub
