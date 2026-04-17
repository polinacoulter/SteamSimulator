VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{A6D9C597-48C2-11D1-8915-000021240899}#2.0#0"; "Haven3WayValve v1.41.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form FO_Service 
   Caption         =   " Fuel Oil Service System"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   9720
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7455
      Left            =   0
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   13150
      Begin VB.TextBox QFO_heater 
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
         Left            =   3900
         TabIndex        =   86
         Top             =   7020
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   8520
         TabIndex        =   85
         Text            =   "85 to 450 psi"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox TFO_filter_B 
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
         Left            =   6960
         TabIndex        =   82
         Text            =   "xxx F"
         Top             =   4980
         Width           =   675
      End
      Begin VB.TextBox FFO_recirc 
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
         Left            =   4680
         TabIndex        =   81
         Text            =   "2.1 kg/s"
         Top             =   1440
         Width           =   915
      End
      Begin VB.TextBox FFO_recirc 
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
         Left            =   2440
         TabIndex        =   80
         Text            =   "2.1 kg/s"
         Top             =   1440
         Width           =   915
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   3780
         Max             =   1000
         TabIndex        =   79
         Top             =   3720
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox TFO_pumps 
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
         Left            =   3720
         TabIndex        =   78
         Text            =   "xxx F"
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox FFO_heater_byp 
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
         Left            =   4320
         TabIndex        =   77
         Text            =   "2.1 kg/s"
         Top             =   6280
         Width           =   1035
      End
      Begin VB.TextBox FFO_heater 
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
         Left            =   4320
         TabIndex        =   76
         Text            =   "2.1 kg/s"
         Top             =   5700
         Width           =   1035
      End
      Begin VB.TextBox FHO_sett_out 
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
         Left            =   2940
         TabIndex        =   75
         Text            =   "2.1 kg/s"
         Top             =   2880
         Width           =   1035
      End
      Begin VB.TextBox FHO_sett_out 
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
         Left            =   180
         TabIndex        =   74
         Text            =   "2.1 kg/s"
         Top             =   2400
         Width           =   1035
      End
      Begin VB.TextBox FDO_sett_out 
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
         Left            =   4980
         TabIndex        =   73
         Text            =   "2.1 kg/s"
         Top             =   2820
         Width           =   1035
      End
      Begin VB.TextBox F_burn 
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
         Left            =   8340
         TabIndex        =   72
         Text            =   "2.1 kg/s"
         Top             =   5460
         Width           =   1035
      End
      Begin VB.TextBox PFO_in 
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
         Left            =   1140
         TabIndex        =   71
         Text            =   "500 psi"
         Top             =   4740
         Width           =   735
      End
      Begin VB.TextBox PFO_out 
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
         Left            =   3720
         TabIndex        =   70
         Text            =   "500 psi"
         Top             =   4740
         Width           =   735
      End
      Begin Haven3WayValvePrj.Haven3WayValve Three_way_VV 
         Height          =   375
         Left            =   420
         TabIndex        =   61
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         OpenSideTri     =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.TextBox ZDO_sett 
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
         Left            =   7080
         TabIndex        =   58
         Text            =   "85 %"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TDO_sett 
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
         Left            =   7080
         TabIndex        =   57
         Text            =   "77 F"
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox THO_sett 
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
         Left            =   3780
         TabIndex        =   56
         Text            =   "approx 180 F"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox ZHO_sett 
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
         Left            =   3780
         TabIndex        =   55
         Text            =   "20 %"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "FO Transfer Pumps"
         Height          =   675
         Left            =   8220
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   480
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   300
         Top             =   5280
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
         Left            =   2880
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6960
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
         Left            =   2820
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3480
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
         Left            =   2400
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   6960
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
         Left            =   2340
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3480
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
         Left            =   1920
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6960
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
         Left            =   1860
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3480
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   1
         Left            =   8000
         TabIndex        =   40
         Top             =   1860
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
      Begin VB.CommandButton Command5 
         Caption         =   "FO Transfer Pumps"
         Height          =   675
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   36
         Left            =   7560
         TabIndex        =   37
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   6
      End
      Begin VB.TextBox PFO_heater 
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
         Left            =   6960
         TabIndex        =   36
         Text            =   "500 psi"
         Top             =   5400
         Width           =   675
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F.O. Recir"
         Height          =   675
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox TFO_heater 
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
         Left            =   6040
         TabIndex        =   32
         Text            =   "xxx F"
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox THO_sett 
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
         Index           =   2
         Left            =   1500
         TabIndex        =   31
         Text            =   "180 F"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox ZHO_sett 
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
         Index           =   2
         Left            =   1500
         TabIndex        =   30
         Text            =   "20 %"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Steam Supply"
         Height          =   495
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton Command1 
         Caption         =   "To Burners"
         Height          =   435
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5940
         Width           =   915
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
         Height          =   315
         Left            =   5880
         TabIndex        =   21
         Text            =   "T"
         Top             =   4770
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
         Height          =   255
         Left            =   8160
         TabIndex        =   20
         Text            =   "P"
         Top             =   5040
         Width           =   195
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   0
         Left            =   6240
         TabIndex        =   0
         Top             =   1710
         Visible         =   0   'False
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
         Height          =   1275
         Index           =   2
         Left            =   6780
         TabIndex        =   1
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2249
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   80
         OrientVertical  =   0   'False
         ColourClosed    =   33023
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   1275
         Index           =   3
         Left            =   1200
         TabIndex        =   2
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2249
         TopColour       =   8421504
         BottomColour    =   8421504
         LineWidth       =   3
         OrientVertical  =   0   'False
         ColourClosed    =   16576
         BackColour      =   12632256
         ValveType       =   7
         ColourOpen      =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   5
         Left            =   840
         TabIndex        =   4
         Top             =   1620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         BottomLineColour=   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   6
         Left            =   840
         TabIndex        =   5
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
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   7
         Left            =   3045
         TabIndex        =   6
         Top             =   1620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   8
         Left            =   3045
         TabIndex        =   9
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
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   10
         Left            =   2340
         TabIndex        =   10
         Top             =   3660
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
         Height          =   435
         Index           =   11
         Left            =   1620
         TabIndex        =   11
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   12
         Left            =   3120
         TabIndex        =   12
         Top             =   4080
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
         Index           =   13
         Left            =   1680
         TabIndex        =   13
         Top             =   5820
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
         Index           =   14
         Left            =   3120
         TabIndex        =   14
         Top             =   5820
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
         Index           =   16
         Left            =   2340
         TabIndex        =   16
         Top             =   6380
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
         Height          =   435
         Index           =   17
         Left            =   1920
         TabIndex        =   17
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   18
         Left            =   2640
         TabIndex        =   18
         Top             =   4800
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
         Index           =   21
         Left            =   5220
         TabIndex        =   23
         Top             =   3660
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
         Index           =   22
         Left            =   5940
         TabIndex        =   24
         Top             =   3660
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
         Index           =   24
         Left            =   5580
         TabIndex        =   26
         Top             =   4160
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
         Index           =   25
         Left            =   2280
         TabIndex        =   28
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
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   26
         Left            =   5400
         TabIndex        =   29
         Top             =   5940
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
         Index           =   9
         Left            =   2400
         TabIndex        =   33
         Top             =   960
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
         Index           =   19
         Left            =   3660
         TabIndex        =   34
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
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   34
         Left            =   2220
         TabIndex        =   38
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   6
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   35
         Left            =   4380
         TabIndex        =   39
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   6
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   20
         Left            =   8220
         TabIndex        =   22
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   15
         Left            =   2280
         TabIndex        =   15
         Top             =   4800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   32768
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   28
         Left            =   7080
         TabIndex        =   42
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   23
         Left            =   1500
         TabIndex        =   53
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
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   1275
         Index           =   4
         Left            =   3420
         TabIndex        =   60
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2249
         TopColour       =   8421504
         BottomColour    =   8421504
         LineWidth       =   3
         OrientVertical  =   0   'False
         ColourClosed    =   16512
         BackColour      =   12632256
         ValveType       =   7
         ColourOpen      =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   29
         Left            =   3540
         TabIndex        =   62
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   30
         Left            =   1320
         TabIndex        =   63
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   31
         Left            =   1080
         TabIndex        =   64
         Top             =   3000
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
         Index           =   32
         Left            =   6180
         TabIndex        =   65
         Top             =   5040
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
         Index           =   33
         Left            =   4260
         TabIndex        =   66
         Top             =   5040
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
         Index           =   0
         Left            =   2280
         TabIndex        =   68
         Top             =   4140
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   69
         Top             =   5880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   27
         Left            =   4620
         TabIndex        =   83
         Top             =   960
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
         Height          =   375
         Index           =   421
         Left            =   6980
         TabIndex        =   94
         Top             =   2280
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
         Height          =   375
         Index           =   192
         Left            =   40
         TabIndex        =   95
         Top             =   3120
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
         Height          =   375
         Index           =   191
         Left            =   420
         TabIndex        =   96
         Top             =   3840
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Make invisible after testing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5640
         TabIndex        =   98
         Top             =   1020
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Remember to link to the NEW screen after testing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   6570
         TabIndex        =   97
         Top             =   6390
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   6960
         Picture         =   "FO_Service.frx":0000
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   255
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   7920
         Picture         =   "FO_Service.frx":030A
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   28
         X1              =   8040
         X2              =   7080
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   27
         X1              =   8040
         X2              =   8040
         Y1              =   5940
         Y2              =   5460
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   26
         X1              =   8040
         X2              =   8040
         Y1              =   5220
         Y2              =   3720
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   7920
         X2              =   7920
         Y1              =   5160
         Y2              =   5580
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   7740
         X2              =   7740
         Y1              =   5160
         Y2              =   5580
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   25
         X1              =   7080
         X2              =   7080
         Y1              =   3720
         Y2              =   2520
      End
      Begin VB.Label HODO_label 
         Caption         =   "HO/DO"
         Height          =   255
         Left            =   240
         TabIndex        =   93
         Top             =   3500
         Width           =   615
      End
      Begin VB.Label HO_DO_label 
         Caption         =   "HO"
         Height          =   255
         Left            =   300
         TabIndex        =   92
         Top             =   4320
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Reg_FO_htr_pos 
         Alignment       =   1  'Right Justify
         Caption         =   "xxx %"
         Height          =   255
         Left            =   5100
         TabIndex        =   91
         Top             =   3480
         Width           =   555
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   6720
         X2              =   6000
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   4740
         Picture         =   "FO_Service.frx":0614
         Stretch         =   -1  'True
         Top             =   4620
         Width           =   255
      End
      Begin VB.Image DO_arrow 
         Height          =   255
         Index           =   1
         Left            =   1560
         Picture         =   "FO_Service.frx":091E
         Stretch         =   -1  'True
         Top             =   3180
         Width           =   255
      End
      Begin VB.Image Arrow_left 
         Height          =   255
         Left            =   780
         Picture         =   "FO_Service.frx":0C28
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   0
         Picture         =   "FO_Service.frx":0F32
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Line DO_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   240
         X2              =   6060
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   20
         X1              =   4140
         X2              =   6840
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   3180
         Picture         =   "FO_Service.frx":123C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   1080
         Picture         =   "FO_Service.frx":1546
         Stretch         =   -1  'True
         Top             =   420
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   24
         X1              =   2640
         X2              =   2760
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   23
         X1              =   2640
         X2              =   2640
         Y1              =   540
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   21
         X1              =   960
         X2              =   2640
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   3
         X1              =   1080
         X2              =   3660
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Shape DO_oil 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   6780
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Shape HO_DO_oil 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   6540
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Shape HO_oil 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   6300
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   8100
         Picture         =   "FO_Service.frx":1850
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   20
         X1              =   6960
         X2              =   6960
         Y1              =   1200
         Y2              =   780
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   7740
         Picture         =   "FO_Service.frx":1B5A
         Stretch         =   -1  'True
         Top             =   660
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   3240
         Picture         =   "FO_Service.frx":1E64
         Stretch         =   -1  'True
         Top             =   780
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "FO Supply Pumps"
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
         Left            =   3000
         TabIndex        =   52
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "# 2"
         Height          =   315
         Left            =   1560
         TabIndex        =   51
         Top             =   7020
         Width           =   405
      End
      Begin VB.Label Label10 
         Caption         =   "# 1"
         Height          =   315
         Left            =   1500
         TabIndex        =   50
         Top             =   3540
         Width           =   405
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   23
         X1              =   8220
         X2              =   8220
         Y1              =   5940
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   21
         X1              =   3180
         X2              =   3180
         Y1              =   5100
         Y2              =   5580
      End
      Begin VB.Image DO_arrow 
         Height          =   255
         Index           =   0
         Left            =   5940
         Picture         =   "FO_Service.frx":216E
         Stretch         =   -1  'True
         Top             =   2460
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   22
         X1              =   6960
         X2              =   8100
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   5940
         X2              =   5700
         Y1              =   5340
         Y2              =   5520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   19
         X1              =   2040
         X2              =   2040
         Y1              =   1260
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   18
         X1              =   4140
         X2              =   4140
         Y1              =   1260
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   17
         X1              =   3000
         X2              =   4140
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   15
         X1              =   5220
         X2              =   4920
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   14
         X1              =   5220
         X2              =   5220
         Y1              =   1260
         Y2              =   480
      End
      Begin VB.Label Label7 
         Caption         =   "To SSDG"
         Height          =   255
         Left            =   8220
         TabIndex        =   41
         Tag             =   "Tossdg"
         Top             =   2940
         Width           =   795
      End
      Begin VB.Line DO_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   6060
         X2              =   6060
         Y1              =   2160
         Y2              =   3300
      End
      Begin VB.Line DO_line 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   6060
         X2              =   6900
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   5760
         X2              =   5760
         Y1              =   3480
         Y2              =   3660
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   3660
         X2              =   3660
         Y1              =   2760
         Y2              =   2400
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   1920
         X2              =   1920
         Y1              =   2760
         Y2              =   2340
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   1440
         X2              =   1440
         Y1              =   2760
         Y2              =   2340
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   4140
         X2              =   4140
         Y1              =   2400
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   12
         X1              =   7620
         X2              =   7620
         Y1              =   2940
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   8
         X1              =   4440
         X2              =   4440
         Y1              =   2880
         Y2              =   2640
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   7
         X1              =   2280
         X2              =   2280
         Y1              =   2820
         Y2              =   2400
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   300
         Picture         =   "FO_Service.frx":2478
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   8460
         Picture         =   "FO_Service.frx":2782
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   5340
         Picture         =   "FO_Service.frx":2A8C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   6780
         Picture         =   "FO_Service.frx":2D96
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   0
         Picture         =   "FO_Service.frx":30A0
         Stretch         =   -1  'True
         Top             =   3660
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   5820
         Picture         =   "FO_Service.frx":33AA
         Stretch         =   -1  'True
         Top             =   5820
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   1080
         Picture         =   "FO_Service.frx":36B4
         Stretch         =   -1  'True
         Top             =   780
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   7
         X1              =   2880
         X2              =   2880
         Y1              =   3960
         Y2              =   4380
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   11
         X1              =   2880
         X2              =   2580
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   10
         X1              =   2880
         X2              =   2880
         Y1              =   1260
         Y2              =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   9
         X1              =   2880
         X2              =   5700
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   30
         X1              =   6900
         X2              =   6420
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   22
         X1              =   6840
         X2              =   6840
         Y1              =   6240
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   19
         X1              =   4140
         X2              =   4140
         Y1              =   6240
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   18
         X1              =   1740
         X2              =   1740
         Y1              =   5100
         Y2              =   5580
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   12
         X1              =   1740
         X2              =   3180
         Y1              =   5580
         Y2              =   5580
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   6720
         X2              =   6720
         Y1              =   4560
         Y2              =   3480
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   6720
         X2              =   6720
         Y1              =   4920
         Y2              =   4740
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   4860
         X2              =   4860
         Y1              =   3960
         Y2              =   5520
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   6420
         X2              =   4860
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   6420
         X2              =   6420
         Y1              =   4440
         Y2              =   3960
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   780
         X2              =   780
         Y1              =   4920
         Y2              =   5340
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   6720
         X2              =   5740
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   8220
         X2              =   8220
         Y1              =   5160
         Y2              =   4620
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   8220
         X2              =   5040
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   8
         X1              =   5940
         X2              =   5940
         Y1              =   5340
         Y2              =   5820
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   5700
         X2              =   5520
         Y1              =   5520
         Y2              =   5340
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   5520
         X2              =   5280
         Y1              =   5340
         Y2              =   5520
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   5280
         X2              =   5100
         Y1              =   5520
         Y2              =   5340
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   5100
         X2              =   4860
         Y1              =   5340
         Y2              =   5520
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7740
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   195
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7740
         Shape           =   3  'Circle
         Top             =   5160
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FO Heater"
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
         Left            =   5040
         TabIndex        =   19
         Top             =   4800
         Width           =   975
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   15
         X1              =   3660
         X2              =   1200
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   14
         X1              =   3660
         X2              =   3660
         Y1              =   6120
         Y2              =   4380
      End
      Begin VB.Line Fuel_line 
         BorderStyle     =   3  'Dot
         Index           =   13
         X1              =   4620
         X2              =   2460
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line control__line 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   2520
         X2              =   2520
         Y1              =   5160
         Y2              =   4740
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   11
         X1              =   2160
         X2              =   2160
         Y1              =   6120
         Y2              =   6660
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   10
         X1              =   2880
         X2              =   2880
         Y1              =   6120
         Y2              =   6660
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   9
         X1              =   2160
         X2              =   2880
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   8
         X1              =   2160
         X2              =   2160
         Y1              =   4380
         Y2              =   3960
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   6
         X1              =   2160
         X2              =   2880
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   5
         X1              =   960
         X2              =   1200
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   4
         X1              =   1080
         X2              =   1080
         Y1              =   6120
         Y2              =   4380
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   2
         X1              =   1080
         X2              =   3660
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   5
         X1              =   2880
         X2              =   2880
         Y1              =   2940
         Y2              =   1920
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   960
         X2              =   960
         Y1              =   4920
         Y2              =   5340
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   780
         Shape           =   3  'Circle
         Top             =   5100
         Width           =   195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   6
         X1              =   120
         X2              =   2880
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   4
         X1              =   2880
         X2              =   4560
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   3
         X1              =   2880
         X2              =   3660
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "Stbd HO Settling Tank"
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
         Left            =   3900
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Port HO Settling Tank"
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
         Left            =   1080
         TabIndex        =   7
         Top             =   120
         Width           =   1875
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   2
         X1              =   120
         X2              =   840
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   780
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   1020
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "DO Settling Tank"
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
         TabIndex        =   3
         Top             =   180
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   2
         X1              =   8580
         X2              =   8580
         Y1              =   2160
         Y2              =   2820
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   7980
         X2              =   8580
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   780
         Shape           =   3  'Circle
         Top             =   4920
         Width           =   195
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   3
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4740
         Top             =   5100
         Width           =   1335
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   17
         X1              =   5940
         X2              =   8220
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   120
         Y1              =   1920
         Y2              =   5100
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   16
         X1              =   960
         X2              =   2040
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   13
         X1              =   3660
         X2              =   4800
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Label PFuel_Control_pos 
         Alignment       =   2  'Center
         Caption         =   "xxx.x % "
         Height          =   195
         Left            =   1860
         TabIndex        =   84
         Top             =   4620
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "DO NOT DELETE"
         Height          =   735
         Left            =   840
         TabIndex        =   67
         Top             =   3420
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OLE Reg_FO_htr 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   5580
         OleObjectBlob   =   "FO_Service.frx":39BE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   89
         Top             =   3540
         Width           =   375
      End
      Begin VB.OLE Reg_FO_htr 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   5580
         OleObjectBlob   =   "FO_Service.frx":BFD6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   88
         Top             =   3560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_FO_htr 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   5580
         OleObjectBlob   =   "FO_Service.frx":149EE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   87
         Top             =   3540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_FO_htr 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   5580
         OleObjectBlob   =   "FO_Service.frx":1CC06
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   90
         Top             =   3560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Stem_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   6420
         X2              =   4860
         Y1              =   3960
         Y2              =   3960
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   59
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "FO_Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()


' 20/07/09
'Burner_FO_system.Show
Burner_FO_system_NEW.Show

End Sub

Private Sub Command2_Click()


' 4/8/00
Stm_Stm_Generator.Show


End Sub

Private Sub Command3_Click()

' 16/07/09
'Burner_FO_system.Show
Burner_FO_system_NEW.Show

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

Bunkers.Show

End Sub

Private Sub Command6_Click()


End Sub

Private Sub Command7_Click()

Bunkers.Show

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idFO_serv)

valve(2).LineWidth = 0
valve(3).LineWidth = 0
valve(4).LineWidth = 0


' === 16/12/02 ===
FDO_sett_out.Visible = False
DO_arrow(0).Visible = False
DO_arrow(1).Visible = False

' 16/07/09 testing
'valve(0).Visible = False

valve(31).Visible = False
HODO_label.Visible = False
DO_line(0).Visible = False
DO_line(1).Visible = False
DO_line(2).Visible = False
Three_way_VV.Visible = False



' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    ONpump(0).Visible = False
    OFFpump(0).Visible = False
    sbyPUMP(0).Visible = False
    ONpump(1).Visible = False
    OFFpump(1).Visible = False
    sbyPUMP(1).Visible = False
    
    ' 21/9/00
    PFuel_Control_pos.Visible = False
    
Else
    ONpump(0).Visible = True
    OFFpump(0).Visible = True
    sbyPUMP(0).Visible = True
    ONpump(1).Visible = True
    OFFpump(1).Visible = True
    sbyPUMP(1).Visible = True
    
    ' 21/9/00
    PFuel_Control_pos.Visible = True
    
    
End If


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idFO_serv)
    HavenStretch1.Visible = True
End If
   
End Sub

Private Sub HScroll1_Change()

    Call set_VSCROLL(7, HScroll1.Value) ' Heat in kW

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
' 15/2/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
   
   
'   Note:  PumpIO(4=OFF,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(4, 29))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(4, 30))
'End If


'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 3))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 15))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 15/2/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
  
   
   
'   Note:  PumpIO_LSS(2=ON,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(2, 29))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(2, 30))
'End If

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 3))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 15))
End If


End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
' 15/2/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
  
   
'   Note:  PumpIO(5=Stand-bye,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(5, 29))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(5, 30))
'End If

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 3))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(5, 15))
End If


End Sub

Private Sub THO_sett_Click(Index As Integer)

    Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_HI = 28  ' Tank # 1 STBD
    nAlarm_LO = 36  ' Tank # 1 STBD
Else
    nAlarm_HI = 29  ' Tank # 2 PORT
    nAlarm_LO = 37  ' Tank # 2 PORT
End If

    Alarm_set_point.Show

End Sub

Private Sub Three_way_VV_Click()

Control_PC_check
If This_Control_status = False Then Exit Sub

' 20/9/99
'Call valve_pcdi_LSS(197)
 Call valve_pcdi_LSS(191)
 

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim green_col  As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' 18/04/01
If VB_steam_logicals(14) Then
  HO_DO_label.Caption = "HO"
Else
  HO_DO_label.Caption = "DO"
End If


' 12/8/00
Call paint_alarm(Me.THO_sett(1), 28, 36, 0, 0)
Call paint_alarm(Me.THO_sett(2), 29, 37, 0, 0)

Call paint_alarm(Me.ZHO_sett(1), 44, 0, 0, 0)
Call paint_alarm(Me.ZHO_sett(2), 45, 0, 0, 0)

' 12/8/00
'Call paint_alarm(Me.ZDO_sett, 43, 0, 0, 0)
Call paint_alarm(Me.ZDO_sett, 43, 47, 0, 0)



' === 11/8/00 ===
'    PFuel_control_pos 0 to 100 %
'    PC_REALS(397) = PFuel_control_pos
PFuel_Control_pos.Caption = Format(LSS_reals(397), "##0.0") & " %"
green_col = LSS_reals(397) * 2.55
valve(15).ColourOpen = RGB(0, green_col, 0)


' 9/8/00
'' PC_REALS(144) = FO_heater_reg_pos * 100#
'green_col = LSS_reals(144) * 2.55
'valve(20).ColourOpen = RGB(0, green_col, 0)

Reg_FO_htr_pos.Caption = Format(LSS_reals(144), "###0") & " %"
Call update_regulator(144, Reg_FO_htr(0), Reg_FO_htr(1), Reg_FO_htr(2), Reg_FO_htr(3))


' 9/8/00
'  PC_REALS(152) = QFO_heater
QFO_heater.text = Format(LSS_reals(152), "###0")



'Tsteam14.Text = Format(LSS_reals(74) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TFO_pumps.text = Format(LSS_reals(195) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TFO_heater.text = Format(LSS_reals(196) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' 7/8/00  (See OV004.F90 )
'TFO_filter_B.text = Format(LSS_reals(200) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TFO_filter_B.text = Format(LSS_reals(199) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


If SI_units Then
    PFO_in.text = Format(LSS_reals(220) * Press_mult, "##0.00") & Press_chars     ' Bar
    PFO_out.text = Format(LSS_reals(221) * Press_mult, "##0.0") & Press_chars     ' Bar
    PFO_heater.text = Format(LSS_reals(227) * Press_mult, "##0.0") & Press_chars     ' Bar

    ' 7/8/00
    PFO_A.text = Format(LSS_reals(230) * Press_mult, "##0.0") & Press_chars     ' Bar

    FFO_recirc(1).text = Format(LSS_reals(234) * Flow_mult, "###0.000") & Flow_chars
    FFO_recirc(2).text = Format(LSS_reals(192) * Flow_mult, "###0.000") & Flow_chars
    
    FFO_heater.text = Format(LSS_reals(197) * Flow_mult, "###0.000") & Flow_chars
    FFO_heater_byp.text = Format(LSS_reals(198) * Flow_mult, "###0.000") & Flow_chars

    F_burn.text = Format(LSS_reals(222) * Flow_mult, "###0.000") & Flow_chars
    FDO_sett_out.text = Format(LSS_reals(225) * Flow_mult, "###0.000") & Flow_chars
    FHO_sett_out(1).text = Format(LSS_reals(223) * Flow_mult, "###0.000") & Flow_chars
    FHO_sett_out(2).text = Format(LSS_reals(224) * Flow_mult, "###0.000") & Flow_chars
Else
    ' === 22/01/02 15.045 inches of Hg for Maryland ===
   'If TARGET_this_sim = MARYLAND Then
    If TARGET_this_sim = MARYLAND Or TARGET_this_sim = TEXAS Then
        PFO_in.text = Format(LSS_reals(220) * Press_mult * 2.036, "###0.0") & " inHg" ' inHg
    Else
        PFO_in.text = Format(LSS_reals(220) * Press_mult, "###0.0") & Press_chars    ' psi
    End If
    
    
    PFO_out.text = Format(LSS_reals(221) * Press_mult, "###0") & Press_chars     ' psi
    PFO_heater.text = Format(LSS_reals(227) * Press_mult, "###0") & Press_chars    ' psi
    
    ' 7/8/00
    PFO_A.text = Format(LSS_reals(230) * Press_mult, "##0") & Press_chars     ' Bar

    FFO_recirc(1).text = Format(LSS_reals(234) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_recirc(2).text = Format(LSS_reals(192) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
    FFO_heater.text = Format(LSS_reals(197) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_heater_byp.text = Format(LSS_reals(198) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
    F_burn.text = Format(LSS_reals(222) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FDO_sett_out.text = Format(LSS_reals(225) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FHO_sett_out(1).text = Format(LSS_reals(223) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FHO_sett_out(2).text = Format(LSS_reals(224) * Flow_mult / 3600#, "###0.000") & " lb/s"
End If


' Stbd HO Settling Tank
If VB_steam_integers(28) = 0 Then
    valve(4).ColourClosed = DO_oil.FillColor
ElseIf VB_steam_integers(28) = 1 Then
    valve(4).ColourClosed = HO_oil.FillColor
Else
    valve(4).ColourClosed = HO_DO_oil.FillColor
End If
    
' Port HO Settling Tank
If VB_steam_integers(29) = 0 Then
    valve(3).ColourClosed = DO_oil.FillColor
ElseIf VB_steam_integers(29) = 1 Then
    valve(3).ColourClosed = HO_oil.FillColor
Else
    valve(3).ColourClosed = HO_DO_oil.FillColor
End If
    
' DO Settling Tank
If VB_steam_integers(32) = 0 Then
    valve(2).ColourClosed = DO_oil.FillColor
ElseIf VB_steam_integers(32) = 1 Then
    valve(2).ColourClosed = HO_oil.FillColor
Else
    valve(2).ColourClosed = HO_DO_oil.FillColor
End If

THO_sett(1).text = Format(LSS_reals(258) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
THO_sett(2).text = Format(LSS_reals(259) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

TDO_sett.text = Format(LSS_reals(260) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

ZHO_sett(1).text = Format(LSS_reals(249), "##0.0") & " %"
valve(4).LineWidth = LSS_reals(249)

ZHO_sett(2).text = Format(LSS_reals(250), "##0.0") & " %"
valve(3).LineWidth = LSS_reals(250)


' DO Settling Tank
ZDO_sett.text = Format(LSS_reals(251), "##0.0") & " %"
valve(2).LineWidth = LSS_reals(251)

' Note: update_pump_LSS(syxxx,index,nPump)
' 20/9/99
'Call update_pump_LSS(FO_Service, 0, 29)      ' FO Supply Pump # 1
'Call update_pump_LSS(FO_Service, 1, 30)      ' FO Supply Pump # 2
Call update_pump_LSS(FO_Service, 0, 3)      ' FO Supply Pump # 1
Call update_pump_LSS(FO_Service, 1, 15)      ' FO Supply Pump # 2


' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)
For i = 0 To 1
   Call update_vvalve_LSS(Me, 160 + i, i, -1)
Next i

' 13/8/00  SKIP valve 10
For i = 5 To 9
   Call update_vvalve_LSS(Me, 160 + i, i, -1)
Next i
For i = 11 To 14
   Call update_vvalve_LSS(Me, 160 + i, i, -1)
Next i

' 4/8/00
'For i = 16 To 35
'   Call update_vvalve_LSS(Me, 160 + i, i, -1)
'Next i
' 13/8/00  SKIP valve 16
'For i = 16 To 19

' === 05/03/01 ===
'For i = 17 To 19
For i = 17 To 18
   Call update_vvalve_LSS(Me, 160 + i, i, -1)
Next i

'  STBD HO Settler valve
   Call update_vvalve_LSS(Me, 375, 19, -1)

' === 05/03/01 ===
'For i = 21 To 35
For i = 21 To 22
   Call update_vvalve_LSS(Me, 160 + i, i, -1)
Next i

'  PORT HO Settler valve
   Call update_vvalve_LSS(Me, 376, 23, -1)

For i = 24 To 35
   Call update_vvalve_LSS(Me, 160 + i, i, -1)
Next i

' === 20/08/02 ===
   Call update_vvalve_LSS(Me, 421, 421, -1)
   
   

'''''''''''''''''''''''
' 13/8/00  Pressure relief valve
'        equivalence(STEAM_logicals(122),L_suc_FO)
'        logical*1                       L_suc_FO

' 24/8/00
'If vv_status_LSS(171) > 0 And vv_status_LSS(172) = 0 And pp_status_LSS(3) > 1 Then
If vv_status_LSS(171) > 0 And pp_status_LSS(3) > 1 Then
    If VB_steam_logicals(122) Then
        ' CLOSE relief V/V #1
        valve(10).OpenBtmRight = False
        valve(10).OpenTopLeft = False
    Else
    
     '   PFO_out.text = Format(LSS_reals(221) * Press_mult, "##0.0") & Press_chars     ' Bar

      ' 24/8/00
      ' 15/9/00  8.107 Fuel pump gets up to 38.0 Bar MAX = 551 psi POP relief at 550 psi
      ' If vv_status_LSS(172) = 0 Or LSS_reals(221) > 520# / 14.5 Then
        If vv_status_LSS(172) = 0 Or LSS_reals(221) > 550# / 14.5 Then
            ' OPEN relief V/V #1
            valve(10).OpenBtmRight = True
            valve(10).OpenTopLeft = True
        Else
            valve(10).OpenBtmRight = False
            valve(10).OpenTopLeft = False
        End If
    End If
Else
    valve(10).OpenBtmRight = False
    valve(10).OpenTopLeft = False
End If

' 24/8/00
'If vv_status_LSS(173) > 0 And vv_status_LSS(174) = 0 And pp_status_LSS(15) > 1 Then
If vv_status_LSS(173) > 0 And pp_status_LSS(15) > 1 Then
    If VB_steam_logicals(122) Then
        ' CLOSE relief V/V #2
        valve(16).OpenBtmRight = False
        valve(16).OpenTopLeft = False
    Else
    
        ' 24/8/00
      ' 15/9/00  8.107 Fuel pump gets up to 38.0 Bar MAX = 551 psi POP relief at 550 psi
      ' If vv_status_LSS(174) = 0 Or LSS_reals(221) > 520# / 14.5 Then
        If vv_status_LSS(174) = 0 Or LSS_reals(221) > 550# / 14.5 Then
            ' OPEN relief V/V #2
            valve(16).OpenBtmRight = True
            valve(16).OpenTopLeft = True
        Else
            valve(16).OpenBtmRight = False
            valve(16).OpenTopLeft = False
        End If
          
    End If
Else
    valve(16).OpenBtmRight = False
    valve(16).OpenTopLeft = False
End If


'''''''''''''''''''''''
     
' Three Way Valve
' 20/9/99  WAS valve 197 but is now valve 191...


' === 16/12/02 ===
'  Call update_vvalve_LSS(FO_Service, 191, 31, -1)
   Call update_vvalve_LSS(FO_Service, 191, 191, 192)
   
If (vv_status_LSS(191) = 0) Then
    Three_way_VV.OpenBtmTri = True
    Three_way_VV.OpenSideTri = False
ElseIf (vv_status_LSS(191) = 3) Then
    Three_way_VV.OpenBtmTri = False
    Three_way_VV.OpenSideTri = True
Else
    Three_way_VV.OpenBtmTri = Not Three_way_VV.OpenBtmTri
    Three_way_VV.OpenSideTri = Not Three_way_VV.OpenSideTri
End If
 
End Sub

Private Sub valve_Click(Index As Integer)

  Control_PC_check
  If This_Control_status = False Then Exit Sub
   
   
   
 ' 14/2/00 Block pump valve clicks on ACT Terminals
If Index > 10 And Index < 15 Then
  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
End If
   
   
 ' 29/4/00 Block relief valve clicks on ACT Terminals
If Index = 10 Or Index = 16 Then

 ' 13/8/00 Block relief valve clicks on ALL Terminals
  'If Not I_AM_INSTRUCTOR And Not PTT Then
  '    ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  'End If
  
End If
   
   
' === 16/12/02 ===
If Index = 192 Then
    Call valve_pcdi_LSS(191)
      Exit Sub
End If
   
   
   
' === 20/08/02 ===
If Index = 421 Then
    ' DO Sett tank to burners
    Call valve_pcdi_LSS(421)
      Exit Sub
End If


   
' 20/9/99  See Fuel_Oil_Service_System_Valve Xref.xls
If (Index < 2 Or Index > 4) Then
  If (Index <> 15) Then
  
    ' === 06/03/01 Point 10.104 ===
    If Index = 19 Then
    
            ' STBD HO Settler tank filling valve
             Call valve_pcdi_LSS(375)
             
    ElseIf Index = 23 Then
    
            ' PORT HO Settler tank filling valve
             Call valve_pcdi_LSS(376)
   
    Else
      
      
          ' 4/8/00
        If (Index <> 20) Then
            Call valve_pcdi_LSS(Index + 160)
        End If
        
    End If
  
  End If
End If


' OLD LSS Valves...
 
'  Valves
'Select Case Index
'  Case 0
'   Call valve_pcdi_LSS(195)
'  Case 1
'    Call valve_pcdi_LSS(173)        ' Diesel Oil to ALTERNATORS
'  Case 5
'    Call valve_pcdi_LSS(174)
'  Case 6
'    Call valve_pcdi_LSS(175)
'  Case 7
'    Call valve_pcdi_LSS(176)
'  Case 8
'    Call valve_pcdi_LSS(177)
'  Case 9
'    Call valve_pcdi_LSS(178)
'  Case 10
'    Call valve_pcdi_LSS(189)        ' Relief Valve (Remove?)
'  Case 11
'   'Call valve_pcdi_LSS(180)
'    Call valve_pcdi_LSS(225)
'  Case 12
'    Call valve_pcdi_LSS(182)
'  Case 13
'    Call valve_pcdi_LSS(183)
'  Case 14
'    Call valve_pcdi_LSS(184)
'  Case 16
'    Call valve_pcdi_LSS(190)        ' Relief Valve (Remove?)
'  Case 17
'    Call valve_pcdi_LSS(186)
'  Case 18
'    Call valve_pcdi_LSS(187)
'
'  Case 19
'    Call valve_pcdi_LSS(154)
'
' Case 21
'    Call valve_pcdi_LSS(192)
'  Case 22
'    Call valve_pcdi_LSS(193)
'
'  Case 23
'    Call valve_pcdi_LSS(155)
'
'  Case 24
'    Call valve_pcdi_LSS(194)
'  Case 25
'    Call valve_pcdi_LSS(188)
'  Case 26
'    Call valve_pcdi_LSS(191)
'  Case 27
'   'Call valve_pcdi_LSS(179)
'   ' Call valve_pcdi_LSS(224)
'    Call valve_pcdi_LSS(226)
'
'  Case 28
'    Call valve_pcdi_LSS(153)
'
'  Case 29
'    Call valve_pcdi_LSS(205)
'  Case 30
'    Call valve_pcdi_LSS(203)
'  'Case 31   ( See Three_way_VV )
'    'Call valve_pcdi_LSS(197)
'  Case 32
'    Call valve_pcdi_LSS(198)
'  Case 33
'    Call valve_pcdi_LSS(196)
'
'  Case 34
'    Call valve_pcdi_LSS(221)
'  Case 35
'    Call valve_pcdi_LSS(222)
'  Case 36
'    Call valve_pcdi_LSS(223)
'
'End Select

End Sub
Sub update_regulator(nReals As Integer, obj_0 As Object, obj_1 As Object, obj_2 As Object, obj_3 As Object)

' 6/11/99
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

Private Sub ZDO_sett_Click()

    Unload Alarm_set_point   ' Do this first

    nAlarm_HI = 43  ' DO Settling Tank
    
    ' 12/8/00
    nAlarm_LO = 47  ' DO Settling Tank

    Alarm_set_point.Show

End Sub

Private Sub ZHO_sett_Click(Index As Integer)

' 12/8/00
    Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_LO = 44  ' Tank # 1 STBD
Else
    nAlarm_LO = 45  ' Tank # 2 PORT
End If

    Alarm_set_point.Show

End Sub
