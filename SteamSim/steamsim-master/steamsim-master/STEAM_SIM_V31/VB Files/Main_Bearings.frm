VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{479F4823-E7F1-11D0-901E-0020AF7543C2}#5.0#0"; "fxrtx50g.ocx"
Begin VB.Form Main_Bearings 
   Caption         =   "Main Bearings"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10425
   ScaleWidth      =   13005
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9915
      Left            =   0
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   17489
      Begin VB.TextBox T_foot_txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2940
         TabIndex        =   115
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1455
         Left            =   360
         Max             =   160
         TabIndex        =   114
         Top             =   2760
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame2 
         Caption         =   "Turning Gear"
         Height          =   1275
         Left            =   4920
         TabIndex        =   100
         Top             =   8160
         Width           =   1215
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
            Left            =   660
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   540
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "START"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "STOP"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   780
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin PumpPrj.Pump HavenPump1 
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   106
            Top             =   840
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            CentrifugalPump =   0   'False
            OrientLeft      =   0   'False
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Turning Gear"
         Height          =   1215
         Left            =   2700
         TabIndex        =   95
         Top             =   8160
         Width           =   2055
         Begin VB.CommandButton TURN_IN 
            Caption         =   "IN"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            TabIndex        =   97
            Top             =   360
            Width           =   555
         End
         Begin VB.CommandButton Turn_OUT 
            Caption         =   "OUT"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            TabIndex        =   96
            Top             =   780
            Width           =   555
         End
         Begin VB.Image Turning_gear_OUT 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   1
            Left            =   180
            Picture         =   "Main_Bearings.frx":0000
            Stretch         =   -1  'True
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IN"
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
            Left            =   720
            MousePointer    =   1  'Arrow
            TabIndex        =   99
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OUT"
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
            Left            =   660
            MousePointer    =   1  'Arrow
            TabIndex        =   98
            Top             =   840
            Width           =   465
         End
         Begin VB.Image Turning_gear_OUT 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   0
            Left            =   180
            Picture         =   "Main_Bearings.frx":030A
            Stretch         =   -1  'True
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image Turning_gear_IN 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   1
            Left            =   180
            Picture         =   "Main_Bearings.frx":0614
            Stretch         =   -1  'True
            Top             =   300
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image Turning_gear_IN 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   0
            Left            =   180
            Picture         =   "Main_Bearings.frx":091E
            Stretch         =   -1  'True
            Top             =   300
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin FXRTX50GLib.FXRtext FXRtext2 
         Height          =   1875
         Left            =   780
         TabIndex        =   89
         Top             =   5880
         Width           =   315
         _Version        =   327680
         _ExtentX        =   556
         _ExtentY        =   3307
         _StockProps     =   70
         Caption         =   "Astern Element"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Persistence     =   -1  'True
         _StdProps       =   -4088998
         ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
         ErrCode         =   909381820
         ErrInfo         =   102046478
         Angle           =   90
         BackColor       =   8421504
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1875
         Left            =   7260
         Max             =   255
         TabIndex        =   37
         Top             =   6180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   420
         Top             =   9060
      End
      Begin VB.Frame Sliding_frame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   107
         Top             =   2400
         Width           =   435
         Begin VB.Line Line21 
            BorderWidth     =   2
            X1              =   120
            X2              =   360
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label16 
            Caption         =   "H"
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
            Left            =   60
            TabIndex        =   111
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label15 
            Caption         =   "N"
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
            Left            =   195
            TabIndex        =   110
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label14 
            Caption         =   "C"
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
            Left            =   345
            TabIndex        =   109
            Top             =   120
            Width           =   135
         End
         Begin VB.Line Line24 
            BorderWidth     =   2
            X1              =   300
            X2              =   300
            Y1              =   0
            Y2              =   120
         End
         Begin VB.Line Line10 
            BorderWidth     =   2
            X1              =   360
            X2              =   360
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line Line22 
            BorderWidth     =   2
            X1              =   120
            X2              =   120
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line Line23 
            BorderWidth     =   2
            X1              =   180
            X2              =   180
            Y1              =   0
            Y2              =   120
         End
         Begin VB.Shape Shape37 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Top             =   0
            Width           =   75
         End
         Begin VB.Shape Shape36 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   180
            Top             =   0
            Width           =   135
         End
         Begin VB.Shape Shape38 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   300
            Top             =   0
            Width           =   75
         End
      End
      Begin VB.Shape Shaft_locked 
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   10740
         Top             =   4740
         Width           =   255
      End
      Begin VB.Shape Shape40 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   10740
         Top             =   3660
         Width           =   255
      End
      Begin VB.Shape Shaft_locked 
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   10740
         Top             =   3480
         Width           =   255
      End
      Begin VB.OLE FOOT_OLE 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1260
         OleObjectBlob   =   "Main_Bearings.frx":0C28
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Arrow.vsd"
         TabIndex        =   117
         Top             =   2160
         Width           =   90
      End
      Begin VB.Shape Shaft_un_locked 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   10740
         Top             =   4860
         Width           =   255
      End
      Begin VB.Shape Shaft_un_locked 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   10740
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Casing"
         Height          =   195
         Left            =   3000
         TabIndex        =   116
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "HP Turbine"
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
         Left            =   1320
         TabIndex        =   113
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Shape oval_shape 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   435
         Left            =   3060
         Shape           =   2  'Oval
         Top             =   2880
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Line HP_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   2340
         X2              =   2340
         Y1              =   3300
         Y2              =   2820
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   720
         Top             =   1260
         Width           =   135
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Sliding Foot"
         Height          =   195
         Left            =   0
         TabIndex        =   108
         Top             =   2220
         Width           =   915
      End
      Begin VB.Image Arrow_875 
         Height          =   255
         Left            =   1140
         Picture         =   "Main_Bearings.frx":2A40
         Stretch         =   -1  'True
         Top             =   660
         Width           =   255
      End
      Begin VB.Shape Shape35 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   5520
         Top             =   1380
         Width           =   135
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   4260
         X2              =   4920
         Y1              =   2040
         Y2              =   1740
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   4260
         X2              =   4920
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4230
         Top             =   1500
         Width           =   735
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   3600
         Top             =   1320
         Width           =   135
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   3420
         X2              =   3600
         Y1              =   1980
         Y2              =   1320
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   3420
         X2              =   3600
         Y1              =   1320
         Y2              =   1980
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   3300
         Top             =   1320
         Width           =   135
      End
      Begin VB.Line Turning_gear_line 
         BorderWidth     =   2
         X1              =   5640
         X2              =   6000
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   7200
         TabIndex        =   2
         ToolTipText     =   "HP Pinion Aft"
         Top             =   8460
         Width           =   435
      End
      Begin VB.Shape Shape30 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   10980
         Top             =   3660
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "Turning Gear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5820
         TabIndex        =   103
         Top             =   1260
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Image Turning_gear_OUT 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   5820
         Picture         =   "Main_Bearings.frx":2D4A
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image Turning_gear_IN 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   5640
         Picture         =   "Main_Bearings.frx":3054
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Shaft_shp 
         Alignment       =   2  'Center
         Caption         =   "xxxxx shp"
         Height          =   255
         Left            =   4680
         TabIndex        =   94
         Top             =   4380
         Width           =   855
      End
      Begin VB.Label Shaft_kW 
         Alignment       =   2  'Center
         Caption         =   "xxxxx kW"
         Height          =   255
         Left            =   4680
         TabIndex        =   93
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Shaft_rpm 
         Alignment       =   2  'Center
         Caption         =   "xxx rpm"
         Height          =   255
         Left            =   4740
         TabIndex        =   92
         Top             =   3780
         Width           =   795
      End
      Begin VB.Shape Shape34 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   5475
         Top             =   5340
         Width           =   135
      End
      Begin VB.Shape Shape33 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   5775
         Top             =   5340
         Width           =   135
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5595
         X2              =   5775
         Y1              =   5340
         Y2              =   6000
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   5595
         X2              =   5775
         Y1              =   6000
         Y2              =   5340
      End
      Begin VB.Shape Shape29 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   5475
         Top             =   2340
         Width           =   135
      End
      Begin VB.Shape Shape23 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   5775
         Top             =   2340
         Width           =   135
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   5595
         X2              =   5775
         Y1              =   2340
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   5595
         X2              =   5775
         Y1              =   3000
         Y2              =   2340
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   4260
         X2              =   4980
         Y1              =   6420
         Y2              =   6780
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   4200
         X2              =   4920
         Y1              =   6780
         Y2              =   6420
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   4200
         Top             =   6600
         Width           =   795
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   6420
         X2              =   7140
         Y1              =   5640
         Y2              =   5280
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   6480
         X2              =   7200
         Y1              =   5280
         Y2              =   5640
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   6420
         Top             =   5460
         Width           =   795
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   6420
         X2              =   7140
         Y1              =   2700
         Y2              =   3060
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   6480
         X2              =   7200
         Y1              =   3120
         Y2              =   2700
      End
      Begin VB.Shape Shape20 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   6420
         Top             =   2460
         Width           =   795
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   10200
         Top             =   3960
         Width           =   2235
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   26
         Left            =   600
         Top             =   1740
         Width           =   75
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   26
         Left            =   600
         Top             =   1260
         Width           =   75
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   5
         Left            =   1080
         Picture         =   "Main_Bearings.frx":335E
         Stretch         =   -1  'True
         Top             =   8580
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   4
         Left            =   600
         Picture         =   "Main_Bearings.frx":3668
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   3
         Left            =   1920
         Picture         =   "Main_Bearings.frx":3972
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   2
         Left            =   2640
         Picture         =   "Main_Bearings.frx":3C7C
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "875 psi Steam"
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
         Left            =   60
         TabIndex        =   87
         Top             =   4500
         Width           =   1335
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   0
         Left            =   1800
         Picture         =   "Main_Bearings.frx":3F86
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   255
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2040
         X2              =   2040
         Y1              =   8040
         Y2              =   8280
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   26
         Left            =   10560
         TabIndex        =   86
         Top             =   8640
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   25
         Left            =   10560
         TabIndex        =   85
         Top             =   8280
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   24
         Left            =   10560
         TabIndex        =   84
         Top             =   8040
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   23
         Left            =   10560
         TabIndex        =   83
         Top             =   7800
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   22
         Left            =   10560
         TabIndex        =   82
         Top             =   7560
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   21
         Left            =   10560
         TabIndex        =   81
         Top             =   7320
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   20
         Left            =   10560
         TabIndex        =   80
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   19
         Left            =   10560
         TabIndex        =   79
         Top             =   6840
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   18
         Left            =   10560
         TabIndex        =   78
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   17
         Left            =   10560
         TabIndex        =   77
         Top             =   6360
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   16
         Left            =   10560
         TabIndex        =   76
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   15
         Left            =   10560
         TabIndex        =   75
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   14
         Left            =   10560
         TabIndex        =   74
         Top             =   5640
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   13
         Left            =   10560
         TabIndex        =   73
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   12
         Left            =   10560
         TabIndex        =   72
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Cross Over Pipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2160
         TabIndex        =   71
         Top             =   3840
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "875 psi Steam"
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
         Left            =   300
         TabIndex        =   70
         Top             =   360
         Width           =   1335
      End
      Begin VB.Line LP_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2760
         X2              =   2760
         Y1              =   2580
         Y2              =   5820
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   1200
         X2              =   1200
         Y1              =   8400
         Y2              =   8700
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   720
         X2              =   720
         Y1              =   4800
         Y2              =   5820
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   11
         Left            =   10620
         TabIndex        =   69
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   10
         Left            =   10620
         TabIndex        =   68
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   9
         Left            =   10620
         TabIndex        =   67
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   8
         Left            =   10620
         TabIndex        =   66
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   7
         Left            =   10620
         TabIndex        =   65
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   6
         Left            =   10620
         TabIndex        =   64
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   5
         Left            =   10620
         TabIndex        =   63
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   4
         Left            =   10620
         TabIndex        =   62
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   3
         Left            =   10620
         TabIndex        =   61
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36 degC"
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
         Index           =   2
         Left            =   10620
         TabIndex        =   60
         Top             =   720
         Width           =   735
      End
      Begin VB.Shape Shape19 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6840
         Shape           =   2  'Oval
         Top             =   2100
         Width           =   135
      End
      Begin VB.Shape Shape18 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6660
         Shape           =   2  'Oval
         Top             =   2160
         Width           =   75
      End
      Begin VB.Line Line50 
         BorderWidth     =   2
         X1              =   6600
         X2              =   7020
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line49 
         BorderWidth     =   2
         X1              =   6840
         X2              =   7020
         Y1              =   1740
         Y2              =   2040
      End
      Begin VB.Line Line48 
         BorderWidth     =   2
         X1              =   6780
         X2              =   6600
         Y1              =   1740
         Y2              =   2040
      End
      Begin VB.Line Line47 
         BorderWidth     =   2
         X1              =   6840
         X2              =   6840
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line Line46 
         BorderWidth     =   2
         X1              =   6780
         X2              =   6780
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4620
         Shape           =   2  'Oval
         Top             =   1140
         Width           =   135
      End
      Begin VB.Shape Shape15 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4440
         Shape           =   2  'Oval
         Top             =   1200
         Width           =   75
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6840
         Shape           =   2  'Oval
         Top             =   6060
         Width           =   75
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6660
         Shape           =   2  'Oval
         Top             =   6180
         Width           =   135
      End
      Begin VB.Line Line45 
         BorderWidth     =   2
         X1              =   6840
         X2              =   7020
         Y1              =   6660
         Y2              =   6360
      End
      Begin VB.Line Line44 
         BorderWidth     =   2
         X1              =   6600
         X2              =   6780
         Y1              =   6360
         Y2              =   6660
      End
      Begin VB.Line Line43 
         BorderWidth     =   2
         X1              =   6600
         X2              =   7020
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line42 
         BorderWidth     =   2
         X1              =   6840
         X2              =   6840
         Y1              =   6600
         Y2              =   6960
      End
      Begin VB.Line Line41 
         BorderWidth     =   2
         X1              =   6780
         X2              =   6780
         Y1              =   6600
         Y2              =   6960
      End
      Begin VB.Shape Shape11 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4620
         Shape           =   2  'Oval
         Top             =   7020
         Width           =   75
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4440
         Shape           =   2  'Oval
         Top             =   7140
         Width           =   135
      End
      Begin VB.Line Line40 
         BorderWidth     =   2
         X1              =   4620
         X2              =   4800
         Y1              =   7620
         Y2              =   7320
      End
      Begin VB.Line Line39 
         BorderWidth     =   2
         X1              =   4380
         X2              =   4560
         Y1              =   7320
         Y2              =   7620
      End
      Begin VB.Line Line38 
         BorderWidth     =   2
         X1              =   4380
         X2              =   4800
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line37 
         BorderWidth     =   2
         X1              =   4620
         X2              =   4620
         Y1              =   7560
         Y2              =   7920
      End
      Begin VB.Line Line36 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   7560
         Y2              =   7920
      End
      Begin VB.Line Line35 
         BorderWidth     =   2
         X1              =   4380
         X2              =   4800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line33 
         BorderWidth     =   2
         X1              =   4620
         X2              =   4800
         Y1              =   780
         Y2              =   1080
      End
      Begin VB.Line Line32 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4380
         Y1              =   780
         Y2              =   1080
      End
      Begin VB.Line Line31 
         BorderWidth     =   2
         X1              =   4620
         X2              =   4620
         Y1              =   480
         Y2              =   840
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   480
         Y2              =   840
      End
      Begin VB.Label List_lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "LO Cooler Outlet"
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
         Index           =   26
         Left            =   8220
         TabIndex        =   57
         Top             =   8700
         Width           =   2235
      End
      Begin VB.Shape Bearing_C 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   13
         Left            =   8940
         Top             =   4440
         Width           =   75
      End
      Begin VB.Shape Bearing_C 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   12
         Left            =   8400
         Top             =   4440
         Width           =   75
      End
      Begin VB.Label List_lbl 
         Caption         =   "25. Stern Tube Brg"
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
         Index           =   25
         Left            =   8100
         TabIndex        =   33
         Top             =   8280
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "24. Line Shaft"
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
         Index           =   24
         Left            =   8100
         TabIndex        =   32
         Top             =   8040
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "23. LP Second Redn Pinion Aft"
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
         Index           =   23
         Left            =   8100
         TabIndex        =   31
         Top             =   7800
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "22. LP Second Redn Pinion Fwd"
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
         Index           =   22
         Left            =   8100
         TabIndex        =   30
         Top             =   7560
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "21. LP First Redn Gear Aft"
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
         Index           =   21
         Left            =   8100
         TabIndex        =   29
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "20. LP First Redn Gear Fwd"
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
         Index           =   20
         Left            =   8100
         TabIndex        =   28
         Top             =   7080
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "19. LP Pinion Aft"
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
         Index           =   19
         Left            =   8100
         TabIndex        =   27
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "18. LP Pinion Fwd"
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
         Index           =   18
         Left            =   8100
         TabIndex        =   26
         Top             =   6600
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "17. LP Thrust Ahead"
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
         Index           =   17
         Left            =   8100
         TabIndex        =   25
         Top             =   6360
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "16. LP Thrust Astern"
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
         Index           =   16
         Left            =   8100
         TabIndex        =   24
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "15. LP Turbine Aft"
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
         Index           =   15
         Left            =   8100
         TabIndex        =   23
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "14. LP Turbine Fwd"
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
         Index           =   14
         Left            =   8100
         TabIndex        =   22
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "13. Main Thrust Brg Astern"
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
         Index           =   13
         Left            =   8100
         TabIndex        =   20
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "12. Main Thrust Brg Ahead"
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
         Index           =   12
         Left            =   8100
         TabIndex        =   19
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label List_lbl 
         Caption         =   "11. Bull Wheel Aft"
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
         Index           =   11
         Left            =   8220
         TabIndex        =   18
         Top             =   2880
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "10. Bull Wheel Fwd"
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
         Index           =   10
         Left            =   8220
         TabIndex        =   17
         Top             =   2640
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "9. HP Second Redn Pinion Aft"
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
         Index           =   9
         Left            =   8220
         TabIndex        =   16
         Top             =   2400
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "8. HP Second Redn Pinion Fwd"
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
         Index           =   8
         Left            =   8220
         TabIndex        =   15
         Top             =   2160
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "7. HP First Redn Gear Aft"
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
         Index           =   7
         Left            =   8220
         TabIndex        =   14
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "6. HP First Redn Gear Fwd."
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
         Index           =   6
         Left            =   8220
         TabIndex        =   13
         Top             =   1680
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "5. HP Pinion Aft"
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
         Index           =   5
         Left            =   8220
         TabIndex        =   12
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "4. HP Pinion Fwd"
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
         Index           =   4
         Left            =   8220
         TabIndex        =   11
         Top             =   1200
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "3. HP Turbine Aft"
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
         Index           =   3
         Left            =   8220
         TabIndex        =   10
         Top             =   960
         Width           =   2355
      End
      Begin VB.Label List_lbl 
         Caption         =   "2. HP Thrust Brg"
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
         Index           =   2
         Left            =   8220
         TabIndex        =   9
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Param_lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "79 F"
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
         Left            =   10620
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label List_lbl 
         Caption         =   "1. HP Turbine Fwd"
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
         Index           =   1
         Left            =   8220
         TabIndex        =   7
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Param_lbl 
         Caption         =   "36 degC"
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
         Index           =   0
         Left            =   6780
         TabIndex        =   6
         Top             =   660
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   25
         Left            =   11940
         Top             =   4440
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   24
         Left            =   9720
         Top             =   4500
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   23
         Left            =   7260
         Top             =   5880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   22
         Left            =   5940
         Top             =   5880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   21
         Left            =   4980
         Top             =   5880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   20
         Left            =   3780
         Top             =   5880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   19
         Left            =   5040
         Top             =   6960
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   18
         Left            =   3780
         Top             =   6960
         Width           =   375
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   17
         Left            =   2100
         Top             =   6840
         Width           =   75
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   16
         Left            =   2460
         Top             =   6840
         Width           =   75
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   15
         Left            =   2820
         Top             =   6960
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   14
         Left            =   180
         Top             =   6960
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   13
         Left            =   8940
         Top             =   4020
         Width           =   75
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   11
         Left            =   7200
         Top             =   4440
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   10
         Left            =   5940
         Top             =   4440
         Width           =   435
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Propellor Shaft"
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
         Left            =   7260
         TabIndex        =   3
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   25
         Left            =   11940
         Top             =   3840
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   24
         Left            =   9720
         Top             =   3720
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   23
         Left            =   7260
         Top             =   5460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   22
         Left            =   5940
         Top             =   5460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   21
         Left            =   4980
         Top             =   5460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   20
         Left            =   3780
         Top             =   5460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   19
         Left            =   5040
         Top             =   6540
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   18
         Left            =   3780
         Top             =   6540
         Width           =   375
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   17
         Left            =   2100
         Top             =   6480
         Width           =   75
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   16
         Left            =   2460
         Top             =   6480
         Width           =   75
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   15
         Left            =   2820
         Top             =   6540
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   14
         Left            =   180
         Top             =   6540
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   13
         Left            =   8940
         Top             =   3540
         Width           =   75
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   2220
         Top             =   6480
         Width           =   195
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   9660
         Top             =   3840
         Width           =   555
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   12
         Left            =   8400
         Top             =   4020
         Width           =   75
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   12
         Left            =   8400
         Top             =   3540
         Width           =   75
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   11
         Left            =   7200
         Top             =   3840
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   10
         Left            =   5940
         Top             =   3840
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   9
         Left            =   7260
         Top             =   2880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   8
         Left            =   5940
         Top             =   2880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   7
         Left            =   4980
         Top             =   2880
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   6
         Left            =   3780
         Top             =   2880
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   7
         Left            =   4980
         Top             =   2460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   9
         Left            =   7260
         Top             =   2460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   8
         Left            =   5940
         Top             =   2460
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   6
         Left            =   3780
         Top             =   2460
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   5
         Left            =   5040
         Top             =   1860
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   4
         Left            =   3780
         Top             =   1860
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   2835
         Top             =   1860
         Width           =   435
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   900
         Top             =   1740
         Width           =   75
      End
      Begin VB.Shape Bearing_B 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   120
         Top             =   1860
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   5
         Left            =   5040
         Top             =   1440
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   4
         Left            =   3780
         Top             =   1440
         Width           =   435
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   900
         Top             =   1260
         Width           =   75
      End
      Begin VB.Shape Bearing_A 
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   120
         Top             =   1440
         Width           =   435
      End
      Begin VB.Shape Shape44 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1575
         Left            =   8520
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Bull Gear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   6540
         TabIndex        =   1
         Top             =   3120
         Width           =   525
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape28 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   45
         Top             =   6660
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Left            =   1380
         TabIndex        =   0
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   6480
         X2              =   7140
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   3420
         X2              =   3600
         Y1              =   7080
         Y2              =   6420
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   3420
         X2              =   3600
         Y1              =   6420
         Y2              =   7080
      End
      Begin VB.Shape Shape32 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   3600
         Top             =   6420
         Width           =   135
      End
      Begin VB.Shape Shape31 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   3300
         Top             =   6420
         Width           =   135
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   4260
         X2              =   4920
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Shape Shape27 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2115
         Left            =   4260
         Top             =   4980
         Width           =   675
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   4260
         X2              =   4920
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Shape Bearing_A 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   2820
         Top             =   1440
         Width           =   435
      End
      Begin VB.Shape Shape24 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2775
         Top             =   6660
         Width           =   2715
      End
      Begin VB.Shape Shape25 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   3555
         Left            =   6480
         Top             =   2400
         Width           =   675
      End
      Begin VB.Shape Shape26 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3780
         Top             =   5580
         Width           =   3975
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Index           =   14
         Left            =   180
         TabIndex        =   45
         ToolTipText     =   "LP Turbine Fwd"
         Top             =   6240
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   15
         Left            =   2820
         TabIndex        =   46
         ToolTipText     =   "LP Turbine Aft"
         Top             =   6300
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   16
         Left            =   2400
         TabIndex        =   47
         ToolTipText     =   "LP Thrust Astern"
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   17
         Left            =   1860
         TabIndex        =   48
         ToolTipText     =   "LP Thrust Ahead"
         Top             =   6240
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   18
         Left            =   3780
         TabIndex        =   49
         ToolTipText     =   "LP Pinion Fwd"
         Top             =   6300
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   19
         Left            =   4980
         TabIndex        =   50
         ToolTipText     =   "LP Pinion Aft"
         Top             =   6300
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   3780
         TabIndex        =   51
         ToolTipText     =   "LP First Redn Fwd"
         Top             =   5220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   21
         Left            =   4980
         TabIndex        =   52
         ToolTipText     =   "LP First Redn Aft"
         Top             =   5220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   6060
         TabIndex        =   53
         ToolTipText     =   "LP First Redn Pinion"
         Top             =   5220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   23
         Left            =   7200
         TabIndex        =   54
         ToolTipText     =   "LP First Redn Pinion Aft"
         Top             =   5220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   24
         Left            =   9720
         TabIndex        =   55
         ToolTipText     =   "Line Shaft"
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Index           =   25
         Left            =   12000
         TabIndex        =   56
         ToolTipText     =   "Stern Tube Brg"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nozzle"
         Height          =   435
         Left            =   4440
         TabIndex        =   58
         ToolTipText     =   "Nozzles are fitted to spray lubricating oil into the meshing teeth to protect them from scuffing damage."
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Nozzle"
         Height          =   675
         Left            =   4560
         TabIndex        =   59
         ToolTipText     =   "Nozzles are fitted to spray lubricating oil into the meshing teeth to protect them from scuffing damage."
         Top             =   7320
         Width           =   795
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2055
         Left            =   4260
         Top             =   1380
         Width           =   675
      End
      Begin VB.Shape Shape17 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3780
         Top             =   2580
         Width           =   3915
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   6
         Left            =   3780
         TabIndex        =   36
         ToolTipText     =   "HP First Redn Fwd"
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   7
         Left            =   5040
         TabIndex        =   38
         ToolTipText     =   "HP First Redn Aft"
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   8
         Left            =   6060
         TabIndex        =   39
         ToolTipText     =   "HP First Redn Pinion Fwd"
         Top             =   2220
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   9
         Left            =   7260
         TabIndex        =   40
         ToolTipText     =   "HP First Redn Pinion Aft"
         Top             =   2220
         Width           =   375
      End
      Begin VB.Shape Shape43 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   5820
         Top             =   3960
         Width           =   4395
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   10
         Left            =   5940
         TabIndex        =   41
         ToolTipText     =   "Bull Wheel Fwd"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   11
         Left            =   7200
         TabIndex        =   42
         ToolTipText     =   "Bull Wheel Aft"
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   13
         Left            =   8880
         TabIndex        =   44
         ToolTipText     =   "Main Thrust Brg Astern"
         Top             =   3300
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   12
         Left            =   8220
         TabIndex        =   43
         ToolTipText     =   "Main Thrust Brg ahead"
         Top             =   3300
         Width           =   255
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4005
         Left            =   660
         OleObjectBlob   =   "Main_Bearings.frx":4290
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   91
         Top             =   4980
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   60
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "HP Turbine Fwd"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   26
         Left            =   540
         TabIndex        =   88
         ToolTipText     =   "HP Thrust Bearing"
         Top             =   1020
         Width           =   195
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "HP Thrust Bearing"
         Top             =   1020
         Width           =   195
      End
      Begin VB.OLE FOOT_OLE_UP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "Main_Bearings.frx":BCA8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "\\Pc00\PC00_C\Refrigeration\Refrigeration_Dwgs\Arrow.vsd"
         TabIndex        =   112
         Top             =   300
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Shape HP_shaft 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2520
         Top             =   1560
         Width           =   3150
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   3
         Left            =   2880
         TabIndex        =   21
         ToolTipText     =   "HP Turbine Aft"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   4
         Left            =   3900
         TabIndex        =   34
         ToolTipText     =   "HP Turbine Fwd"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Bearing_lbl 
         Alignment       =   2  'Center
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   5
         Left            =   4980
         TabIndex        =   35
         ToolTipText     =   "HP Pinion Aft"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Line Line25 
         Visible         =   0   'False
         X1              =   2640
         X2              =   3000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Shaft_locked_lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Shaft Locked"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11040
         TabIndex        =   119
         Top             =   3420
         Width           =   1035
      End
      Begin VB.Label Shaft_un_locked_lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Shaft Un-Locked"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   11040
         TabIndex        =   118
         Top             =   3300
         Width           =   1335
      End
      Begin VB.OLE HP_OLE 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1965
         Left            =   1200
         OleObjectBlob   =   "Main_Bearings.frx":DAC0
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\My Documents\visiotbne.vsd"
         TabIndex        =   90
         Top             =   660
         Width           =   1620
      End
      Begin VB.Line LINE_875 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   1260
         X2              =   1260
         Y1              =   720
         Y2              =   1200
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   1920
         X2              =   1920
         Y1              =   2340
         Y2              =   2880
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2460
      TabIndex        =   104
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Main_Bearings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim set_dummy_focus As Integer
Dim HP_ole_width As Single
Dim HP_ole_orig_width As Single
Dim T_foot_val As Integer
Dim T_foot_old As Integer

' 5/9/00
Dim HP_ole_left As Single
Dim HP_ole_orig_left As Single


Private Sub Command1_Click()


End Sub

Private Sub Command3D12_Click()



End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idMain_Bearings)

    ' 27/7/00
HP_ole_width = HP_OLE.Width
HP_ole_orig_width = HP_OLE.Width

' 5/9/00
HP_ole_left = HP_OLE.Left
HP_ole_orig_left = HP_OLE.Left

T_foot_old = -100

VScroll2.Value = T_foot_val


' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' === 14/03/01 ===
    'TURN_IN.Visible = False
    'Turn_OUT.Visible = False
    
    'Frame1.Visible = False
    
     Frame2.Visible = False
    
Else

    TURN_IN.Visible = True
    Turn_OUT.Visible = True
    
    Frame1.Visible = True
    Frame2.Visible = True
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idMain_Bearings)
    HavenStretch1.Visible = True
    
    ' 27/7/00
    HP_ole_width = HP_OLE.Width
    T_foot_old = -123
    
    ' 4/9/00
    HP_ole_left = HP_OLE.Left
    
    
End If

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 3/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


' Turning Gear STOP (Pump number 28)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 28))
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


' Turning Gear START (Pump number 28)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 28))
End If

End Sub

Private Sub Param_lbl_Click(Index As Integer)

' === 10/10/01 ===
If Index < 26 Then
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 120
    Alarm_set_point.Show
End If


End Sub

Private Sub Timer1_Timer()
Dim Pshaft As Single

Dim RED_value As Integer
Dim RED_value_quartered As Integer

Dim HP_inc As Single


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 8/9/00
If VB_steam_logicals(132) Then

  Shaft_locked_lbl.Visible = True
  Shaft_un_locked_lbl.Visible = False
  
  Shaft_locked(0).Visible = True
  Shaft_locked(1).Visible = True
  
  Shaft_un_locked(0).Visible = False
  Shaft_un_locked(1).Visible = False
  
Else
  Shaft_locked_lbl.Visible = False
  Shaft_un_locked_lbl.Visible = True
  
  Shaft_locked(0).Visible = False
  Shaft_locked(1).Visible = False
  
  Shaft_un_locked(0).Visible = True
  Shaft_un_locked(1).Visible = True
End If




' 27/7/00
'   PC_REALS(143) = T_sliding_foot

' 5/9/00 test by commenting out....
T_foot_val = LSS_reals(143)

' 5/9/00
'T_foot_txt.text = Format(LSS_reals(143) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_foot_txt.text = Format(T_foot_val * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

If T_foot_old <> T_foot_val Then

  T_foot_old = T_foot_val
  
  
 ' TESTING
 ' HP_inc = ( VScroll2.Value)  * HP_ole_width / HP_ole_orig_width
   HP_inc = (T_foot_val * 0.5) * HP_ole_width / HP_ole_orig_width
  
  HP_OLE.Width = HP_ole_width + HP_inc
  
  ' 5/9/00
  HP_OLE.Left = HP_ole_left - HP_inc

  ' 5/9/00
 'FOOT_OLE.Left = HP_OLE.Left + HP_OLE.Width - FOOT_OLE.Width
  FOOT_OLE.Left = HP_OLE.Left + FOOT_OLE.Width / 2
  LINE_875.X1 = FOOT_OLE.Left
  LINE_875.X2 = FOOT_OLE.Left
  Arrow_875.Left = HP_OLE.Left - Arrow_875.Width / 4

  ' 5/9/00
  'HP_pipe.X1 = FOOT_OLE.Left
  'HP_pipe.X2 = HP_pipe.X1

  oval_shape.Left = HP_pipe.X1
  
  ' 15/8/00   Got a CRASH when LP_pipe.X1 was 2760 and HP_pipe.X1 was 2779
  '  SO
  '  oval_shape.Width = LP_pipe.X1 - HP_pipe.X1
  If LP_pipe.X1 > HP_pipe.X1 Then
     oval_shape.Width = LP_pipe.X1 - HP_pipe.X1
  Else
     oval_shape.Width = 15   ' Seems to be the MIN value
  End If
  
End If


' 23/11/99
'single_char = Mid$(PC_INT_STRING, 362, 1)
'If single_char = "1" Then
If VB_steam_logicals(94) Then

  '15/2/00
    Turning_gear_line.Visible = True

    Turning_gear_IN(0).Visible = False    ' GREY image
    Turning_gear_IN(1).Visible = True     ' GREEN image
    Turning_gear_IN(2).Visible = True     ' GREEN image
    
   Turning_gear_OUT(0).Visible = True     ' GREY image
   Turning_gear_OUT(1).Visible = False    ' RED image
   Turning_gear_OUT(2).Visible = False    ' RED image
   
Else

  '15/2/00
    Turning_gear_line.Visible = False
    
    Turning_gear_IN(0).Visible = True     ' GREY image
    Turning_gear_IN(1).Visible = False    ' GREEN image
    Turning_gear_IN(2).Visible = False    ' GREEN image
    
   Turning_gear_OUT(0).Visible = False    ' GREY image
   Turning_gear_OUT(1).Visible = True     ' RED image
   Turning_gear_OUT(2).Visible = True     ' RED image
   
End If


' 11/11/99  From Boilers.frm  (TEST FORM )
'    List8.AddItem "shaft rpm " & Format(LSS_reals(80), "#####0.00")
'
'    Pshaft = LSS_reals(797) + LSS_reals(798) + LSS_reals(799) + LSS_reals(800)
'
'    List8.AddItem "shaft power = " & Format(Pshaft, "##0.0") & " kW"
'    List8.AddItem "shaft power = " & Format(Pshaft / 0.746, "##0.0") & " hp"

' 11/11/99
shaft_rpm.Caption = Format(LSS_reals(80), "##0") & " rpm"


'cc 11 / 2 / 0
'        PC_REALS(801) = Power_astern

'  11/2/00
'Pshaft = LSS_reals(797) + LSS_reals(798) + LSS_reals(799) + LSS_reals(800)
Pshaft = LSS_reals(797) + LSS_reals(798) + LSS_reals(799) + LSS_reals(800) + LSS_reals(801)

Shaft_kW.Caption = Format(Pshaft, "#####0") & " kW"
Shaft_shp.Caption = Format(Pshaft / 0.746, "#####0") & " shp"


' LO Cooler outlet Temperature
Param_lbl(26).Caption = Format(LSS_reals(206) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' Main Bearing Temperatures
For i = 1 To 25

   Param_lbl(i).Caption = Format(LSS_reals(i + 89) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

       ' === 10/10/01 ===
       ' if( (T_bearing(i) * 1.8 + 32.0)  .gt. ALARM_HILO(1,120) ) then
       '     LSS_LOGICALS(i+64) = .true.

        If LSS_LOGICALS(i + 64) Then
            Param_lbl(i).BackColor = &HFF&          ' RED
        Else
            Param_lbl(i).BackColor = &HFF00&        ' GREEN
        End If
       ' === 10/10/01 ===
        
        
        

    ' 2/9/99
    ' 6/89/99  400.0 changed to 200.0
    If (LSS_reals(i + 89)) < 200# Then
        If (LSS_reals(i + 89)) > 0# Then
        
            ' 11/11/99  To reduce screen flicker
            'VScroll1.Value = Int(LSS_reals(i + 89) * 255# / 200#)
            RED_value_quartered = Int(LSS_reals(i + 89) * 63# / 200#)
        Else
            'VScroll1.Value = 0
            RED_value_quartered = 0
        End If
    Else
           'VScroll1.Value = 255
            RED_value_quartered = 63
    End If
    
    
      'If (old_bearing_value <> VScroll1.Value) Then
       If (old_bearing_value(i) <> RED_value_quartered) Then
       
                       RED_value = RED_value_quartered * 4
                       
            old_bearing_value(i) = RED_value_quartered
            
                Bearing_A(i).FillColor = RGB(RED_value, 0, 0)
            Bearing_A(i).BorderColor = RGB(RED_value, 0, 0)
    
            Bearing_B(i).FillColor = RGB(RED_value, 0, 0)
            Bearing_B(i).BorderColor = RGB(RED_value, 0, 0)
          
                    Bearing_C(12).FillColor = Bearing_A(12).FillColor
            Bearing_C(12).BorderColor = Bearing_A(12).BorderColor
      
                Bearing_C(13).FillColor = Bearing_A(13).FillColor
            Bearing_C(13).BorderColor = Bearing_A(13).BorderColor

       End If


Next i

' Note: update_pump_LSS(syxxx,index,nPump)

' 23/11/99 Turning Gear
Call update_pump_LSS(Me, 0, 28)      ' Turning Gear


End Sub

Private Sub TURN_IN_Click()

' 23/11/99
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 3/4/00  NOT from ACT's...

' === 13/03/01 MUST now be operable from ACT's ===
'If Not I_AM_INSTRUCTOR And Not PTT Then
'      ' I MUST be an ACT Terminal (Full Mission) so
'      Exit Sub
'End If



Call setPCDI(246)    ' Turning Gear IN

End Sub

Private Sub Turn_OUT_Click()

' 23/11/99
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 3/4/00  NOT from ACT's...

' === 13/03/01 MUST now be operable from ACT's ===
'If Not I_AM_INSTRUCTOR And Not PTT Then
'      ' I MUST be an ACT Terminal (Full Mission) so
'      Exit Sub
'End If


Call setPCDI(243)    ' Turning Gear OUT


End Sub

Private Sub VScroll1_Change()


'For i = 1 To 25
'      Bearing_A(i).FillColor = RGB(VScroll1.Value, 0, 0)
'    Bearing_A(i).BorderColor = RGB(VScroll1.Value, 0, 0)
'
'      Bearing_B(i).FillColor = RGB(VScroll1.Value, 0, 0)
'    Bearing_B(i).BorderColor = RGB(VScroll1.Value, 0, 0)
'
'Next i
'
'        Bearing_C(12).FillColor = RGB(VScroll1.Value, 0, 0)
'      Bearing_C(12).BorderColor = RGB(VScroll1.Value, 0, 0)
'
'        Bearing_C(13).FillColor = RGB(VScroll1.Value, 0, 0)
'      Bearing_C(13).BorderColor = RGB(VScroll1.Value, 0, 0)

End Sub

Private Sub VScroll2_Change()

T_foot_val = VScroll2.Value

End Sub
