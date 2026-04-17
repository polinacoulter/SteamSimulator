VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Main_Turbines 
   Caption         =   "Main Turbines"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Main_Turbines.frx":0000
   ScaleHeight     =   10455
   ScaleWidth      =   11850
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   10035
      Left            =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   17701
      Begin VB.CommandButton ONpump 
         BackColor       =   &H0000FF00&
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
         Left            =   8100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   9300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H008080FF&
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
         Left            =   8640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   9300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AUTO"
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
         Left            =   9240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   9300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox Psteam6 
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
         Left            =   1020
         TabIndex        =   132
         Text            =   "879 psi"
         Top             =   3180
         Width           =   675
      End
      Begin VB.Frame Turbine_Power_frame 
         Caption         =   "Turbine Power"
         Height          =   1515
         Left            =   7080
         TabIndex        =   121
         Top             =   300
         Width           =   3855
         Begin VB.Label Power_label 
            Alignment       =   2  'Center
            Caption         =   "xxxx kW"
            Height          =   255
            Index           =   3
            Left            =   1860
            TabIndex        =   129
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Power_label 
            Alignment       =   2  'Center
            Caption         =   "xxxx kW"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   128
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Power_label 
            Alignment       =   2  'Center
            Caption         =   "xxxx kW"
            Height          =   255
            Index           =   1
            Left            =   1860
            TabIndex        =   127
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Power_label 
            Alignment       =   2  'Center
            Caption         =   "xxxx kW"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   126
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            Caption         =   "HP 2nd Stage"
            Height          =   255
            Left            =   1920
            TabIndex        =   125
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "LP 1st Stage"
            Height          =   255
            Left            =   420
            TabIndex        =   124
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "LP 2nd Stage"
            Height          =   255
            Left            =   1920
            TabIndex        =   123
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            Caption         =   "HP 1st Stage"
            Height          =   255
            Left            =   420
            TabIndex        =   122
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.TextBox Fsteam16 
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
         Left            =   5160
         TabIndex        =   120
         Text            =   "xx .x kg/s"
         Top             =   8280
         Width           =   855
      End
      Begin VB.TextBox Fsteam15 
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
         Left            =   6060
         TabIndex        =   119
         Text            =   "xx .x kg/s"
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox Fsteam_LP_out 
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
         Left            =   1680
         TabIndex        =   118
         Text            =   "xx .x kg/s"
         Top             =   8820
         Width           =   855
      End
      Begin VB.TextBox Hsteam_LP_out 
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
         Left            =   1680
         TabIndex        =   117
         Text            =   "xxxx kJ/kg"
         Top             =   8520
         Width           =   855
      End
      Begin VB.TextBox Hsteam13 
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
         Left            =   5160
         TabIndex        =   116
         Text            =   "xxxx kJ/kg"
         Top             =   7980
         Width           =   855
      End
      Begin VB.TextBox Fsteam12 
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
         Left            =   4020
         TabIndex        =   115
         Text            =   "xx .x kg/s"
         Top             =   5820
         Width           =   855
      End
      Begin VB.TextBox Hsteam12 
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
         Left            =   5700
         TabIndex        =   114
         Text            =   "xxxx kJ/kg"
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox Fsteam14 
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
         Left            =   2700
         TabIndex        =   113
         Text            =   "xx .x kg/s"
         Top             =   2940
         Width           =   915
      End
      Begin VB.TextBox Fsteam11 
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
         Left            =   5700
         TabIndex        =   112
         Text            =   "xx .x kg/s"
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox Hsteam11 
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
         Left            =   2700
         TabIndex        =   111
         Text            =   "xxxx kJ/kg"
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox Hsteam10 
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
         Left            =   2940
         TabIndex        =   110
         Text            =   "xxxx kJ/kg"
         Top             =   0
         Width           =   915
      End
      Begin VB.TextBox Fsteam10 
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
         Left            =   2940
         TabIndex        =   109
         Text            =   "xx .x kg/s"
         Top             =   300
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
         Height          =   240
         Left            =   7920
         TabIndex        =   104
         Text            =   "P"
         Top             =   2520
         Width           =   195
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1455
         Left            =   6480
         Max             =   3
         TabIndex        =   103
         Top             =   3840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Left            =   180
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   8820
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton AUTO_guardian_VV 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Auto"
         Height          =   255
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6720
         Width           =   555
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Close"
         Height          =   255
         Left            =   60
         TabIndex        =   78
         Top             =   7440
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open"
         Height          =   255
         Left            =   60
         TabIndex        =   77
         Top             =   7200
         Visible         =   0   'False
         Width           =   556
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1455
         Left            =   4680
         Max             =   3
         TabIndex        =   65
         Top             =   3600
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   135
         Left            =   6480
         Max             =   3
         TabIndex        =   64
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   9240
         Max             =   3
         TabIndex        =   63
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   19
         Left            =   5700
         TabIndex        =   55
         Top             =   2340
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   795
         Index           =   3
         Left            =   1740
         TabIndex        =   44
         Top             =   7380
         Width           =   675
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Index           =   4
            X1              =   120
            X2              =   600
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Index           =   4
            X1              =   360
            X2              =   360
            Y1              =   120
            Y2              =   660
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            FillColor       =   &H000000FF&
            Height          =   555
            Index           =   4
            Left            =   60
            Shape           =   3  'Circle
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Port Boiler"
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2280
         Width           =   675
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Index           =   0
         Left            =   1680
         TabIndex        =   42
         Top             =   1320
         Width           =   675
         Begin VB.Shape Shape6 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   3
            FillColor       =   &H0000FF00&
            Height          =   555
            Index           =   1
            Left            =   0
            Shape           =   3  'Circle
            Top             =   120
            Width           =   615
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   3
            Index           =   1
            X1              =   300
            X2              =   300
            Y1              =   120
            Y2              =   660
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   3
            Index           =   1
            X1              =   60
            X2              =   540
            Y1              =   360
            Y2              =   360
         End
      End
      Begin VB.CommandButton Mollier_Command 
         Caption         =   "Mollier Diagram"
         Height          =   1155
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   7380
         Width           =   1035
      End
      Begin VB.TextBox Tsteam13 
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
         TabIndex        =   40
         Text            =   "354 F"
         Top             =   8880
         Width           =   675
      End
      Begin VB.TextBox Psteam13 
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
         TabIndex        =   39
         Text            =   "55 psi"
         Top             =   8580
         Width           =   675
      End
      Begin VB.TextBox Psteam11 
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
         Left            =   3120
         TabIndex        =   38
         Text            =   "xx psi"
         Top             =   2340
         Width           =   675
      End
      Begin VB.TextBox Tsteam11 
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
         Left            =   3120
         TabIndex        =   37
         Text            =   "xx  F"
         Top             =   2040
         Width           =   675
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   1080
         Max             =   1000
         TabIndex        =   33
         Top             =   7260
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   1320
         Max             =   1000
         TabIndex        =   32
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Tsteam_LP_out 
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
         Left            =   2640
         TabIndex        =   30
         Text            =   "91 F"
         Top             =   9300
         Width           =   675
      End
      Begin VB.TextBox Psteam14 
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
         Left            =   2640
         TabIndex        =   29
         Text            =   "0.05 psi"
         Top             =   9000
         Width           =   675
      End
      Begin VB.TextBox Tsteam12 
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
         Left            =   5160
         TabIndex        =   28
         Text            =   "354 F"
         Top             =   6720
         Width           =   675
      End
      Begin VB.TextBox Psteam12 
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
         Left            =   5160
         TabIndex        =   27
         Text            =   "55 psi"
         Top             =   6420
         Width           =   675
      End
      Begin VB.TextBox Psteam10 
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
         Left            =   3840
         TabIndex        =   26
         Text            =   "879 psi"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Tsteam10 
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
         Left            =   1020
         TabIndex        =   25
         Text            =   "955 F"
         Top             =   3480
         Width           =   675
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Condensate system"
         Height          =   555
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stbd Boiler"
         Height          =   675
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   675
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Feed Water System"
         Height          =   735
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Main Condenser"
         Height          =   735
         Left            =   9300
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   8100
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   1
         Left            =   6540
         TabIndex        =   3
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
         Index           =   3
         Left            =   420
         TabIndex        =   5
         Top             =   1380
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
         Index           =   5
         Left            =   9300
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   7
         Left            =   4440
         TabIndex        =   18
         Top             =   9120
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
         Left            =   420
         TabIndex        =   34
         Top             =   4380
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
         Height          =   255
         Index           =   90
         Left            =   1860
         TabIndex        =   35
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
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
         Index           =   10
         Left            =   1920
         TabIndex        =   36
         Top             =   6780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   13
         Left            =   7200
         TabIndex        =   47
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
         Height          =   375
         Index           =   24
         Left            =   3780
         TabIndex        =   48
         Top             =   3180
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
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
         Index           =   14
         Left            =   3780
         TabIndex        =   49
         Top             =   3780
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
         Left            =   3780
         TabIndex        =   50
         Top             =   4440
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
         Left            =   2520
         TabIndex        =   51
         Top             =   4140
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   17
         Left            =   4320
         TabIndex        =   52
         Top             =   4140
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
         Index           =   18
         Left            =   6900
         TabIndex        =   53
         Top             =   3060
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   20
         Left            =   4260
         TabIndex        =   56
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   21
         Left            =   8880
         TabIndex        =   58
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
         Index           =   8
         Left            =   9540
         TabIndex        =   66
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
         Index           =   12
         Left            =   7080
         TabIndex        =   67
         Top             =   7860
         Visible         =   0   'False
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
         Left            =   9240
         TabIndex        =   68
         Top             =   2940
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   2
         Left            =   1140
         TabIndex        =   69
         Top             =   6180
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
         Height          =   495
         Index           =   0
         Left            =   1740
         TabIndex        =   71
         Top             =   420
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   873
         TopColour       =   16384
         BottomColour    =   16384
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16384
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   11
         Left            =   1920
         TabIndex        =   72
         Top             =   6180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   9
         Left            =   1800
         TabIndex        =   73
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   23
         Left            =   5520
         TabIndex        =   93
         Top             =   3840
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
         Index           =   25
         Left            =   5520
         TabIndex        =   94
         Top             =   4500
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
         Index           =   26
         Left            =   6100
         TabIndex        =   96
         Top             =   4260
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
         Index           =   6
         Left            =   6900
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   27
         Left            =   5100
         TabIndex        =   107
         Top             =   4200
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   1440
         X2              =   1440
         Y1              =   780
         Y2              =   1140
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   2040
         X2              =   2400
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   10380
         Picture         =   "Main_Turbines.frx":0C42
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   45
         X1              =   9840
         X2              =   10620
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   44
         X1              =   7560
         X2              =   8940
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   43
         X1              =   5640
         X2              =   5640
         Y1              =   3720
         Y2              =   3960
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shaft Roll-over"
         Height          =   255
         Left            =   8100
         TabIndex        =   136
         Top             =   9000
         Width           =   1635
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   3780
         Picture         =   "Main_Turbines.frx":0F4C
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   3900
         X2              =   3900
         Y1              =   1980
         Y2              =   3840
      End
      Begin VB.Label VV_pos_throttle_astern 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2280
         TabIndex        =   131
         Top             =   6180
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   5640
         X2              =   8520
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Image Image26 
         Height          =   255
         Left            =   8640
         Picture         =   "Main_Turbines.frx":1256
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   255
      End
      Begin VB.Image Image25 
         Height          =   255
         Left            =   9000
         Picture         =   "Main_Turbines.frx":1560
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   255
      End
      Begin VB.Line Line4 
         Index           =   6
         X1              =   4800
         X2              =   4980
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   2400
         X2              =   3480
         Y1              =   8820
         Y2              =   8820
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   4140
         X2              =   5160
         Y1              =   8580
         Y2              =   8280
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   3600
         X2              =   3900
         Y1              =   2940
         Y2              =   2580
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   6420
         X2              =   5820
         Y1              =   1680
         Y2              =   2280
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   6180
         X2              =   4980
         Y1              =   1080
         Y2              =   2280
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   3360
         X2              =   3360
         Y1              =   480
         Y2              =   840
      End
      Begin VB.Line Line3 
         X1              =   3900
         X2              =   3900
         Y1              =   1680
         Y2              =   2040
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bulkhead Valve"
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
         Left            =   120
         TabIndex        =   106
         Top             =   3960
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image24 
         Height          =   255
         Left            =   120
         Picture         =   "Main_Turbines.frx":186A
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   40
         X1              =   240
         X2              =   240
         Y1              =   5160
         Y2              =   4680
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bulkhead Valve"
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
         Left            =   120
         TabIndex        =   105
         Top             =   960
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image23 
         Height          =   255
         Left            =   120
         Picture         =   "Main_Turbines.frx":1B74
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   31
         X1              =   240
         X2              =   240
         Y1              =   2160
         Y2              =   1680
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   7920
         X2              =   7680
         Y1              =   2640
         Y2              =   2400
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   7680
         X2              =   7080
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "To Evap."
         Height          =   255
         Left            =   6480
         TabIndex        =   98
         Top             =   5520
         Width           =   735
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   5520
         Picture         =   "Main_Turbines.frx":1E7E
         Stretch         =   -1  'True
         Top             =   5100
         Width           =   255
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "35# Dump toTo Main Condenser"
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
         Left            =   5400
         TabIndex        =   97
         Top             =   5400
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   39
         X1              =   5640
         X2              =   6180
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   38
         X1              =   6180
         X2              =   6180
         Y1              =   3720
         Y2              =   5040
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Auxiliary  Exhaust from Turbo feed pumps and Cargo pumps"
         Height          =   1395
         Index           =   1
         Left            =   7440
         TabIndex        =   92
         Top             =   5160
         Width           =   915
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   6840
         Picture         =   "Main_Turbines.frx":2188
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   36
         X1              =   6960
         X2              =   6960
         Y1              =   3720
         Y2              =   5280
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Throttle (Astn.)"
         Height          =   315
         Left            =   1080
         TabIndex        =   76
         Top             =   8160
         Width           =   1695
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Throttle (Ahd.)"
         Height          =   255
         Left            =   1200
         TabIndex        =   75
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Warming_label 
         BackStyle       =   0  'Transparent
         Caption         =   " Warming through valve"
         Height          =   195
         Left            =   1080
         TabIndex        =   74
         Top             =   0
         Width           =   1875
      End
      Begin VB.Line Warm_line 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   2700
         X2              =   2700
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Warm_line 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   1440
         X2              =   1440
         Y1              =   360
         Y2              =   780
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   1440
         X2              =   1800
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   1260
         TabIndex        =   70
         Top             =   5880
         Width           =   195
      End
      Begin VB.Image Image8 
         Height          =   360
         Left            =   -600
         Picture         =   "Main_Turbines.frx":2492
         Stretch         =   -1  'True
         Top             =   10800
         Width           =   285
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   34
         X1              =   8700
         X2              =   10200
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   33
         X1              =   8700
         X2              =   8700
         Y1              =   2760
         Y2              =   3240
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   32
         X1              =   10200
         X2              =   10200
         Y1              =   2760
         Y2              =   3240
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   4860
         Picture         =   "Main_Turbines.frx":279C
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   2700
         Picture         =   "Main_Turbines.frx":2AA6
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   2700
         Picture         =   "Main_Turbines.frx":2DB0
         Stretch         =   -1  'True
         Top             =   640
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   3900
         X2              =   4800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   4860
         Picture         =   "Main_Turbines.frx":30BA
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   255
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4980
         X2              =   4980
         Y1              =   2280
         Y2              =   6660
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   9720
         Picture         =   "Main_Turbines.frx":33C4
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Index           =   19
         X1              =   9180
         X2              =   10080
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   5280
         Picture         =   "Main_Turbines.frx":36CE
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Cascade"
         Height          =   255
         Left            =   4140
         TabIndex        =   62
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HP Bleed"
         Height          =   255
         Left            =   2700
         TabIndex        =   61
         Tag             =   "H"
         Top             =   3300
         Width           =   1035
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   1680
         X2              =   2520
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   29
         X1              =   5820
         X2              =   5820
         Y1              =   2760
         Y2              =   2280
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   5160
         Picture         =   "Main_Turbines.frx":39D8
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   28
         X1              =   4980
         X2              =   5820
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   7440
         Picture         =   "Main_Turbines.frx":3CE2
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   255
      End
      Begin VB.Image Image20 
         Height          =   255
         Left            =   7860
         Picture         =   "Main_Turbines.frx":3FEC
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Bleed"
         Height          =   315
         Left            =   5280
         TabIndex        =   60
         Top             =   2040
         Width           =   915
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   30
         X1              =   7560
         X2              =   7560
         Y1              =   3720
         Y2              =   5100
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "875/35 # Regulator"
         Height          =   255
         Left            =   8940
         TabIndex        =   59
         Tag             =   "35 "
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   4380
         TabIndex        =   57
         Top             =   2280
         Width           =   195
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   26
         X1              =   7860
         X2              =   7860
         Y1              =   2760
         Y2              =   3360
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   6360
         X2              =   6360
         Y1              =   2760
         Y2              =   3360
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IP Regulator"
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
         Left            =   6600
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   4200
         TabIndex        =   54
         Top             =   3240
         Width           =   195
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   27
         X1              =   6360
         X2              =   7860
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   3600
         Picture         =   "Main_Turbines.frx":42F6
         Stretch         =   -1  'True
         Top             =   5340
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   1020
         X2              =   1020
         Y1              =   6480
         Y2              =   780
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   4020
         Picture         =   "Main_Turbines.frx":4600
         Stretch         =   -1  'True
         Top             =   8880
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HP Bleed Regulator"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   4080
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   1680
         X2              =   1680
         Y1              =   6480
         Y2              =   7080
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   23
         X1              =   1020
         X2              =   3000
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   22
         X1              =   1020
         X2              =   3360
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   2520
         X2              =   2520
         Y1              =   7080
         Y2              =   6480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   20
         X1              =   240
         X2              =   1020
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ahead Throttle"
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
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LP Bleed"
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
         Left            =   4260
         TabIndex        =   24
         Top             =   9540
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image9 
         Height          =   360
         Left            =   8940
         Picture         =   "Main_Turbines.frx":490A
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   285
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   1920
         Picture         =   "Main_Turbines.frx":4C14
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "150 psi"
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
         Left            =   1680
         TabIndex        =   23
         Top             =   4620
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   3720
         X2              =   3720
         Y1              =   5040
         Y2              =   5400
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   7980
         X2              =   7980
         Y1              =   3720
         Y2              =   5100
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To Condenser"
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
         Left            =   2880
         TabIndex        =   21
         Top             =   9660
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "12 psi"
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
         Left            =   5220
         TabIndex        =   20
         Top             =   9000
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   5880
         Picture         =   "Main_Turbines.frx":4F1E
         Stretch         =   -1  'True
         Top             =   9300
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To LP Heater"
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
         Left            =   4920
         TabIndex        =   19
         Top             =   9180
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   4140
         X2              =   6180
         Y1              =   9420
         Y2              =   9420
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   15
         X1              =   4140
         X2              =   4140
         Y1              =   9420
         Y2              =   8160
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To Feed Pumps"
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
         Left            =   9000
         TabIndex        =   16
         Top             =   5760
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image15 
         Height          =   315
         Left            =   8820
         Picture         =   "Main_Turbines.frx":5228
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Index           =   14
         X1              =   8940
         X2              =   8940
         Y1              =   6840
         Y2              =   5640
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "De-Aerator"
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
         Left            =   8580
         TabIndex        =   14
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "35 psi"
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
         Left            =   7860
         TabIndex        =   13
         Top             =   3180
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   8400
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   8400
         Picture         =   "Main_Turbines.frx":5532
         Stretch         =   -1  'True
         Top             =   3420
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   13
         X1              =   8520
         X2              =   8520
         Y1              =   4200
         Y2              =   2760
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   2400
         X2              =   2400
         Y1              =   1140
         Y2              =   780
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   8
         X1              =   240
         X2              =   1020
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Warm_line 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   1440
         X2              =   2700
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "To HP Heater"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   5580
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   2100
         X2              =   4440
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   4440
         X2              =   4440
         Y1              =   3600
         Y2              =   5040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   3900
         X2              =   4440
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   3300
         Picture         =   "Main_Turbines.frx":583C
         Stretch         =   -1  'True
         Top             =   9420
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5010
         Top             =   7260
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2680
         Top             =   1200
         Width           =   675
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
         Left            =   3480
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5040
         Top             =   1200
         Width           =   555
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
         Left            =   3540
         TabIndex        =   1
         Top             =   7200
         Width           =   1035
      End
      Begin VB.Shape Shape28 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2565
         Top             =   7260
         Width           =   435
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3000
         X2              =   3000
         Y1              =   6480
         Y2              =   6660
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3420
         X2              =   3420
         Y1              =   8820
         Y2              =   9420
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3360
         X2              =   3360
         Y1              =   780
         Y2              =   840
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   735
         Left            =   5040
         TabIndex        =   0
         Top             =   3000
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   8580
         Shape           =   2  'Oval
         Top             =   3780
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   8280
         Shape           =   3  'Circle
         Top             =   4620
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Guardian v/v"
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
         TabIndex        =   22
         Top             =   6960
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1965
         Left            =   3300
         OleObjectBlob   =   "Main_Turbines.frx":5B46
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\My Documents\visiotbne.vsd"
         TabIndex        =   45
         Top             =   360
         Width           =   1755
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3345
         Left            =   2940
         OleObjectBlob   =   "Main_Turbines.frx":D35E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   46
         Top             =   6000
         Width           =   2235
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Astern Throttle"
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
         Left            =   1800
         TabIndex        =   10
         Top             =   5700
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Index           =   41
         X1              =   9240
         X2              =   8760
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   5160
         X2              =   6600
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   37
         X1              =   5640
         X2              =   5640
         Y1              =   4800
         Y2              =   5340
      End
      Begin VB.Label VV_pos_throttle_ahead 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2100
         TabIndex        =   130
         Top             =   480
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   42
         X1              =   3900
         X2              =   3900
         Y1              =   4740
         Y2              =   5040
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   5520
         OleObjectBlob   =   "Main_Turbines.frx":14D76
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   95
         Top             =   4200
         Width           =   375
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   5520
         OleObjectBlob   =   "Main_Turbines.frx":1D58E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   101
         Top             =   4200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5520
         OleObjectBlob   =   "Main_Turbines.frx":253A6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   100
         Top             =   4200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   6900
         OleObjectBlob   =   "Main_Turbines.frx":2D3BE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   84
         Top             =   2400
         Width           =   315
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   6900
         OleObjectBlob   =   "Main_Turbines.frx":359D6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   87
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   6900
         OleObjectBlob   =   "Main_Turbines.frx":3D7EE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   85
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   6900
         OleObjectBlob   =   "Main_Turbines.frx":45A06
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   86
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5520
         OleObjectBlob   =   "Main_Turbines.frx":4E41E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   99
         Top             =   4200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3780
         OleObjectBlob   =   "Main_Turbines.frx":56436
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   80
         Top             =   4140
         Width           =   375
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   3780
         OleObjectBlob   =   "Main_Turbines.frx":5EC4E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   83
         Top             =   4140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3780
         OleObjectBlob   =   "Main_Turbines.frx":66A66
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   82
         Top             =   4140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   3780
         OleObjectBlob   =   "Main_Turbines.frx":6EA7E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   81
         Top             =   4140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   9240
         OleObjectBlob   =   "Main_Turbines.frx":76A96
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   88
         Top             =   2400
         Width           =   315
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   9240
         OleObjectBlob   =   "Main_Turbines.frx":7F0AE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   91
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   9240
         OleObjectBlob   =   "Main_Turbines.frx":86EC6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   90
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   9240
         OleObjectBlob   =   "Main_Turbines.frx":8F0DE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   89
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   108
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Main_Turbines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub AUTO_guardian_VV_Click()

set_dummy_focus = True

' 17/10/99
    Call setPCDI(304)

End Sub

Private Sub Command1_Click()


' 13/07/09
'Feed_Water.Show
Feed_Water_Alone.Show

End Sub

Private Sub Command10_Click()

Main_Condenser.Show

End Sub

Private Sub Command2_Click()

' 16/11/99
'Boiler_Stbd.Show
Stbd_Boiler.Show

End Sub

Private Sub Command22_Click()


End Sub

Private Sub Command3_Click()

Condensate.Show

End Sub

Private Sub Command4_Click()

' 16/11/99
'BOILER_PORT.Show
Port_Boiler.Show


End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
Dim i As Integer

On Error Resume Next
Call JYAScreenInit(Me, idMain_Turbines)


' === 13/12/02 17.062 ======
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    For i = 0 To 6
        Steam_pipe(i).Visible = False
    Next i
    Valve(9).Visible = False
    Valve(10).Visible = False
Else
    For i = 0 To 6
        Steam_pipe(i).Visible = True
    Next i
    Valve(9).Visible = True
    Valve(10).Visible = True
End If
' === 13/12/02 17.062 end ===


' === 25/09/02 17.044 ======
If (TARGET_this_sim = CALIFORNIA) Then
    If I_AM_INSTRUCTOR Then
        Valve(90).Visible = True
        Warm_line(0).Visible = True
        Warm_line(1).Visible = True
        Warm_line(2).Visible = True
        Warming_label.Visible = True
    Else
        Valve(90).Visible = False
        Warm_line(0).Visible = False
        Warm_line(1).Visible = False
        Warm_line(2).Visible = False
        Warming_label.Visible = False
    End If
End If
' === 25/09/02 17.044 end ===


' === 14/11/02 =======
update_VBAI(17) = True
update_VBAI(18) = True


Reg_IP(1).Top = Reg_IP(0).Top
Reg_IP(2).Top = Reg_IP(0).Top
Reg_IP(3).Top = Reg_IP(0).Top

Reg_IP(1).Left = Reg_IP(0).Left
Reg_IP(2).Left = Reg_IP(0).Left
Reg_IP(3).Left = Reg_IP(0).Left


' 1/12/99
HScroll1.Visible = PTT
HScroll2.Visible = PTT


' 21/9/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    VV_pos_throttle_ahead.Visible = False
    VV_pos_throttle_astern.Visible = False
Else
    VV_pos_throttle_ahead.Visible = True
    VV_pos_throttle_astern.Visible = True
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idMain_Turbines)
    HavenStretch1.Visible = True
End If
   
End Sub

Private Sub HScroll1_Change()
Dim AIvalue As Integer

V_ahead_pos = HScroll1.Value * 0.1

' 7/7/99

' 1/12/99
'Call set_VSCROLL(3, HScroll1.Value)

AIvalue = HScroll1.Value / 4
Call setAI(17, AIvalue)


End Sub

Private Sub HScroll2_Change()
Dim AIvalue As Integer


' 1/12/99
AIvalue = HScroll2.Value / 4
Call setAI(18, AIvalue)


End Sub

Private Sub Mollier_Command_Click()

Mollier.Show

End Sub

Private Sub OFFpump_Click(Index As Integer)

' 13/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

Call setPCDI(291)

End Sub

Private Sub ONpump_Click(Index As Integer)

' 13/9/00
set_dummy_focus = True

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

' 13/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

Call setPCDI(292)

End Sub

Private Sub Timer1_Timer()
Dim green_colour As Integer
Dim i As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

Text1.text = HScroll4.Value


' === 14/11/02 =======
If update_VBAI(17) Then
     update_VBAI(17) = False
     
     ' === 19/02/03 ===
    ' HScroll1.Value = VB_AI(17) * 4
      Call update_the_slider(HScroll1, VB_AI(17) * 4)

End If
If update_VBAI(18) Then
     update_VBAI(18) = False
     
     'HScroll2.Value = VB_AI(18) * 4
      Call update_the_slider(HScroll2, VB_AI(18) * 4)

End If


' 12/5/00
' 7/12/00
'Call paint_alarm(Me.Tsteam14, 89, 0, 0, 0)
Call paint_alarm(Me.Tsteam_LP_out, 89, 0, 0, 0)


' 13/9/00
'    equivalence(STEAM_logicals(137),SHAFT_rollover_ON)
'
'    equivalence(STEAM_logicals(138),SHAFT_rollover_AUTO)
If VB_steam_logicals(137) Then
  ONpump(0).BackColor = &HFF00&     ' green
  OFFpump(0).BackColor = &HC0C0C0   ' grey
Else
  ONpump(0).BackColor = &HC0C0C0   ' grey
  OFFpump(0).BackColor = &H8080FF   ' pink
End If
If VB_steam_logicals(138) Then
  sbyPUMP(0).BackColor = &HFFFF&    ' yellow
Else
  sbyPUMP(0).BackColor = &HC0C0C0   ' grey
End If



If Advanced_mode Then

        ' === 22/11/02 17.034 ===
        'If Not I_AM_INSTRUCTOR And Not PTT Then
        '    ' I MUST be an ACT Terminal (Full Mission) so
        '    Mollier_Command.Visible = False
        'Else
            Mollier_Command.Visible = True
        'End If

  Turbine_Power_frame.Visible = True

  Hsteam10.Visible = True
  Hsteam11.Visible = True
  Hsteam12.Visible = True
  Hsteam13.Visible = True
  
 ' Hsteam14.Visible = True
   Hsteam_LP_out.Visible = True
  
  Fsteam10.Visible = True
  Fsteam11.Visible = True
  Fsteam12.Visible = True
  
  'Fsteam13.Visible = True
   Fsteam_LP_out.Visible = True
  
  Fsteam14.Visible = True
  Fsteam15.Visible = True
  Fsteam16.Visible = True
  
    For i = 0 To 6
      Line4(i).Visible = True
    Next i
    
    Power_label(0).Caption = Format(LSS_reals(797), "#####0.0") & " kW"
    Power_label(1).Caption = Format(LSS_reals(798), "#####0.0") & " kW"
    Power_label(2).Caption = Format(LSS_reals(799), "#####0.0") & " kW"
    Power_label(3).Caption = Format(LSS_reals(800), "#####0.0") & " kW"
      
            Fsteam(10) = LSS_reals(88)
            Fsteam(11) = LSS_reals(89)
            Fsteam(12) = LSS_reals(157)
            Fsteam(13) = LSS_reals(158)
            Fsteam(14) = LSS_reals(159)
      
      
      ' 20/10/99
   If SI_units Then
      Fsteam10.text = Format(LSS_reals(88) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
      Fsteam11.text = Format(LSS_reals(89) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
      Fsteam12.text = Format(LSS_reals(157) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
     
     ' 7/12/00
    ' Fsteam13.text = Format(LSS_reals(158) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
      Fsteam_LP_out.text = Format(LSS_reals(158) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
      
     Fsteam14.text = Format(LSS_reals(159) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
      
      ' 5/5/00
     ' Fsteam15.text = Format(LSS_reals(240) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
       Fsteam15.text = Format(LSS_reals(296) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
      
      Fsteam16.text = Format(LSS_reals(471) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
    Else
      Fsteam10.text = Format(LSS_reals(88) * Flow_mult / 3600#, "###0.0") & " lb/s"
      Fsteam11.text = Format(LSS_reals(89) * Flow_mult / 3600#, "###0.0") & " lb/s"
      Fsteam12.text = Format(LSS_reals(157) * Flow_mult / 3600#, "###0.0") & " lb/s"
      
     ' 7/12/00
    ' Fsteam13.text = Format(LSS_reals(158) * Flow_mult / 3600#, "###0.0") & " lb/s"
      Fsteam_LP_out.text = Format(LSS_reals(158) * Flow_mult / 3600#, "###0.0") & " lb/s"
      
     Fsteam14.text = Format(LSS_reals(159) * Flow_mult / 3600#, "###0.0") & " lb/s"
      
    '  Fsteam15.text = Format(LSS_reals(240) * Flow_mult / 3600#, "###0.0") & " lb/s"
       Fsteam15.text = Format(LSS_reals(296) * Flow_mult / 3600#, "###0.0") & " lb/s"
      
      Fsteam16.text = Format(LSS_reals(471) * Flow_mult / 3600#, "###0.0") & " lb/s"
    End If
  
    
    Hsteam10.text = Format(LSS_reals(65), "##0") & " kJ/kg"
    Hsteam11.text = Format(LSS_reals(66), "##0") & " kJ/kg"
    Hsteam12.text = Format(LSS_reals(67), "##0") & " kJ/kg"
    Hsteam13.text = Format(LSS_reals(68), "##0") & " kJ/kg"
    
    ' 7/12/00
    'Hsteam14.text = Format(LSS_reals(69), "##0") & " kJ/kg"
     Hsteam_LP_out.text = Format(LSS_reals(69), "##0") & " kJ/kg"

    
Else

  Mollier_Command.Visible = False

  Turbine_Power_frame.Visible = False
  
  Hsteam10.Visible = False
  Hsteam11.Visible = False
  Hsteam12.Visible = False
  Hsteam13.Visible = False
  
  ' 7/12/00
  ' Hsteam14.Visible = False
    Hsteam_LP_out.Visible = False
  
  Fsteam10.Visible = False
  Fsteam11.Visible = False
  Fsteam12.Visible = False
  
 ' Fsteam13.Visible = False
   Fsteam_LP_out.Visible = False
  
  Fsteam14.Visible = False
  Fsteam15.Visible = False
  Fsteam16.Visible = False

  
    For i = 0 To 6
      Line4(i).Visible = False
    Next i
                
End If



'  12/11/99
' PC_REALS(81) = V_ahead_hand_pos ( 0 to 100.0 )
If LSS_reals(81) > 99# Then
    green_colour = 255
ElseIf LSS_reals(81) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(81) * 2.55
End If

' 11/10/99   AHEAD HAND-WHEEL Throttling valve...
Valve(9).ColourClosed = RGB(0, green_colour, 0)


' 11/2/00''''''''''''

' PC_REALS(472) = V_astern_hand_pos ( 0 to 100.0 )
If LSS_reals(472) > 99# Then
    green_colour = 255
ElseIf LSS_reals(472) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(472) * 2.55
End If

' ASTERN HAND-WHEEL Throttling valve...
Valve(10).ColourClosed = RGB(0, green_colour, 0)

' '''''''''''''''''''



' 12/11/99  AHEAD MAIN THROTTLE VALVE...
'   PC_REALS(2) = V_ahead_throttle_pos   ( 0 to 100.0 )
VV_pos_throttle_ahead.Caption = Format(LSS_reals(2), "##0") & " %"
If LSS_reals(2) > 99# Then
    green_colour = 255
ElseIf LSS_reals(2) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(2) * 2.55
End If
Valve(0).ColourClosed = RGB(0, green_colour, 0)


' 26/11/99  ASTERN MAIN THROTTLE VALVE...
' PC_REALS(1) = V_astern_throttle_pos
VV_pos_throttle_astern.Caption = Format(LSS_reals(1), "##0") & " %"
If LSS_reals(1) > 99# Then
    green_colour = 255
ElseIf LSS_reals(1) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(1) * 2.55
End If
Valve(11).ColourClosed = RGB(0, green_colour, 0)




' Auto Guardian valve
If (D_OUTPUT(1053)) Then
  AUTO_guardian_VV.BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_guardian_VV.BackColor = &HC0C0C0           ' GREY
End If



If SI_units Then
    ' 11/2/00
    Psteam6.text = Format(LSS_reals(788) * Press_mult, "##0.0") & Press_chars
    
    Psteam10.text = Format(LSS_reals(60) * Press_mult, "##0.0") & Press_chars
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0.0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0.0") & Press_chars
    Psteam13.text = Format(LSS_reals(63) * Press_mult, "##0.00") & Press_chars
    Psteam14.text = Format(LSS_reals(64) * Press_mult, "##0.000") & Press_chars
Else

    ' 11/2/00
    Psteam6.text = Format(LSS_reals(788) * Press_mult, "##0") & Press_chars
    
    Psteam10.text = Format(LSS_reals(60) * Press_mult, "##0") & Press_chars
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0") & Press_chars
    Psteam13.text = Format(LSS_reals(63) * Press_mult, "##0.0") & Press_chars
    Psteam14.text = Format(LSS_reals(64) * Press_mult, "##0.00") & Press_chars
End If

Tsteam10.text = Format(LSS_reals(70) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam11.text = Format(LSS_reals(71) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam12.text = Format(LSS_reals(72) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam13.text = Format(LSS_reals(73) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' 7/12/00
'Tsteam14.text = Format(LSS_reals(74) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam_LP_out.text = Format(LSS_reals(74) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' update_vvalve_LSS(syxxx,nValve,index1,index2)

' 16/9/99 old LSS...
'Call update_vvalve_LSS(Me, 16, 4, -1)    ' VV16
'Call update_vvalve_LSS(Me, 18, 3, -1)    ' VV18
  

   Call update_vvalve_LSS(Me, 112, 1, -1)
   Call update_vvalve_LSS(Me, 282, 2, -1)
   Call update_vvalve_LSS(Me, 283, 3, -1)
   Call update_vvalve_LSS(Me, 284, 4, -1)
   Call update_vvalve_LSS(Me, 98, 5, -1)
   Call update_vvalve_LSS(Me, 110, 6, -1)
   Call update_vvalve_LSS(Me, 287, 7, -1)
   Call update_vvalve_LSS(Me, 108, 8, -1)

   Call update_vvalve_LSS(Me, 111, 13, -1)
   Call update_vvalve_LSS(Me, 76, 14, -1)
   Call update_vvalve_LSS(Me, 92, 15, -1)
   
   Call update_vvalve_LSS(Me, 74, 19, -1)
   Call update_vvalve_LSS(Me, 79, 20, -1)
   Call update_vvalve_LSS(Me, 107, 21, -1)

   Call update_vvalve_LSS(Me, 77, 23, -1)
   Call update_vvalve_LSS(Me, 94, 24, -1)
   Call update_vvalve_LSS(Me, 81, 25, -1)
   
  ' 17/10/99
   Call update_vvalve_LSS(Me, 293, 90, -1)
   
   
   
'   === 11/11/99 By-Pass Valves ===

'    PC_REALS(367) = Eight75_35_bypass * 100#
Call update_green_VV(367, Valve(22))

' PC_REALS(369) = Dump_35_bypass * 100#
Call update_green_VV(369, Valve(26))

'    PC_REALS(373) = IP_bypass * 100#
Call update_green_VV(373, Valve(18))

' 11/11/99 Added
'    PC_REALS(371) = HP_bypass * 100#
Call update_green_VV(371, Valve(17))


'  === 11/11/99 Regulator  Valves ===

Call update_regulator(366, Reg_875_35(0), Reg_875_35(1), Reg_875_35(2), Reg_875_35(3))
Call update_regulator(368, Dump_35(0), Dump_35(1), Dump_35(2), Dump_35(3))
Call update_regulator(372, Reg_IP(0), Reg_IP(1), Reg_IP(2), Reg_IP(3))

' 11/11/99 Added
Call update_regulator(370, Reg_HP(0), Reg_HP(1), Reg_HP(2), Reg_HP(3))


End Sub

Private Sub Tsteam14_Click()

 ' '12/5/00
 '   Unload Alarm_set_point   ' Do this first
 '   nAlarm_HI = 89
 '
 '   Alarm_set_point.Show

End Sub

Private Sub Tsteam_LP_out_Click()

  '7/12/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 89
    
    Alarm_set_point.Show


End Sub

Private Sub valve_Click(Index As Integer)



 ' 7/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        
    '    LP
    If Index = 7 Then
        Exit Sub
    End If
    
End If

 
' See Main Turbines.xls

Select Case Index

  Case 1
    Call valve_pcdi_LSS(112)
  Case 2
    Call valve_pcdi_LSS(282)
  Case 3
    Call valve_pcdi_LSS(283)
  Case 4
    Call valve_pcdi_LSS(284)
  Case 5
    Call valve_pcdi_LSS(98)
  Case 6
    Call valve_pcdi_LSS(110)
  Case 7
    Call valve_pcdi_LSS(287)
  Case 8
    Call valve_pcdi_LSS(108)
  Case 13
    Call valve_pcdi_LSS(111)
  Case 14
    Call valve_pcdi_LSS(76)
  Case 15
    Call valve_pcdi_LSS(92)
  Case 19
    Call valve_pcdi_LSS(74)
  Case 20
    Call valve_pcdi_LSS(79)
  Case 21
    Call valve_pcdi_LSS(107)
  Case 23
    Call valve_pcdi_LSS(77)
  Case 24
    Call valve_pcdi_LSS(94)
  Case 25
    Call valve_pcdi_LSS(81)
    
    ' 17/10/99 Warming through V/V
  Case 90
    Call valve_pcdi_LSS(293)
    
    
End Select



End Sub

Sub update_green_VV(nReals As Integer, obj As Object)
Dim green_colour As Integer

' 6/11/99
If LSS_reals(nReals) > 99# Then
    green_colour = 255
ElseIf LSS_reals(nReals) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(nReals) * 2.55
End If

obj.ColourClosed = RGB(0, green_colour, 0)

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
Sub update_the_slider(Slider As Object, Pot_Value As Integer)

Dim S_max As Integer
Dim S_min As Integer

' TESTING...
S_max = Slider.max
S_min = Slider.Min

' === 19/02/03 ===
If Pot_Value >= Slider.max Then
    Slider.Value = Slider.max
ElseIf Pot_Value <= Slider.Min Then
    Slider.Value = Slider.Min
Else
    Slider.Value = Pot_Value
End If


End Sub


