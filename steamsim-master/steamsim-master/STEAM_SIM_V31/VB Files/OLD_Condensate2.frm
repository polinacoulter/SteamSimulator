VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.0#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form OLD_Condensate 
   Caption         =   "Condensate"
   ClientHeight    =   10845
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10845
   ScaleWidth      =   15240
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   10215
      Left            =   60
      Top             =   60
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   18018
      Begin VB.VScrollBar VScroll6 
         Height          =   1335
         Left            =   6600
         Max             =   1000
         TabIndex        =   180
         Top             =   540
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
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
         Left            =   650
         TabIndex        =   179
         Text            =   "P"
         Top             =   1440
         Width           =   195
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Distiller"
         Height          =   675
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   8100
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Main Condenser"
         Height          =   675
         Left            =   13260
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   3780
         Width           =   915
      End
      Begin VB.TextBox Z_CDT 
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
         Left            =   13140
         TabIndex        =   176
         Text            =   "20 %"
         Top             =   7860
         Width           =   855
      End
      Begin VB.TextBox T_DWT 
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
         Left            =   1440
         TabIndex        =   168
         Text            =   "xx F"
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox T_DFT 
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
         Left            =   4080
         TabIndex        =   167
         Text            =   "xx F"
         Top             =   3180
         Width           =   795
      End
      Begin VB.TextBox T_LP_htr 
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
         TabIndex        =   166
         Text            =   "xx F"
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox T_cond_recirc 
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
         Left            =   5400
         TabIndex        =   165
         Text            =   "xx F"
         Top             =   6840
         Width           =   855
      End
      Begin VB.TextBox T_FWG 
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
         Left            =   6480
         TabIndex        =   164
         Text            =   "xx F"
         Top             =   7140
         Width           =   855
      End
      Begin VB.TextBox T_air_eject 
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
         Left            =   7620
         TabIndex        =   163
         Text            =   "xx F"
         Top             =   7140
         Width           =   855
      End
      Begin VB.TextBox T_ADT 
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
         Left            =   8160
         TabIndex        =   162
         Text            =   "xx F"
         Top             =   9360
         Width           =   855
      End
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
         Left            =   12060
         TabIndex        =   161
         Text            =   "xx F"
         Top             =   5400
         Width           =   855
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   1335
         Left            =   1200
         Max             =   1000
         TabIndex        =   160
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox F_DWT_dump 
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
         TabIndex        =   158
         Text            =   "xx.x kg/s"
         Top             =   7320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_DFT_dump 
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
         Left            =   1380
         TabIndex        =   157
         Text            =   "xx.x kg/s"
         Top             =   3300
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_DFT_main_IN 
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
         TabIndex        =   156
         Text            =   "xx.x kg/s"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_cond_htrs 
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
         Left            =   7560
         TabIndex        =   155
         Text            =   "xx.x kg/s"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_cond_recirc 
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
         Left            =   9660
         TabIndex        =   154
         Text            =   "xx.x kg/s"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_cond_recirc 
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
         Left            =   8700
         TabIndex        =   153
         Text            =   "xx.xkg/s"
         Top             =   5940
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_cond_recirc 
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
         Left            =   8400
         TabIndex        =   152
         Text            =   "xx.x kg/s"
         Top             =   4680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox F_DFT_hotwell 
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
         Left            =   6000
         TabIndex        =   151
         Text            =   "xx.x kg/s"
         Top             =   3420
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_hotwell_out 
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
         Left            =   11820
         TabIndex        =   150
         Text            =   "xx.x kg/s"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_ADT_out 
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
         Left            =   3600
         TabIndex        =   149
         Text            =   "xx.x kg/s"
         Top             =   8100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox PY_cond 
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
         TabIndex        =   148
         Text            =   "x.x Bar"
         Top             =   3960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1335
         Left            =   12960
         Max             =   1000
         TabIndex        =   146
         Top             =   4800
         Width           =   135
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   6060
         Max             =   1000
         TabIndex        =   132
         Top             =   3960
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   3540
         Max             =   100
         TabIndex        =   131
         Top             =   1020
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   1020
         Max             =   100
         TabIndex        =   130
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox P_ADT_out 
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
         Left            =   3420
         TabIndex        =   124
         Text            =   "x.x Bar"
         Top             =   9540
         Width           =   795
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1335
         Left            =   5040
         Max             =   1000
         TabIndex        =   122
         Top             =   2040
         Width           =   135
      End
      Begin VB.TextBox P_cond_out 
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
         Left            =   8940
         TabIndex        =   101
         Text            =   "x.x Bar"
         Top             =   6480
         Width           =   795
      End
      Begin VB.TextBox Z_DWT 
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
         Left            =   1440
         TabIndex        =   100
         Text            =   "20 %"
         Top             =   6060
         Width           =   855
      End
      Begin VB.TextBox Z_DFT 
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
         Left            =   4080
         TabIndex        =   99
         Text            =   "20 %"
         Top             =   2880
         Width           =   795
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Left            =   60
         Max             =   1000
         TabIndex        =   98
         Top             =   6360
         Width           =   135
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1035
         Left            =   480
         Max             =   1000
         TabIndex        =   97
         Top             =   3960
         Width           =   135
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
         Left            =   10680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   7920
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
         Left            =   10800
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   6360
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Index           =   1
         Left            =   2640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   9240
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
         Left            =   2640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   8640
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Left            =   10200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   7920
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
         Left            =   10320
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   6360
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
         Left            =   2100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   9240
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Left            =   2100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   8640
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Left            =   9720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   7920
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
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   9240
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Left            =   1560
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   8640
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Left            =   9840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   6360
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   540
         Top             =   9120
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Gland Seal"
         Height          =   495
         Left            =   6720
         TabIndex        =   54
         Top             =   5340
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Hotwell_tank 
         Height          =   705
         Left            =   11820
         TabIndex        =   9
         Top             =   4875
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   1244
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   13
         Left            =   9660
         TabIndex        =   34
         Top             =   5340
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
      Begin VB.CommandButton Command1 
         Caption         =   "Feed System"
         Height          =   675
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6480
         Width           =   795
      End
      Begin VB.TextBox Z_Hotwell 
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
         Left            =   12060
         TabIndex        =   1
         Text            =   "20 %"
         Top             =   5100
         Width           =   855
      End
      Begin VB.TextBox Z_ADT 
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
         Left            =   8160
         TabIndex        =   0
         Text            =   "20 %"
         Top             =   9000
         Width           =   855
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   3
         Left            =   10380
         TabIndex        =   3
         Top             =   7260
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   273
         Left            =   11280
         TabIndex        =   4
         Top             =   7200
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
         Index           =   272
         Left            =   9960
         TabIndex        =   5
         Top             =   6960
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
         Height          =   495
         Index           =   2
         Left            =   10320
         TabIndex        =   6
         Top             =   5820
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   271
         Left            =   11280
         TabIndex        =   7
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
         Index           =   270
         Left            =   9960
         TabIndex        =   8
         Top             =   5520
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
         Height          =   255
         Index           =   0
         Left            =   3780
         TabIndex        =   17
         Top             =   8640
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   1
         Left            =   3780
         TabIndex        =   18
         Top             =   9180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   80
         Left            =   4260
         TabIndex        =   19
         Top             =   3540
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
         Height          =   315
         Index           =   2
         Left            =   7560
         TabIndex        =   24
         Top             =   4260
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
         Index           =   3
         Left            =   8100
         TabIndex        =   25
         Top             =   5280
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
         Index           =   8
         Left            =   8460
         TabIndex        =   26
         Top             =   5280
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
         Index           =   9
         Left            =   8220
         TabIndex        =   27
         Top             =   4260
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
         Height          =   255
         Index           =   120
         Left            =   60
         TabIndex        =   40
         Top             =   900
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
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
         Height          =   435
         Index           =   111
         Left            =   1740
         TabIndex        =   41
         Top             =   80
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
         Index           =   112
         Left            =   1020
         TabIndex        =   42
         Top             =   80
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
         Index           =   113
         Left            =   1440
         TabIndex        =   43
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
         Index           =   98
         Left            =   4500
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
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
         Height          =   435
         Index           =   107
         Left            =   3540
         TabIndex        =   45
         Top             =   80
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
         Index           =   108
         Left            =   4290
         TabIndex        =   46
         Top             =   80
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
         Index           =   109
         Left            =   3960
         TabIndex        =   47
         Top             =   540
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
         Height          =   1275
         Index           =   302
         Left            =   3840
         TabIndex        =   48
         Top             =   2280
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   2249
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   11
         Left            =   3360
         TabIndex        =   52
         Top             =   8340
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   14
         Left            =   3360
         TabIndex        =   53
         Top             =   8880
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   77
         Left            =   180
         TabIndex        =   60
         Top             =   1860
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
         Index           =   81
         Left            =   180
         TabIndex        =   61
         Top             =   2520
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
         Index           =   91
         Left            =   840
         TabIndex        =   62
         Top             =   2280
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
         Height          =   315
         Index           =   22
         Left            =   9240
         TabIndex        =   64
         Top             =   7560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
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
         Index           =   274
         Left            =   11115
         TabIndex        =   65
         Top             =   6840
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   275
         Left            =   10740
         TabIndex        =   66
         Top             =   5520
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   15
         Left            =   900
         TabIndex        =   67
         Top             =   6480
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
         Left            =   900
         TabIndex        =   68
         Top             =   7140
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
         Left            =   240
         TabIndex        =   69
         Top             =   6720
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
         Height          =   375
         Index           =   10
         Left            =   4140
         TabIndex        =   71
         Top             =   8460
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
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
         Index           =   12
         Left            =   4140
         TabIndex        =   72
         Top             =   9000
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 ADT_tank 
         Height          =   855
         Left            =   4860
         TabIndex        =   2
         Top             =   8700
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1508
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   18
         Left            =   1320
         TabIndex        =   77
         Top             =   4020
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
         Left            =   1320
         TabIndex        =   78
         Top             =   4680
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
         Index           =   21
         Left            =   660
         TabIndex        =   79
         Top             =   4320
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
      Begin HavenValve1_2Prj.HavenValveV1_2 DWT_tank 
         Height          =   735
         Left            =   0
         TabIndex        =   80
         Top             =   5280
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1296
         TopColour       =   16711680
         BottomColour    =   16711680
         LineWidth       =   6
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   110
         Left            =   1800
         TabIndex        =   102
         Top             =   480
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
         Index           =   19
         Left            =   1800
         TabIndex        =   120
         Top             =   3840
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
         Index           =   24
         Left            =   1560
         TabIndex        =   121
         Top             =   7140
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
         Index           =   23
         Left            =   5820
         TabIndex        =   125
         Top             =   4380
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Height          =   435
         Index           =   1
         Left            =   6060
         TabIndex        =   126
         Top             =   4560
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
         Index           =   5
         Left            =   6480
         TabIndex        =   128
         Top             =   4020
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
         Index           =   4
         Left            =   6820
         TabIndex        =   127
         Top             =   4560
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
      Begin VB.TextBox PX_cond 
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
         Left            =   10020
         TabIndex        =   147
         Text            =   "x.x Bar"
         Top             =   9360
         Visible         =   0   'False
         Width           =   795
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 CDT_tank 
         Height          =   915
         Left            =   12180
         TabIndex        =   170
         Top             =   7260
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1614
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   88
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   583
         Left            =   11820
         TabIndex        =   174
         Top             =   7725
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
      Begin VB.Line PY_line 
         Index           =   1
         X1              =   5520
         X2              =   5760
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line PY_line 
         Index           =   0
         X1              =   5520
         X2              =   5520
         Y1              =   4200
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   14
         X1              =   7160
         X2              =   7160
         Y1              =   7560
         Y2              =   8820
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Caption         =   "Gland Exhaust Condenser"
         Height          =   675
         Left            =   6360
         TabIndex        =   183
         Top             =   7680
         Width           =   735
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Caption         =   "De-Aerator Flow Rate"
         Height          =   435
         Left            =   6300
         TabIndex        =   182
         Top             =   120
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label38 
         Caption         =   "0"
         Height          =   315
         Left            =   6840
         TabIndex        =   181
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   8280
         X2              =   8280
         Y1              =   5040
         Y2              =   4920
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   8280
         X2              =   8280
         Y1              =   4790
         Y2              =   4080
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   7920
         Picture         =   "OLD_Condensate2.frx":0000
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   41
         X1              =   10920
         X2              =   7080
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   8520
         X2              =   8520
         Y1              =   5700
         Y2              =   5040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   29
         X1              =   8160
         X2              =   7440
         Y1              =   5220
         Y2              =   5220
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   40
         X1              =   7740
         X2              =   7740
         Y1              =   5150
         Y2              =   5040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   35
         X1              =   7620
         X2              =   7620
         Y1              =   5220
         Y2              =   4920
      End
      Begin VB.Line Line26 
         BorderStyle     =   3  'Dot
         X1              =   720
         X2              =   540
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line19 
         BorderStyle     =   3  'Dot
         X1              =   720
         X2              =   720
         Y1              =   1740
         Y2              =   2340
      End
      Begin VB.Image Image25 
         Height          =   255
         Left            =   3780
         Picture         =   "OLD_Condensate2.frx":030A
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   78
         X1              =   3540
         X2              =   3540
         Y1              =   3540
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   77
         X1              =   3420
         X2              =   3900
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   76
         X1              =   3420
         X2              =   3420
         Y1              =   4950
         Y2              =   5520
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   15
         X1              =   7740
         X2              =   7740
         Y1              =   5880
         Y2              =   5340
      End
      Begin VB.Shape Sight_Glass 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   12930
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   135
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To Slop Tank"
         Height          =   495
         Left            =   13680
         TabIndex        =   175
         Top             =   8340
         Width           =   855
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   14
         Left            =   14040
         Picture         =   "OLD_Condensate2.frx":0614
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   75
         X1              =   14160
         X2              =   14160
         Y1              =   7440
         Y2              =   8340
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   74
         X1              =   13080
         X2              =   14160
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   73
         X1              =   11640
         X2              =   12240
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   70
         X1              =   11640
         X2              =   11640
         Y1              =   8040
         Y2              =   8520
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Sight Glass"
         Height          =   255
         Left            =   13100
         TabIndex        =   173
         Top             =   7440
         Width           =   855
      End
      Begin VB.Shape Shape41 
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   915
         Left            =   12900
         Shape           =   4  'Rounded Rectangle
         Top             =   7260
         Width           =   195
      End
      Begin VB.Image Image24 
         Height          =   255
         Left            =   9180
         Picture         =   "OLD_Condensate2.frx":091E
         Stretch         =   -1  'True
         Top             =   8400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   72
         X1              =   7980
         X2              =   7980
         Y1              =   8880
         Y2              =   8520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   71
         X1              =   7980
         X2              =   11640
         Y1              =   8520
         Y2              =   8520
      End
      Begin VB.Image Image23 
         Height          =   255
         Left            =   13200
         Picture         =   "OLD_Condensate2.frx":0C28
         Stretch         =   -1  'True
         Top             =   6960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "Steam Heating Returns"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   12960
         TabIndex        =   172
         Top             =   6360
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   69
         X1              =   12600
         X2              =   12600
         Y1              =   7500
         Y2              =   7080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   68
         X1              =   12600
         X2              =   13500
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "Contaminated Drain Tank"
         Height          =   255
         Left            =   11640
         TabIndex        =   171
         Top             =   8280
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   67
         X1              =   8520
         X2              =   7800
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Label Label32 
         Caption         =   "Misc returns"
         Height          =   255
         Left            =   8040
         TabIndex        =   169
         Top             =   8160
         Width           =   855
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   13
         Left            =   7680
         Picture         =   "OLD_Condensate2.frx":0F32
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   63
         X1              =   7800
         X2              =   7800
         Y1              =   8160
         Y2              =   8760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   66
         X1              =   3420
         X2              =   3420
         Y1              =   4680
         Y2              =   3580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   65
         X1              =   3900
         X2              =   3900
         Y1              =   5520
         Y2              =   7560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   64
         X1              =   3900
         X2              =   3900
         Y1              =   7800
         Y2              =   7980
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Hotwell"
         Height          =   255
         Left            =   10980
         TabIndex        =   159
         Top             =   4920
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   32
         X1              =   7440
         X2              =   5760
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label COND_recirc_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6840
         TabIndex        =   145
         Top             =   4380
         Width           =   555
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   10980
         TabIndex        =   139
         Top             =   4560
         Width           =   195
      End
      Begin VB.Shape Shape40 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   10920
         Top             =   4500
         Width           =   315
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   33
         X1              =   300
         X2              =   300
         Y1              =   1440
         Y2              =   1920
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   5640
         Picture         =   "OLD_Condensate2.frx":123C
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   34
         X1              =   7440
         X2              =   7440
         Y1              =   4860
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   46
         X1              =   8160
         X2              =   10020
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line18 
         X1              =   9660
         X2              =   9780
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   3420
         Picture         =   "OLD_Condensate2.frx":1546
         Stretch         =   -1  'True
         Top             =   7860
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   40
         X1              =   3300
         X2              =   3900
         Y1              =   7980
         Y2              =   7980
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   12
         Left            =   5280
         Picture         =   "OLD_Condensate2.frx":1850
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   62
         X1              =   5400
         X2              =   5400
         Y1              =   7560
         Y2              =   8760
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   8520
         X2              =   7740
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label28 
         Caption         =   "Air Ejector"
         Height          =   255
         Left            =   7680
         TabIndex        =   84
         Top             =   7640
         Width           =   855
      End
      Begin VB.Image Image20 
         Height          =   360
         Left            =   180
         Picture         =   "OLD_Condensate2.frx":1B5A
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   285
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   360
         Picture         =   "OLD_Condensate2.frx":1E64
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   1920
         Picture         =   "OLD_Condensate2.frx":216E
         Stretch         =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   3240
         Picture         =   "OLD_Condensate2.frx":2478
         Stretch         =   -1  'True
         Top             =   1500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   9360
         Picture         =   "OLD_Condensate2.frx":2782
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   30
         X1              =   10200
         X2              =   10200
         Y1              =   3360
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   17
         X1              =   9180
         X2              =   10200
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line9 
         X1              =   4800
         X2              =   5040
         Y1              =   5040
         Y2              =   4920
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   11
         Left            =   3120
         Picture         =   "OLD_Condensate2.frx":2A8C
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   255
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Distilled Water tank"
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   2160
         MousePointer    =   1  'Arrow
         TabIndex        =   81
         Top             =   5400
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   29
         X1              =   1440
         X2              =   3240
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line29 
         BorderStyle     =   3  'Dot
         X1              =   2460
         X2              =   2460
         Y1              =   2580
         Y2              =   4500
      End
      Begin VB.Line Line28 
         BorderStyle     =   3  'Dot
         X1              =   2400
         X2              =   1680
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   30
         X1              =   780
         X2              =   780
         Y1              =   3960
         Y2              =   5100
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1440
         X2              =   780
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1440
         X2              =   780
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Line23 
         BorderStyle     =   3  'Dot
         X1              =   3120
         X2              =   2400
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line22 
         BorderStyle     =   3  'Dot
         X1              =   3060
         X2              =   2880
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PTd"
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   5820
         Width           =   195
      End
      Begin VB.Shape Shape39 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2700
         Top             =   5760
         Width           =   315
      End
      Begin VB.Line Line21 
         BorderStyle     =   3  'Dot
         X1              =   2880
         X2              =   1260
         Y1              =   7020
         Y2              =   7020
      End
      Begin VB.Line Line17 
         BorderStyle     =   3  'Dot
         X1              =   2880
         X2              =   2880
         Y1              =   3240
         Y2              =   7020
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   3120
         TabIndex        =   75
         Top             =   2940
         Width           =   195
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   3120
         TabIndex        =   74
         Top             =   2280
         Width           =   195
      End
      Begin VB.Image Image18 
         Height          =   360
         Left            =   6600
         Picture         =   "OLD_Condensate2.frx":2D96
         Stretch         =   -1  'True
         Top             =   8400
         Width           =   285
      End
      Begin VB.Line Line16 
         BorderStyle     =   3  'Dot
         X1              =   4020
         X2              =   3900
         Y1              =   8940
         Y2              =   8820
      End
      Begin VB.Line Line15 
         BorderStyle     =   3  'Dot
         X1              =   4020
         X2              =   3960
         Y1              =   8940
         Y2              =   9180
      End
      Begin VB.Line Line14 
         BorderStyle     =   3  'Dot
         X1              =   4620
         X2              =   4080
         Y1              =   8940
         Y2              =   8940
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   4680
         TabIndex        =   73
         Top             =   8880
         Width           =   195
      End
      Begin VB.Shape Shape31 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4620
         Top             =   8820
         Width           =   315
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   2760
         Picture         =   "OLD_Condensate2.frx":30A0
         Stretch         =   -1  'True
         Top             =   7620
         Width           =   375
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   10
         Left            =   4800
         Picture         =   "OLD_Condensate2.frx":33AA
         Stretch         =   -1  'True
         Top             =   8100
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   9
         Left            =   900
         Picture         =   "OLD_Condensate2.frx":36B4
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   4920
         X2              =   4920
         Y1              =   7740
         Y2              =   8820
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   1020
         X2              =   4920
         Y1              =   7740
         Y2              =   7740
      End
      Begin VB.Label Label7 
         Caption         =   "Make up Feed Water to ADT"
         Height          =   255
         Left            =   660
         TabIndex        =   70
         Top             =   7800
         Width           =   2295
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1020
         X2              =   360
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1020
         X2              =   360
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   360
         X2              =   360
         Y1              =   6360
         Y2              =   7560
      End
      Begin VB.Line Line13 
         BorderStyle     =   3  'Dot
         X1              =   11220
         X2              =   10860
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Shape Shape19 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   10860
         Shape           =   3  'Circle
         Top             =   4800
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   5760
         X2              =   6120
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   8
         Left            =   7920
         Picture         =   "OLD_Condensate2.frx":39BE
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   8040
         X2              =   8040
         Y1              =   4080
         Y2              =   3420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   8880
         X2              =   9780
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   24
         X1              =   9780
         X2              =   9780
         Y1              =   7140
         Y2              =   7800
      End
      Begin VB.Line Line7 
         X1              =   3120
         X2              =   3540
         Y1              =   5100
         Y2              =   4860
      End
      Begin VB.Line Line8 
         X1              =   4800
         X2              =   5040
         Y1              =   4800
         Y2              =   4680
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3540
         Y1              =   4860
         Y2              =   4620
      End
      Begin VB.Label Label30 
         Caption         =   "To Main Condenser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         TabIndex        =   63
         Top             =   3360
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   180
         Picture         =   "OLD_Condensate2.frx":3CC8
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   38
         X1              =   960
         X2              =   960
         Y1              =   1440
         Y2              =   3120
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   39
         X1              =   300
         X2              =   960
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   1200
         Picture         =   "OLD_Condensate2.frx":3FD2
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   36
         X1              =   1320
         X2              =   1320
         Y1              =   1440
         Y2              =   1860
      End
      Begin VB.Label Label37 
         Caption         =   "To Evap."
         Height          =   435
         Left            =   1440
         TabIndex        =   59
         Top             =   1680
         Width           =   435
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   315
         Left            =   7620
         OleObjectBlob   =   "OLD_Condensate2.frx":42DC
         SizeMode        =   1  'Stretch
         TabIndex        =   58
         Top             =   5970
         Width           =   195
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   315
         Left            =   7350
         OleObjectBlob   =   "OLD_Condensate2.frx":5EF4
         SizeMode        =   1  'Stretch
         TabIndex        =   57
         Top             =   6000
         Width           =   195
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   495
         Left            =   8040
         OleObjectBlob   =   "OLD_Condensate2.frx":7B0C
         SizeMode        =   1  'Stretch
         TabIndex        =   56
         Top             =   5790
         Width           =   255
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   495
         Left            =   8400
         OleObjectBlob   =   "OLD_Condensate2.frx":9724
         SizeMode        =   1  'Stretch
         TabIndex        =   55
         Top             =   5790
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   3
         Left            =   5220
         Picture         =   "OLD_Condensate2.frx":B33C
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Astern Turbine"
         Height          =   495
         Left            =   10260
         TabIndex        =   32
         Top             =   240
         Width           =   555
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   10740
         Picture         =   "OLD_Condensate2.frx":B646
         Stretch         =   -1  'True
         Top             =   1980
         Width           =   255
      End
      Begin VB.Line Line38 
         BorderStyle     =   3  'Dot
         X1              =   5340
         X2              =   5340
         Y1              =   2820
         Y2              =   5220
      End
      Begin VB.Line Line37 
         BorderStyle     =   3  'Dot
         X1              =   11520
         X2              =   5340
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Line Line36 
         BorderStyle     =   3  'Dot
         X1              =   11520
         X2              =   11520
         Y1              =   2040
         Y2              =   2820
      End
      Begin VB.Line Line10 
         BorderStyle     =   3  'Dot
         X1              =   10860
         X2              =   10860
         Y1              =   6060
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   58
         X1              =   3780
         X2              =   3960
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   57
         X1              =   3540
         X2              =   3960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "DFT Water level"
         Height          =   615
         Left            =   4080
         TabIndex        =   49
         Top             =   2220
         Width           =   555
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   4920
         Picture         =   "OLD_Condensate2.frx":B950
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   23
         X1              =   4800
         X2              =   3360
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   25
         X1              =   840
         X2              =   840
         Y1              =   900
         Y2              =   300
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   24
         X1              =   2280
         X2              =   2280
         Y1              =   900
         Y2              =   360
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   20
         X1              =   2280
         X2              =   840
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   21
         X1              =   3360
         X2              =   3360
         Y1              =   840
         Y2              =   360
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   22
         X1              =   4800
         X2              =   4800
         Y1              =   840
         Y2              =   360
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   2700
         Picture         =   "OLD_Condensate2.frx":BC5A
         Stretch         =   -1  'True
         Top             =   1020
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "No 2"
         Height          =   255
         Left            =   2280
         TabIndex        =   39
         Top             =   9000
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "No 1"
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   8280
         Width           =   375
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   7
         Left            =   11400
         Picture         =   "OLD_Condensate2.frx":BF64
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "L.P. Extraction"
         Height          =   495
         Left            =   11760
         TabIndex        =   33
         Top             =   1680
         Width           =   735
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   6
         Left            =   7380
         Picture         =   "OLD_Condensate2.frx":C26E
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   5
         Left            =   6900
         Picture         =   "OLD_Condensate2.frx":C578
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   6180
         Picture         =   "OLD_Condensate2.frx":C882
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   8940
         Picture         =   "OLD_Condensate2.frx":CB8C
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   6540
         Picture         =   "OLD_Condensate2.frx":CE96
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   9300
         Picture         =   "OLD_Condensate2.frx":D1A0
         Stretch         =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "L.P Ahead Turbine"
         Height          =   375
         Left            =   11220
         TabIndex        =   31
         Top             =   300
         Width           =   975
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   10860
         X2              =   10860
         Y1              =   1380
         Y2              =   3360
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   11520
         X2              =   11520
         Y1              =   1980
         Y2              =   1200
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   7980
         X2              =   7980
         Y1              =   6300
         Y2              =   5940
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   7380
         X2              =   7980
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Label Label4 
         Caption         =   "L.P. Extraction"
         Height          =   495
         Left            =   6000
         TabIndex        =   30
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "150 psi."
         Height          =   255
         Left            =   7800
         TabIndex        =   29
         Top             =   3120
         Width           =   615
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   19
         X1              =   8280
         X2              =   7620
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line20 
         BorderStyle     =   3  'Dot
         X1              =   8100
         X2              =   8880
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   8880
         X2              =   10320
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   8880
         X2              =   8880
         Y1              =   5760
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   61
         X1              =   3240
         X2              =   3240
         Y1              =   6420
         Y2              =   5040
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   1800
         Picture         =   "OLD_Condensate2.frx":D4AA
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   2160
         Picture         =   "OLD_Condensate2.frx":D7B4
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Shape Shape38 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8230
         Top             =   6660
         Width           =   285
      End
      Begin VB.Shape Shape18 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7690
         Top             =   6660
         Width           =   290
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   8220
         Shape           =   3  'Circle
         Top             =   6420
         Width           =   315
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   7680
         Shape           =   3  'Circle
         Top             =   6420
         Width           =   315
      End
      Begin VB.Shape Shape37 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   6915
         Top             =   6600
         Width           =   225
      End
      Begin VB.Shape Shape36 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   6900
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   60
         X1              =   7140
         X2              =   7140
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   59
         X1              =   6900
         X2              =   6900
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Shape Shape15 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   7995
         Top             =   6600
         Width           =   225
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   7980
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   54
         X1              =   8220
         X2              =   8220
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   53
         X1              =   7980
         X2              =   7980
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Shape Shape34 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   7455
         Top             =   6600
         Width           =   225
      End
      Begin VB.Shape Shape33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   7440
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   56
         X1              =   7680
         X2              =   7680
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   55
         X1              =   7440
         X2              =   7440
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   6135
         Top             =   6600
         Width           =   225
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   6120
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   52
         X1              =   6360
         X2              =   6360
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   41
         X1              =   6120
         X2              =   6120
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Shape Shape30 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   5160
         Top             =   6600
         Width           =   225
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   5160
         Shape           =   3  'Circle
         Top             =   7080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   51
         X1              =   5400
         X2              =   5400
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   50
         X1              =   5160
         X2              =   5160
         Y1              =   6600
         Y2              =   7260
      End
      Begin VB.Shape Shape28 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8655
         Top             =   7620
         Width           =   210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   49
         X1              =   8880
         X2              =   8880
         Y1              =   7440
         Y2              =   7920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   48
         X1              =   8640
         X2              =   8640
         Y1              =   7560
         Y2              =   7920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   47
         X1              =   8520
         X2              =   8520
         Y1              =   6900
         Y2              =   6660
      End
      Begin VB.Shape Shape26 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   11340
         Shape           =   3  'Circle
         Top             =   5160
         Width           =   135
      End
      Begin VB.Shape Shape25 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   10140
         Shape           =   3  'Circle
         Top             =   4500
         Width           =   135
      End
      Begin VB.Shape Shape20 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   9960
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   9
         X1              =   9360
         X2              =   9360
         Y1              =   4620
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   6
         X1              =   9360
         X2              =   10200
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   9360
         X2              =   9360
         Y1              =   6840
         Y2              =   4920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   9360
         X2              =   9360
         Y1              =   7440
         Y2              =   7080
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   8160
         X2              =   8160
         Y1              =   5760
         Y2              =   5220
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   6180
         Picture         =   "OLD_Condensate2.frx":DABE
         Stretch         =   -1  'True
         Top             =   5340
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   6900
         Picture         =   "OLD_Condensate2.frx":DDC8
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   14
         X1              =   7440
         X2              =   7440
         Y1              =   5880
         Y2              =   5220
      End
      Begin VB.Shape Shape24 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   8040
         Top             =   5700
         Width           =   255
      End
      Begin VB.Shape Shape23 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   8400
         Top             =   5700
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   7620
         Top             =   5880
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Atmospheric Drain Tank"
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   9600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "875#/35#"
         Height          =   195
         Left            =   5040
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   4800
         Picture         =   "OLD_Condensate2.frx":E0D2
         Stretch         =   -1  'True
         Top             =   2220
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "IP Bleed "
         Height          =   435
         Left            =   60
         TabIndex        =   20
         Top             =   420
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   7020
         X2              =   7020
         Y1              =   5580
         Y2              =   6300
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   7620
         X2              =   7620
         Y1              =   4080
         Y2              =   4800
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   5340
         X2              =   5340
         Y1              =   5220
         Y2              =   6300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   45
         X1              =   10020
         X2              =   10020
         Y1              =   3840
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   44
         X1              =   5040
         X2              =   7920
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   43
         X1              =   3540
         X2              =   4800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   42
         X1              =   7500
         X2              =   7500
         Y1              =   7560
         Y2              =   8760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   39
         X1              =   3300
         X2              =   3900
         Y1              =   9180
         Y2              =   9180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   38
         X1              =   3300
         X2              =   3900
         Y1              =   8640
         Y2              =   8640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   26
         X1              =   3300
         X2              =   3300
         Y1              =   7980
         Y2              =   9180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   25
         X1              =   4560
         X2              =   4860
         Y1              =   9420
         Y2              =   9420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   21
         X1              =   4560
         X2              =   4560
         Y1              =   8760
         Y2              =   9420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   20
         X1              =   3960
         X2              =   4560
         Y1              =   9300
         Y2              =   9300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   19
         X1              =   3960
         X2              =   4560
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   10140
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   16
         X1              =   4920
         X2              =   4920
         Y1              =   4980
         Y2              =   6600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   15
         X1              =   6280
         X2              =   6280
         Y1              =   7560
         Y2              =   8760
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   1920
         X2              =   1920
         Y1              =   1440
         Y2              =   2100
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   2280
         X2              =   2280
         Y1              =   1440
         Y2              =   2160
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   13
         X1              =   2820
         X2              =   2820
         Y1              =   1800
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   13
         X1              =   9780
         X2              =   11460
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   9780
         X2              =   9780
         Y1              =   5280
         Y2              =   5820
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   9780
         X2              =   10020
         Y1              =   7260
         Y2              =   7260
      End
      Begin VB.Line Line12 
         BorderStyle     =   3  'Dot
         X1              =   11230
         X2              =   11230
         Y1              =   7500
         Y2              =   5400
      End
      Begin VB.Line Line11 
         BorderStyle     =   3  'Dot
         X1              =   10860
         X2              =   10980
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   10020
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   22
         X1              =   7140
         X2              =   7440
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   6300
         X2              =   6300
         Y1              =   6300
         Y2              =   6060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   5
         X1              =   4920
         X2              =   5160
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   11
         X1              =   6360
         X2              =   6900
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   5400
         X2              =   6120
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Image Image9 
         Height          =   360
         Left            =   3240
         Picture         =   "OLD_Condensate2.frx":E3DC
         Stretch         =   -1  'True
         Top             =   1035
         Width           =   285
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   3120
         X2              =   4920
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Evap"
         Height          =   255
         Left            =   6060
         TabIndex        =   15
         Top             =   5700
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   7
         X1              =   9780
         X2              =   10560
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1290
         Left            =   6780
         Top             =   6300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   12
         X1              =   3240
         X2              =   3240
         Y1              =   4800
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   23
         X1              =   8520
         X2              =   9780
         Y1              =   6915
         Y2              =   6915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   8880
         X2              =   9360
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   32
         X1              =   9780
         X2              =   10620
         Y1              =   7260
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   33
         X1              =   9780
         X2              =   9780
         Y1              =   5820
         Y2              =   7260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   34
         X1              =   10680
         X2              =   11760
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   35
         X1              =   10680
         X2              =   11760
         Y1              =   7500
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   36
         X1              =   11760
         X2              =   11760
         Y1              =   7500
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   37
         X1              =   5760
         X2              =   5760
         Y1              =   4320
         Y2              =   6600
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "LP Heater"
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
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   6960
         Width           =   795
      End
      Begin VB.Label Label19 
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
         Height          =   195
         Left            =   9480
         TabIndex        =   12
         Top             =   8160
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         Caption         =   "No 1"
         Height          =   255
         Left            =   10380
         TabIndex        =   11
         Top             =   5460
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "No 2"
         Height          =   255
         Left            =   10380
         TabIndex        =   10
         Top             =   6960
         Width           =   495
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   9600
         Shape           =   3  'Circle
         Top             =   1635
         Width           =   15
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   0
         Left            =   9540
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   15
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   11940
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   255
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1515
         Left            =   9840
         Top             =   3360
         Width           =   2595
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1755
         Index           =   2
         Left            =   9060
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1980
         Left            =   11640
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   2820
         X2              =   300
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   7320
         Top             =   5880
         Width           =   255
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   8640
         Shape           =   3  'Circle
         Top             =   7740
         Width           =   255
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   7260
         Top             =   6300
         Width           =   1515
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   5040
         Top             =   6300
         Width           =   555
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   6000
         Top             =   6300
         Width           =   555
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   6000
         Top             =   5640
         Width           =   615
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   885
         Left            =   10700
         OleObjectBlob   =   "OLD_Condensate2.frx":E6E6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hotwell.vsd"
         TabIndex        =   50
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   3345
         Left            =   10140
         OleObjectBlob   =   "OLD_Condensate2.frx":15EFE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   51
         Top             =   -1380
         Width           =   3015
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         Height          =   195
         Left            =   10860
         TabIndex        =   14
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   6300
         X2              =   6300
         Y1              =   5760
         Y2              =   5340
      End
      Begin VB.Shape Shape32 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3060
         Top             =   2220
         Width           =   315
      End
      Begin VB.Shape Shape35 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3060
         Top             =   2880
         Width           =   315
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   2700
         Top             =   1815
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   5
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   2880
         Shape           =   3  'Circle
         Top             =   1395
         Width           =   735
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   37
         X1              =   300
         X2              =   300
         Y1              =   2820
         Y2              =   3360
      End
      Begin VB.Label Label27 
         Caption         =   "from Aux Exhaust"
         Height          =   255
         Left            =   8640
         TabIndex        =   83
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label26 
         Caption         =   "DFT Dump Regulator"
         Height          =   615
         Left            =   1800
         TabIndex        =   82
         Top             =   4680
         Width           =   795
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1320
         OleObjectBlob   =   "OLD_Condensate2.frx":1D916
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   103
         Top             =   4380
         Width           =   405
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1320
         OleObjectBlob   =   "OLD_Condensate2.frx":2592E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   106
         Top             =   4380
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1320
         OleObjectBlob   =   "OLD_Condensate2.frx":2D746
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   105
         Top             =   4380
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DFT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1320
         OleObjectBlob   =   "OLD_Condensate2.frx":3575E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   104
         Top             =   4380
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   1440
         X2              =   1440
         Y1              =   3720
         Y2              =   5280
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   900
         OleObjectBlob   =   "OLD_Condensate2.frx":3DF76
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   111
         Top             =   6840
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   900
         OleObjectBlob   =   "OLD_Condensate2.frx":45F8E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   114
         Top             =   6840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   900
         OleObjectBlob   =   "OLD_Condensate2.frx":4DDA6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   113
         Top             =   6840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE DWT_dump 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   900
         OleObjectBlob   =   "OLD_Condensate2.frx":565BE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   112
         Top             =   6840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   1020
         X2              =   1020
         Y1              =   5880
         Y2              =   7740
      End
      Begin VB.Label DWT_dump_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   141
         Top             =   6480
         Width           =   555
      End
      Begin VB.Label DFT_dump_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   142
         Top             =   4320
         Width           =   555
      End
      Begin VB.Label DWT_dump_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1260
         TabIndex        =   140
         Top             =   6780
         Width           =   555
      End
      Begin VB.Label DFT_dump_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   780
         TabIndex        =   143
         Top             =   4080
         Width           =   555
      End
      Begin VB.Label COND_recirc_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6420
         TabIndex        =   144
         Top             =   5040
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   18
         X1              =   4920
         X2              =   4920
         Y1              =   1620
         Y2              =   4740
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   560
         Index           =   0
         Left            =   1380
         OleObjectBlob   =   "OLD_Condensate2.frx":5E5D6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   115
         Top             =   0
         Width           =   405
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   1380
         OleObjectBlob   =   "OLD_Condensate2.frx":665EE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   135
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   1380
         OleObjectBlob   =   "OLD_Condensate2.frx":6E406
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   134
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   1380
         OleObjectBlob   =   "OLD_Condensate2.frx":7661E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   180
         OleObjectBlob   =   "OLD_Condensate2.frx":7F036
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   107
         Top             =   2220
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   180
         OleObjectBlob   =   "OLD_Condensate2.frx":8704E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   110
         Top             =   2220
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   180
         OleObjectBlob   =   "OLD_Condensate2.frx":8EE66
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   109
         Top             =   2220
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   180
         OleObjectBlob   =   "OLD_Condensate2.frx":9767E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   108
         Top             =   2220
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   560
         Index           =   0
         Left            =   3900
         OleObjectBlob   =   "OLD_Condensate2.frx":9F696
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   116
         Top             =   0
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   560
         Index           =   3
         Left            =   3900
         OleObjectBlob   =   "OLD_Condensate2.frx":A76AE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   138
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   560
         Index           =   2
         Left            =   3900
         OleObjectBlob   =   "OLD_Condensate2.frx":AF4C6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   137
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   560
         Index           =   1
         Left            =   3900
         OleObjectBlob   =   "OLD_Condensate2.frx":B76DE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   26
         X1              =   5280
         X2              =   300
         Y1              =   360
         Y2              =   360
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   6420
         OleObjectBlob   =   "OLD_Condensate2.frx":C00F6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   129
         Top             =   4500
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   6420
         OleObjectBlob   =   "OLD_Condensate2.frx":C810E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   119
         Top             =   4500
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   6420
         OleObjectBlob   =   "OLD_Condensate2.frx":D0126
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   118
         Top             =   4500
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   6420
         OleObjectBlob   =   "OLD_Condensate2.frx":D833E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   117
         Top             =   4500
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Auxiliary Exhaust from Turbo feed and Cargo Pumps"
         Height          =   975
         Left            =   1560
         TabIndex        =   22
         Top             =   2160
         Width           =   915
      End
   End
   Begin VB.CommandButton Pump_Stby 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stby"
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
      Index           =   0
      Left            =   10440
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7860
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Pump_Stop 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Stop"
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
      Left            =   9960
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7860
      Width           =   495
   End
   Begin VB.CommandButton Pump_Start 
      BackColor       =   &H0000FF00&
      Caption         =   "Start"
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
      Index           =   0
      Left            =   9480
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7860
      Width           =   495
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   123
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "OLD_Condensate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()


' 6/4/00  set_dummy_focus = True

Feed_Water.Show


End Sub

Private Sub Command2_Click()

' 14/2/00
'Shaft_rpm.Caption = Format(LSS_reals(80), "##0") & " rpm"
If LSS_reals(80) > 25# Then
  Gland_Steam_Full_away.Show
Else
  Gland_Steam.Show
End If

End Sub

Private Sub Command3_Click()


' 4/9/00
Main_Condenser.ZOrder

End Sub

Private Sub Command4_Click()


' 4/9/00  Point 8.147
Distil.ZOrder


End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idCondensate)



' 3/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' 11/12/00
    valve(13).Visible = False
    Line1(2).Visible = False
    Line1(13).Visible = False
    
    ONpump(2).Visible = False
    OFFpump(2).Visible = False
    sbyPUMP(2).Visible = False
    
    ONpump(3).Visible = False
    OFFpump(3).Visible = False
    sbyPUMP(3).Visible = False
    
    ' 21/9/00
    COND_recirc_bypass.Visible = False
    COND_recirc_reg.Visible = False
    DWT_dump_bypass.Visible = False
    DWT_dump_reg.Visible = False
    DFT_dump_bypass.Visible = False
    DFT_dump_reg.Visible = False
    
Else


    ' 11/12/00
    valve(13).Visible = True
    Line1(2).Visible = True
    Line1(13).Visible = True

    ONpump(2).Visible = True
    OFFpump(2).Visible = True
    sbyPUMP(2).Visible = True
    
    ONpump(3).Visible = True
    OFFpump(3).Visible = True
    sbyPUMP(3).Visible = True
    
    ' 21/9/00
    COND_recirc_bypass.Visible = True
    COND_recirc_reg.Visible = True
    DWT_dump_bypass.Visible = True
    DWT_dump_reg.Visible = True
    DFT_dump_bypass.Visible = True
    DFT_dump_reg.Visible = True
    
    
End If


' 1/12/99
VScroll1.Visible = PTT
VScroll4.Visible = PTT

HScroll3.Visible = PTT

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then

    HavenStretch1.Visible = False


    Call JYAScreenResize(Me, idCondensate)
    
    HavenStretch1.Visible = True
    


End If

End Sub

Private Sub HScroll3_Change()
Dim AIvalue As Integer

' 29/10/99  COND_recirc_bypass

' 1/12/99
'Call set_VSCROLL(17, HScroll3.Value)
AIvalue = HScroll3.Value / 4
Call setAI(11, AIvalue)


End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

'   Note:  PumpIO(4=OFF,nPump)
' Steam Sim
Select Case Index
  Case 0
    Call setPCDI(PumpIO_LSS(4, 35)) ' ADT pump #1
  Case 1
    Call setPCDI(PumpIO_LSS(4, 36)) ' ADT pump #2
  Case 2
    Call setPCDI(PumpIO_LSS(4, 23)) ' Condensate pump #1
  Case 3
    Call setPCDI(PumpIO_LSS(4, 24)) ' Condensate pump #2
End Select

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)
' Steam Sim
Select Case Index
  Case 0
    Call setPCDI(PumpIO_LSS(2, 35)) ' ADT pump #1
  Case 1
    Call setPCDI(PumpIO_LSS(2, 36)) ' ADT pump #2
  Case 2
    Call setPCDI(PumpIO_LSS(2, 23)) ' Condensate pump #1
  Case 3
    Call setPCDI(PumpIO_LSS(2, 24)) ' Condensate pump #2
End Select

End Sub

Private Sub P_cond_out_Click()

' 12/12/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 99
Alarm_set_point.Show

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(5=Stand-bye,nPump)
' Steam Sim
Select Case Index
  Case 0
    Call setPCDI(PumpIO_LSS(5, 35)) ' ADT pump #1
  Case 1
    Call setPCDI(PumpIO_LSS(5, 36)) ' ADT pump #2
  Case 2
    Call setPCDI(PumpIO_LSS(5, 23)) ' Condensate pump #1
  Case 3
    Call setPCDI(PumpIO_LSS(5, 24)) ' Condensate pump #2
End Select

End Sub

Private Sub T_air_eject_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 93
Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim VV_pos As Integer
Dim green_colour As Integer

Dim Hotwell_level As Integer
Dim DFT_level As Integer
Dim DWT_level As Integer
Dim ADT_level As Integer
Dim CDT_level As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 7/1/00
Call paint_alarm(Me.Z_DFT, 76, 81, 0, 0)
Call paint_alarm(Me.Z_ADT, 90, 96, 0, 0)

' 5/5/00  ( See S Green point 5.111 )
'Call paint_alarm(Me.Z_DWT, 80, 96, 0, 0)
Call paint_alarm(Me.Z_DWT, 80, 0, 0, 0)

Call paint_alarm(Me.Z_Hotwell, 92, 0, 0, 0)
Call paint_alarm(Me.T_air_eject, 93, 0, 0, 0)

' 12/12/00
Call paint_alarm(Me.P_cond_out, 99, 0, 0, 0)


'  === 9/11/99 STEAM Regulator  Valves ===
Call update_regulator(366, Reg_875_35(0), Reg_875_35(1), Reg_875_35(2), Reg_875_35(3))
Call update_regulator(368, Dump_35(0), Dump_35(1), Dump_35(2), Dump_35(3))
Call update_regulator(372, Reg_IP(0), Reg_IP(1), Reg_IP(2), Reg_IP(3))

'  === 6/11/99 STEAM Bypass  Valves ===
Call update_bypass(367, valve(109))
Call update_bypass(369, valve(91))
Call update_bypass(373, valve(113))



'          PC_REALS(300) = F_DFT_main_IN
'          PC_REALS(301) = F_DFT_dump
'          PC_REALS(302) = F_DWT_dump
'          PC_REALS(303) = F_DFT_hotwell

' 18 / 10 / 99
'        PC_REALS(451) = Z_DFT_inches
'        PC_REALS(452) = Z_DWT * 100#
'        PC_REALS(453) = Z_ADT * 100#
'
'        PC_REALS(454) = DFT_dump_reg * 100#
'        PC_REALS(455) = DFT_dump_bypass * 100#
'
'        PC_REALS(456) = DWT_dump_reg * 100#
'        PC_REALS(457) = DWT_dump_bypass * 100#

'        PC_REALS(458) = P_cond_out
'        PC_REALS(459) = P_ADT_out

' 29/10/99
'        PC_REALS(460) = COND_recirc_reg * 100#
'        PC_REALS(461) = COND_recirc_bypass * 100#

' 1/11/99
'        PC_REALS(462) = PX_COND
'        PC_REALS(463) = PY_COND
'        PC_REALS(464) = F_ADT_OUT
'
'        PC_REALS(465) = Z_hotwell_inches
'
'        PC_REALS(466) = F_hotwell_out
'        PC_REALS(467) = F_cond_recirc(1)
'        PC_REALS(468) = F_cond_recirc(2)
'        PC_REALS(469) = F_cond_recirc(3)
'        PC_REALS(470) = F_cond_htrs


'    10/11/99
'          PC_REALS(304) = T_hotwell
'          PC_REALS(305) = T_air_eject
'          PC_REALS(306) = T_FWG
'          PC_REALS(307) = T_cond_recirc
'          PC_REALS(308) = T_LP_htr
'          PC_REALS(309) = T_DFT
'          PC_REALS(310) = T_HP_htr
'          PC_REALS(311) = T_DWT
'          PC_REALS(312) = T_ADT

T_hotwell.text = Format(LSS_reals(304) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_air_eject.text = Format(LSS_reals(305) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_FWG.text = Format(LSS_reals(306) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_cond_recirc.text = Format(LSS_reals(307) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_LP_htr.text = Format(LSS_reals(308) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_DFT.text = Format(LSS_reals(309) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

T_DWT.text = Format(LSS_reals(311) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_ADT.text = Format(LSS_reals(312) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char



    VScroll2.Visible = Advanced_mode
    VScroll3.Visible = Advanced_mode
  

    F_DFT_main_IN.Visible = Advanced_mode
    F_DFT_dump.Visible = Advanced_mode
    F_DWT_dump.Visible = Advanced_mode
    F_DFT_hotwell.Visible = Advanced_mode
    
    F_ADT_out.Visible = Advanced_mode
    
    F_hotwell_out.Visible = Advanced_mode
    F_cond_recirc(1).Visible = Advanced_mode
    F_cond_recirc(2).Visible = Advanced_mode
    F_cond_recirc(3).Visible = Advanced_mode
    F_cond_htrs.Visible = Advanced_mode
    
    
  '   28/8/00
  ' PX_cond.Visible = Advanced_mode
    P_cond_out.Visible = Advanced_mode
   
    PY_cond.Visible = Advanced_mode
    
' === 22/01/02 ===
    PY_line(0).Visible = Advanced_mode
    PY_line(1).Visible = Advanced_mode



If SI_units Then
    F_DFT_main_IN.text = Format(LSS_reals(300) * Flow_mult, "###0.0") & Flow_chars
    F_DFT_dump.text = Format(LSS_reals(301) * Flow_mult, "###0.0") & Flow_chars
    F_DWT_dump.text = Format(LSS_reals(302) * Flow_mult, "###0.0") & Flow_chars
    F_DFT_hotwell.text = Format(LSS_reals(303) * Flow_mult, "###0.0") & Flow_chars
    
    F_ADT_out.text = Format(LSS_reals(464) * Flow_mult, "###0.0") & Flow_chars
    
    F_hotwell_out.text = Format(LSS_reals(466) * Flow_mult, "###0.0") & Flow_chars
    F_cond_recirc(1).text = Format(LSS_reals(467) * Flow_mult, "###0.0") & Flow_chars
    F_cond_recirc(2).text = Format(LSS_reals(468) * Flow_mult, "###0.0") & Flow_chars
    F_cond_recirc(3).text = Format(LSS_reals(469) * Flow_mult, "###0.0") & Flow_chars
    F_cond_htrs.text = Format(LSS_reals(470) * Flow_mult, "###0.0") & Flow_chars

    Z_DFT.text = Format(LSS_reals(451) * 25.4, "###0.0") & " mm"
    Z_Hotwell.text = Format(LSS_reals(465) * 25.4, "###0.0") & " mm"
    
    P_cond_out.text = Format(LSS_reals(458) * Press_mult, "##0.0") & Press_chars     ' Bar
    P_ADT_out.text = Format(LSS_reals(459) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    PX_cond.text = Format(LSS_reals(462) * Press_mult, "##0.0") & Press_chars     ' Bar
    PY_cond.text = Format(LSS_reals(463) * Press_mult, "##0.0") & Press_chars     ' Bar
Else
    F_DFT_main_IN.text = Format(LSS_reals(300) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_DFT_dump.text = Format(LSS_reals(301) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_DWT_dump.text = Format(LSS_reals(302) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_DFT_hotwell.text = Format(LSS_reals(303) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_ADT_out.text = Format(LSS_reals(464) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_hotwell_out.text = Format(LSS_reals(466) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_cond_recirc(1).text = Format(LSS_reals(467) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_cond_recirc(2).text = Format(LSS_reals(468) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_cond_recirc(3).text = Format(LSS_reals(469) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_cond_htrs.text = Format(LSS_reals(470) * Flow_mult / 3600#, "###0.0") & " lb/s"

    Z_DFT.text = Format(LSS_reals(451), "###0.0") & " in"
    Z_Hotwell.text = Format(LSS_reals(465), "###0.0") & " in"
    
    P_cond_out.text = Format(LSS_reals(458) * Press_mult, "###0") & Press_chars    ' psi
    P_ADT_out.text = Format(LSS_reals(459) * Press_mult, "###0") & Press_chars    ' psi
    
    PX_cond.text = Format(LSS_reals(462) * Press_mult, "###0") & Press_chars    ' psi
    PY_cond.text = Format(LSS_reals(463) * Press_mult, "###0") & Press_chars    ' psi
End If


'  Hotwell Tank  ( 0 inches to +8.0 inches )
Hotwell_level = LSS_reals(465) * 100# / 8#
If Hotwell_level > 99 Then
  Hotwell_tank.LineWidth = 100
ElseIf Hotwell_level < 1 Then
  Hotwell_tank.LineWidth = 0
Else
  Hotwell_tank.LineWidth = Hotwell_level
End If

'  DFT Tank  ( -35 inches to +35 inches )
DFT_level = (LSS_reals(451) + 35#) * 100# / 70#
If DFT_level > 99 Then
  valve(302).LineWidth = 100
ElseIf DFT_level < 1 Then
  valve(302).LineWidth = 0
Else
  valve(302).LineWidth = DFT_level
End If

'  DWT Tank  ( 0 to 100 % )
Z_DWT.text = Format(LSS_reals(452), "###0") & " %"
DWT_level = LSS_reals(452)
If DWT_level > 99 Then
  DWT_tank.LineWidth = 100
ElseIf DWT_level < 1 Then
  DWT_tank.LineWidth = 0
Else
  DWT_tank.LineWidth = DWT_level
End If

'  ADT Tank  ( 0 to 100 % )
Z_ADT.text = Format(LSS_reals(453), "###0") & " %"
ADT_level = LSS_reals(453)
If ADT_level > 99 Then
  ADT_tank.LineWidth = 100
ElseIf ADT_level < 1 Then     ' 11/11/99 WAS DWT_level corrected to ADT_level
  ADT_tank.LineWidth = 0
Else
  ADT_tank.LineWidth = ADT_level
End If

' 10/8/00
'  CDT Tank  ( 0 to 100 % )
Z_CDT.text = Format(LSS_reals(154), "###0") & " %"
CDT_level = LSS_reals(154)
If CDT_level > 99 Then
  CDT_tank.LineWidth = 100
ElseIf CDT_level < 1 Then
  CDT_tank.LineWidth = 0
Else
  CDT_tank.LineWidth = CDT_level
End If

If CDT_level < 83 Then
    Sight_Glass.FillColor = &H0&         ' Black
Else
    If VB_steam_logicals(120) Then
        ' see Fault(621) = Oil in the CDT
        Sight_Glass.FillColor = &HFFFF&      ' Yellow
    Else
        Sight_Glass.FillColor = &HFF0000     ' Blue
    End If
End If


' COND Re-circ Regulator
COND_recirc_reg.Caption = Format(LSS_reals(460), "###0") & " %"

Call update_regulator(460, Reg_cond_recirc(0), Reg_cond_recirc(1), Reg_cond_recirc(2), Reg_cond_recirc(3))


' COND Re-circ Bypass
COND_recirc_bypass.Caption = Format(LSS_reals(461), "###0") & " %"

Call update_bypass(461, valve(5))



' DFT Dump Regulator
DFT_dump_reg.Caption = Format(LSS_reals(454), "###0") & " %"
' 6/11/99
Call update_regulator(454, DFT_dump(0), DFT_dump(1), DFT_dump(2), DFT_dump(3))

' DFT Dump Bypass
DFT_dump_bypass.Caption = Format(LSS_reals(455), "###0") & " %"

Call update_bypass(455, valve(21))


' DWT Dump Regulator
DWT_dump_reg.Caption = Format(LSS_reals(456), "###0") & " %"
' 6/11/99
Call update_regulator(456, DWT_dump(0), DWT_dump(1), DWT_dump(2), DWT_dump(3))



' DWT Dump Bypass
DWT_dump_bypass.Caption = Format(LSS_reals(457), "###0") & " %"

Call update_bypass(457, valve(17))



' Note: update_pump_LSS(syxxx,index,nPump)
Call update_pump_LSS(Me, 0, 35)      ' ADT Pump # 1
Call update_pump_LSS(Me, 1, 36)      ' ADT Pump # 2

Call update_pump_LSS(Me, 2, 23)      ' Condensate Pump # 1
Call update_pump_LSS(Me, 3, 24)      ' Condensate Pump # 2


' update_vvalve_LSS(syxxx,nValve,index1,index2)

' 26/10/99 ADDED
   Call update_vvalve_LSS(Me, 1, 1, -1)
   Call update_vvalve_LSS(Me, 4, 4, -1)
   
'  Valve 5 should be ANALOG ? but where is the ER signal?? AI(??)
'  See Condensate  Xref  in EXCEL
'   Call update_vvalve_LSS(Me, 5, 5, -1)


   Call update_vvalve_LSS(Me, 2, 2, -1)
   Call update_vvalve_LSS(Me, 3, 3, -1)
   Call update_vvalve_LSS(Me, 8, 8, -1)
   Call update_vvalve_LSS(Me, 9, 9, -1)
   
 For i = 10 To 16
   Call update_vvalve_LSS(Me, i, i, -1)
 Next i
   
   Call update_vvalve_LSS(Me, 18, 18, -1)
   Call update_vvalve_LSS(Me, 20, 20, -1)
   Call update_vvalve_LSS(Me, 22, 22, -1)
   Call update_vvalve_LSS(Me, 23, 23, -1)

   Call update_vvalve_LSS(Me, 77, 77, -1)
   Call update_vvalve_LSS(Me, 80, 80, -1)
   Call update_vvalve_LSS(Me, 81, 81, -1)
   
   Call update_vvalve_LSS(Me, 107, 107, -1)
   Call update_vvalve_LSS(Me, 108, 108, -1)
   
   Call update_vvalve_LSS(Me, 111, 111, -1)
   Call update_vvalve_LSS(Me, 112, 112, -1)
   
 For i = 270 To 275
   Call update_vvalve_LSS(Me, i, i, -1)
 Next i
   
' 10/8/00   CDT Outlet Valve
   Call update_vvalve_LSS(Me, 583, 583, -1)
   
   

End Sub

Private Sub valve_Click(Index As Integer)

'   14/10/99
'Sub valve_pcdi_LSS(nValve As Integer)

'   Call valve_pcdi_LSS(Index + 70)
    Call valve_pcdi_LSS(Index)

End Sub

Private Sub VScroll1_Change()
Dim AIvalue As Integer

' 18/10/99  DWT_dump_bypass

' 1/12/99
'Call set_VSCROLL(15, VScroll1.Value)

AIvalue = VScroll1.Value / 4
Call setAI(9, AIvalue)

End Sub

Private Sub VScroll2_Change()

' 18/10/99  Z_DFT   ( TESTING )
' 7/1/00
'Call set_VSCROLL(16, VScroll2.Value)
Call set_VSCROLL(16, (VScroll2.max - VScroll2.Value))


End Sub

Private Sub VScroll3_Change()

' 1/11/99  Z_hotwell   ( TESTING )
' 7/1/00
'Call set_VSCROLL(18, VScroll3.Value)
Call set_VSCROLL(18, (VScroll3.max - VScroll3.Value))

End Sub

Private Sub VScroll4_Change()
Dim AIvalue As Integer

' 18/10/99   DFT_dump_bypass

' 1/12/99
'Call set_VSCROLL(14, VScroll4.Value)

AIvalue = VScroll4.Value / 4
Call setAI(10, AIvalue)


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
Sub update_bypass(nReals As Integer, obj As Object)
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

Private Sub VScroll6_Change()


' === 12/12/01 TESTING R_de_aerator ===
Call set_VSCROLL(37, VScroll6.Value)
Label38.Caption = VScroll6.Value

End Sub

Private Sub Z_ADT_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 96
nAlarm_LO = 90
Alarm_set_point.Show

End Sub

Private Sub Z_DFT_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 81
nAlarm_LO = 76
Alarm_set_point.Show

End Sub

Private Sub Z_DWT_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 80
Alarm_set_point.Show

End Sub

Private Sub Z_Hotwell_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 92
Alarm_set_point.Show

End Sub
