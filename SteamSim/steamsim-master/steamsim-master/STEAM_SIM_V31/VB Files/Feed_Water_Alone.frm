VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Feed_Water_Alone 
   Caption         =   "Feed Water"
   ClientHeight    =   12360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12360
   ScaleWidth      =   16200
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   11835
      Left            =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   20876
      Begin VB.HScrollBar HScroll_Aux_feed_PORT 
         Height          =   135
         Left            =   8430
         Max             =   1000
         TabIndex        =   187
         Top             =   2460
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll_Aux_feed_STBD 
         Height          =   135
         Left            =   5640
         Max             =   1000
         TabIndex        =   186
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   4350
         TabIndex        =   185
         Top             =   10710
         Width           =   285
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2520
         TabIndex        =   184
         Top             =   7260
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   10650
         TabIndex        =   182
         Top             =   3930
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   10320
         TabIndex        =   181
         Top             =   2490
         Width           =   285
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   3030
         TabIndex        =   180
         Top             =   5550
         Width           =   285
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   2880
         TabIndex        =   179
         Top             =   2820
         Width           =   285
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   11250
         TabIndex        =   178
         Top             =   2820
         Width           =   285
      End
      Begin VB.TextBox Psteam101 
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
         Left            =   12120
         MousePointer    =   1  'Arrow
         TabIndex        =   174
         Text            =   "xx Bar"
         Top             =   2700
         Width           =   735
      End
      Begin VB.TextBox Drum_water_level_port 
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
         Left            =   12120
         MousePointer    =   1  'Arrow
         TabIndex        =   173
         Text            =   "xx in"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Drum_water_level_stbd 
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
         Left            =   1620
         MousePointer    =   1  'Arrow
         TabIndex        =   172
         Text            =   "-xx in"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Psteam1 
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
         Left            =   1620
         MousePointer    =   1  'Arrow
         TabIndex        =   171
         Text            =   "xx Bar"
         Top             =   2700
         Width           =   735
      End
      Begin VB.TextBox T_econ_IN 
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
         Left            =   4080
         TabIndex        =   148
         Text            =   "xxx  F"
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox F_main_feed 
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
         Left            =   13650
         TabIndex        =   147
         Text            =   "xx.x kg/s"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox T_main_feed 
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
         Left            =   4260
         TabIndex        =   146
         Text            =   "xxx  F"
         Top             =   1770
         Width           =   675
      End
      Begin VB.TextBox T_econ_IN 
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
         Left            =   9960
         TabIndex        =   127
         Text            =   "xxx  F"
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox T_main_feed 
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
         Left            =   9780
         TabIndex        =   126
         Text            =   "xxx  F"
         Top             =   1800
         Width           =   675
      End
      Begin VB.TextBox F_main_feed 
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
         Left            =   13680
         TabIndex        =   125
         Text            =   "xx.x kg/s"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox TLO_main_feed 
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
         Left            =   13380
         TabIndex        =   120
         Text            =   "xx F"
         Top             =   9240
         Width           =   795
      End
      Begin VB.TextBox TLO_main_feed 
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
         Left            =   12540
         TabIndex        =   119
         Text            =   "xx F"
         Top             =   9240
         Width           =   795
      End
      Begin VB.TextBox P_main_out 
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
         Left            =   12540
         TabIndex        =   118
         Text            =   "x.x Bar"
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox PLO_total_feed_pump 
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
         Left            =   13380
         TabIndex        =   117
         Text            =   "x.x Bar"
         Top             =   9540
         Width           =   795
      End
      Begin VB.TextBox PLO_total_feed_pump 
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
         Left            =   12540
         TabIndex        =   116
         Text            =   "x.x Bar"
         Top             =   9540
         Width           =   795
      End
      Begin VB.CommandButton REMOTE_feed_control 
         BackColor       =   &H00C0C0C0&
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   13560
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   10680
         Width           =   255
      End
      Begin VB.CommandButton REMOTE_feed_control 
         BackColor       =   &H0000FFFF&
         Caption         =   "R"
         Height          =   255
         Index           =   1
         Left            =   12960
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   10680
         Width           =   255
      End
      Begin VB.TextBox P_main_out 
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
         Left            =   11220
         TabIndex        =   111
         Text            =   "x.x Bar"
         Top             =   120
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton AUTO_feed_pump 
         BackColor       =   &H0000FFFF&
         Caption         =   "A"
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
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   6420
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton AUTO_feed_pump 
         BackColor       =   &H0000FFFF&
         Caption         =   "A"
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
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   4620
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   13440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   10380
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   12840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   10380
         UseMaskColor    =   -1  'True
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
         Index           =   4
         Left            =   13440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   10140
         UseMaskColor    =   -1  'True
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
         Index           =   3
         Left            =   12840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   10140
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
         Index           =   4
         Left            =   13440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   9900
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
         Index           =   3
         Left            =   12840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   9900
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   135
         Left            =   3240
         Max             =   1000
         TabIndex        =   93
         Top             =   9240
         Width           =   1215
      End
      Begin VB.TextBox T_HP_htr 
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
         Left            =   1500
         TabIndex        =   92
         Text            =   "xx F"
         Top             =   8520
         Width           =   795
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
         Left            =   13020
         TabIndex        =   91
         Text            =   "xx F"
         Top             =   5700
         Width           =   795
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
         Left            =   13020
         TabIndex        =   90
         Text            =   "20 %"
         Top             =   5400
         Width           =   795
      End
      Begin VB.TextBox F_aux_feed 
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
         Left            =   8190
         TabIndex        =   87
         Text            =   "xx.x kg/s"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_aux_feed 
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
         Left            =   5610
         TabIndex        =   86
         Text            =   "xx.x kg/s"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_main_feed 
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
         Left            =   8670
         TabIndex        =   85
         Text            =   "xx.x kg/s"
         Top             =   1170
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_main_feed 
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
         Left            =   4920
         TabIndex        =   84
         Text            =   "xx.x kg/s"
         Top             =   1170
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_aux 
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
         Left            =   5400
         TabIndex        =   83
         Text            =   "xx.x kg/s"
         Top             =   9900
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_aux 
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
         Left            =   3540
         TabIndex        =   82
         Text            =   "xx.x kg/s"
         Top             =   9840
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_recirc 
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
         Left            =   12300
         TabIndex        =   81
         Text            =   "xx.x kg/s"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_recirc 
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
         Left            =   12840
         TabIndex        =   80
         Text            =   "xx.x kg/s"
         Top             =   7080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF 
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
         Left            =   4020
         TabIndex        =   79
         Text            =   "xx.x kg/s"
         Top             =   7740
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF 
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
         Left            =   4020
         TabIndex        =   78
         Text            =   "xx.x kg/s"
         Top             =   8640
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_main 
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
         Left            =   10620
         TabIndex        =   77
         Text            =   "xx.x kg/s"
         Top             =   8040
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_main 
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
         Left            =   10620
         TabIndex        =   76
         Text            =   "xx.x kg/s"
         Top             =   8880
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox P_AF 
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
         Left            =   1500
         TabIndex        =   75
         Text            =   "x.x Bar"
         Top             =   10020
         Width           =   795
      End
      Begin VB.TextBox P_MF 
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
         Left            =   1500
         TabIndex        =   74
         Text            =   "x.x Bar"
         Top             =   7980
         Width           =   795
      End
      Begin VB.TextBox P_main 
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
         Left            =   5100
         TabIndex        =   73
         Text            =   "x.x Bar"
         Top             =   8640
         Width           =   795
      End
      Begin VB.TextBox P_main 
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
         Left            =   5100
         TabIndex        =   72
         Text            =   "x.x Bar"
         Top             =   7740
         Width           =   795
      End
      Begin VB.CommandButton CLOSE_VV 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   2
         Left            =   12900
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   8640
         Width           =   675
      End
      Begin VB.CommandButton CLOSE_VV 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   1
         Left            =   12900
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   7980
         Width           =   675
      End
      Begin VB.CommandButton OPEN_VV 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPEN"
         Height          =   255
         Index           =   2
         Left            =   12300
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   8640
         Width           =   615
      End
      Begin VB.CommandButton OPEN_VV 
         BackColor       =   &H0000FF00&
         Caption         =   "OPEN"
         Height          =   255
         Index           =   1
         Left            =   12300
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   7980
         Width           =   615
      End
      Begin VB.CommandButton AUTO_VV 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
         Height          =   255
         Index           =   2
         Left            =   13560
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   8640
         Width           =   615
      End
      Begin VB.CommandButton AUTO_VV 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
         Height          =   255
         Index           =   1
         Left            =   13560
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   7980
         Width           =   615
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   135
         Left            =   4680
         Max             =   1000
         TabIndex        =   47
         Top             =   5700
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   135
         Left            =   4680
         Max             =   1000
         TabIndex        =   46
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   14400
         Top             =   360
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   135
         Left            =   5220
         Max             =   1000
         TabIndex        =   44
         Top             =   9660
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   135
         Left            =   3420
         Max             =   1000
         TabIndex        =   43
         Top             =   9660
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Condensate System"
         Height          =   675
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3720
         Width           =   1095
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
         Left            =   10440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   10980
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
         Index           =   0
         Left            =   10020
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   10980
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
         Left            =   9540
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   10980
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox P_aux_feed_out 
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
         Left            =   8700
         TabIndex        =   12
         Text            =   "75 Bar"
         Top             =   10500
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   70
         Left            =   9600
         TabIndex        =   0
         Top             =   10140
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
         Index           =   71
         Left            =   10320
         TabIndex        =   1
         Top             =   7560
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
         Index           =   72
         Left            =   10500
         TabIndex        =   3
         Top             =   10380
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
         Left            =   3600
         TabIndex        =   4
         Top             =   8760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   32768
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   83
         Left            =   4320
         TabIndex        =   6
         Top             =   6540
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
         Index           =   84
         Left            =   4320
         TabIndex        =   7
         Top             =   4980
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
         Height          =   375
         Index           =   15
         Left            =   14340
         TabIndex        =   8
         Top             =   5040
         Visible         =   0   'False
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   9960
         TabIndex        =   17
         Top             =   10425
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   9360
         TabIndex        =   18
         Top             =   7610
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   80
         Left            =   12240
         TabIndex        =   20
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   2
         Left            =   7860
         TabIndex        =   22
         Top             =   9050
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   87
         Left            =   8760
         TabIndex        =   23
         Top             =   9000
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
         Index           =   73
         Left            =   3600
         TabIndex        =   2
         Top             =   7320
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
         Height          =   375
         Index           =   16
         Left            =   14280
         TabIndex        =   25
         Top             =   5460
         Visible         =   0   'False
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   95
         Left            =   9780
         TabIndex        =   29
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   96
         Left            =   8220
         TabIndex        =   30
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   97
         Left            =   5940
         TabIndex        =   31
         Top             =   4980
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
         Index           =   99
         Left            =   5940
         TabIndex        =   32
         Top             =   6540
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
         Index           =   100
         Left            =   4680
         TabIndex        =   33
         Top             =   6540
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
         Index           =   101
         Left            =   4680
         TabIndex        =   34
         Top             =   4980
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
         Left            =   4980
         TabIndex        =   35
         Top             =   4260
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
         Height          =   1335
         Index           =   44
         Left            =   12780
         TabIndex        =   37
         Top             =   4620
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   2355
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   48
         Left            =   13920
         TabIndex        =   39
         Top             =   180
         Visible         =   0   'False
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
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   12
         Left            =   6480
         TabIndex        =   41
         Top             =   9660
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
         Index           =   51
         Left            =   4740
         TabIndex        =   42
         Top             =   9660
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   3
         Left            =   9900
         TabIndex        =   94
         Top             =   5220
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   65535
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   4
         Left            =   8340
         TabIndex        =   95
         Top             =   6720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   65535
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   54
         Left            =   8820
         TabIndex        =   105
         Top             =   7200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
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
         Height          =   375
         Index           =   55
         Left            =   8820
         TabIndex        =   106
         Top             =   6840
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
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
         Height          =   315
         Index           =   49
         Left            =   5520
         TabIndex        =   40
         Top             =   7140
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   33
         Left            =   5040
         TabIndex        =   36
         Top             =   5820
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
         Index           =   85
         Left            =   12420
         TabIndex        =   121
         Top             =   7380
         Visible         =   0   'False
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
         Index           =   86
         Left            =   12360
         TabIndex        =   122
         Top             =   8880
         Visible         =   0   'False
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
         Height          =   435
         Index           =   478
         Left            =   9480
         TabIndex        =   128
         Top             =   660
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
         Index           =   460
         Left            =   10380
         TabIndex        =   129
         Top             =   180
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   479
         Left            =   8490
         TabIndex        =   130
         Top             =   1920
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
         Index           =   469
         Left            =   10680
         TabIndex        =   131
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
         Index           =   466
         Left            =   10680
         TabIndex        =   132
         Top             =   660
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
         Index           =   470
         Left            =   10440
         TabIndex        =   133
         Top             =   1200
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
         Height          =   435
         Index           =   477
         Left            =   9210
         TabIndex        =   134
         Top             =   1920
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
         Index           =   482
         Left            =   8820
         TabIndex        =   135
         Top             =   660
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
         Index           =   480
         Left            =   8460
         TabIndex        =   144
         Top             =   660
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
         Index           =   449
         Left            =   5850
         TabIndex        =   149
         Top             =   1920
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
         Index           =   450
         Left            =   5880
         TabIndex        =   150
         Top             =   660
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
         Index           =   436
         Left            =   3600
         TabIndex        =   151
         Top             =   660
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
         Index           =   439
         Left            =   3600
         TabIndex        =   152
         Top             =   1500
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
         Index           =   440
         Left            =   3960
         TabIndex        =   153
         Top             =   1200
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
         Height          =   435
         Index           =   447
         Left            =   5190
         TabIndex        =   154
         Top             =   1920
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
         Index           =   452
         Left            =   5520
         TabIndex        =   155
         Top             =   660
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
         Index           =   448
         Left            =   4860
         TabIndex        =   165
         Top             =   660
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
      Begin VB.Image Arrow_875 
         Height          =   255
         Index           =   4
         Left            =   2520
         Picture         =   "Feed_Water_Alone.frx":0000
         Stretch         =   -1  'True
         Top             =   7650
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "150 #"
         Height          =   285
         Left            =   2160
         TabIndex        =   183
         Top             =   6990
         Width           =   1065
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   960
         Picture         =   "Feed_Water_Alone.frx":030A
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   9
         X1              =   900
         X2              =   2520
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   10650
         Picture         =   "Feed_Water_Alone.frx":0614
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   10770
         X2              =   10770
         Y1              =   3660
         Y2              =   3450
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "35 # Dump"
         Height          =   315
         Index           =   8
         Left            =   10290
         TabIndex        =   177
         Top             =   3690
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   10320
         Picture         =   "Feed_Water_Alone.frx":091E
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "35 # Regulated"
         Height          =   255
         Index           =   7
         Left            =   9750
         TabIndex        =   176
         Top             =   2850
         Width           =   1365
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   10440
         X2              =   10440
         Y1              =   3420
         Y2              =   3060
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   0
         Left            =   10500
         Picture         =   "Feed_Water_Alone.frx":0C28
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   255
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   10500
         X2              =   10800
         Y1              =   6060
         Y2              =   5940
      End
      Begin VB.Line Line5 
         Index           =   3
         X1              =   10500
         X2              =   10800
         Y1              =   5820
         Y2              =   5700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   60
         X1              =   10620
         X2              =   10620
         Y1              =   5760
         Y2              =   4320
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   4
         Left            =   7020
         Picture         =   "Feed_Water_Alone.frx":0F32
         Stretch         =   -1  'True
         Top             =   7500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   3
         Left            =   7020
         Picture         =   "Feed_Water_Alone.frx":123C
         Stretch         =   -1  'True
         Top             =   8940
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   5
         Left            =   7500
         Picture         =   "Feed_Water_Alone.frx":1546
         Stretch         =   -1  'True
         Top             =   8520
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   2
         Left            =   9960
         Picture         =   "Feed_Water_Alone.frx":1850
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   1
         Left            =   9960
         Picture         =   "Feed_Water_Alone.frx":1B5A
         Stretch         =   -1  'True
         Top             =   7020
         Width           =   255
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2175
         Index           =   2
         Left            =   8040
         Top             =   6900
         Width           =   75
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2235
         Index           =   1
         Left            =   9600
         Top             =   5400
         Width           =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   33
         X1              =   7860
         X2              =   10620
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   30
         X1              =   7920
         X2              =   9480
         Y1              =   7620
         Y2              =   7620
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   0
         X1              =   8520
         X2              =   9360
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   2
         Left            =   4320
         Picture         =   "Feed_Water_Alone.frx":1E64
         Stretch         =   -1  'True
         Top             =   8940
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   1
         Left            =   4320
         Picture         =   "Feed_Water_Alone.frx":216E
         Stretch         =   -1  'True
         Top             =   7500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   4
         Left            =   8910
         Picture         =   "Feed_Water_Alone.frx":2478
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2880
         Top             =   10740
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   6
         Left            =   7830
         Picture         =   "Feed_Water_Alone.frx":2782
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   5
         Left            =   7740
         Picture         =   "Feed_Water_Alone.frx":2A8C
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   65
         X1              =   7470
         X2              =   9930
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   4
         Left            =   1140
         Picture         =   "Feed_Water_Alone.frx":2D96
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   3
         Left            =   780
         Picture         =   "Feed_Water_Alone.frx":30A0
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   2
         Left            =   1140
         Picture         =   "Feed_Water_Alone.frx":33AA
         Stretch         =   -1  'True
         Top             =   7740
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   1
         Left            =   780
         Picture         =   "Feed_Water_Alone.frx":36B4
         Stretch         =   -1  'True
         Top             =   7740
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   64
         X1              =   7680
         X2              =   7680
         Y1              =   3600
         Y2              =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   63
         X1              =   11820
         X2              =   11820
         Y1              =   10680
         Y2              =   9240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   59
         X1              =   7380
         X2              =   7380
         Y1              =   3420
         Y2              =   960
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   6750
         Picture         =   "Feed_Water_Alone.frx":39BE
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Drum"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   170
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Boiler"
         Height          =   255
         Index           =   5
         Left            =   1230
         TabIndex        =   169
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Economiser"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   168
         Top             =   1320
         Width           =   915
      End
      Begin VB.Image Image16 
         Height          =   255
         Index           =   4
         Left            =   5550
         Picture         =   "Feed_Water_Alone.frx":3CC8
         Stretch         =   -1  'True
         Top             =   2100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Index           =   3
         Left            =   4560
         Picture         =   "Feed_Water_Alone.frx":3FD2
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   3
         Left            =   10680
         Picture         =   "Feed_Water_Alone.frx":42DC
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Index           =   2
         Left            =   3120
         Picture         =   "Feed_Water_Alone.frx":45E6
         Stretch         =   -1  'True
         Top             =   2100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Index           =   1
         Left            =   3360
         Picture         =   "Feed_Water_Alone.frx":48F0
         Stretch         =   -1  'True
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   2
         Left            =   10680
         Picture         =   "Feed_Water_Alone.frx":4BFA
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   1
         Left            =   9960
         Picture         =   "Feed_Water_Alone.frx":4F04
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00004080&
         BorderWidth     =   3
         Height          =   1215
         Index           =   1
         Left            =   1380
         Shape           =   3  'Circle
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   58
         X1              =   2460
         X2              =   4860
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   56
         X1              =   4830
         X2              =   7230
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   55
         X1              =   2220
         X2              =   4080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   54
         X1              =   4080
         X2              =   4080
         Y1              =   960
         Y2              =   2040
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Drum"
         Height          =   255
         Index           =   3
         Left            =   12060
         TabIndex        =   167
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Economiser"
         Height          =   255
         Index           =   2
         Left            =   11220
         TabIndex        =   166
         Top             =   1320
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   53
         X1              =   2220
         X2              =   2220
         Y1              =   1560
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   52
         X1              =   3240
         X2              =   3240
         Y1              =   1260
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   49
         X1              =   2220
         X2              =   2220
         Y1              =   960
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   48
         X1              =   2220
         X2              =   3240
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   47
         X1              =   2220
         X2              =   3240
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   46
         X1              =   2220
         X2              =   4080
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Feed Water Control"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   164
         Top             =   120
         Width           =   1875
      End
      Begin VB.OLE Feed_reg_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   5220
         OleObjectBlob   =   "Feed_Water_Alone.frx":520E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   160
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   14100
         TabIndex        =   159
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Feed"
         Height          =   255
         Left            =   6450
         TabIndex        =   158
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Feed"
         Height          =   255
         Left            =   6480
         TabIndex        =   157
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label V_feed_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   156
         Top             =   360
         Width           =   555
      End
      Begin VB.Line Advanced_line 
         X1              =   4500
         X2              =   4080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line3 
         X1              =   4380
         X2              =   4380
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port Boiler"
         Height          =   255
         Index           =   0
         Left            =   12240
         TabIndex        =   145
         Top             =   780
         Width           =   915
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00004080&
         BorderWidth     =   3
         Height          =   1215
         Index           =   0
         Left            =   11880
         Shape           =   3  'Circle
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   44
         X1              =   12180
         X2              =   12180
         Y1              =   1560
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   42
         X1              =   11160
         X2              =   11160
         Y1              =   1260
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   41
         X1              =   12180
         X2              =   12180
         Y1              =   960
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   40
         X1              =   11160
         X2              =   12180
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   39
         X1              =   11160
         X2              =   12180
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   38
         X1              =   10560
         X2              =   12180
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   37
         X1              =   10560
         X2              =   12240
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   36
         X1              =   9900
         X2              =   12000
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   34
         X1              =   10560
         X2              =   10560
         Y1              =   960
         Y2              =   2040
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Feed"
         Height          =   255
         Index           =   9
         Left            =   1380
         TabIndex        =   139
         Top             =   5280
         Width           =   795
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Feed"
         Height          =   315
         Index           =   10
         Left            =   13800
         TabIndex        =   138
         Top             =   3420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Feed Water Control"
         Height          =   255
         Index           =   11
         Left            =   8460
         TabIndex        =   137
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label V_feed_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   136
         Top             =   360
         Width           =   555
      End
      Begin VB.Line Line4 
         X1              =   10290
         X2              =   10560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   10260
         X2              =   10260
         Y1              =   720
         Y2              =   960
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   12
         X1              =   2640
         X2              =   2640
         Y1              =   8160
         Y2              =   7680
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   4
         X1              =   4020
         X2              =   6480
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Label BYPASS_feed_2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   124
         Top             =   6180
         Width           =   555
      End
      Begin VB.Label BYPASS_feed_1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5340
         TabIndex        =   123
         Top             =   4560
         Width           =   555
      End
      Begin VB.Label Tripped_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Tripped"
         Height          =   255
         Index           =   1
         Left            =   9180
         TabIndex        =   115
         Top             =   4140
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Left            =   9480
         TabIndex        =   114
         Top             =   4740
         Width           =   255
      End
      Begin VB.Label AUX_2 
         Caption         =   "# 2"
         Height          =   195
         Left            =   13920
         TabIndex        =   104
         Top             =   9900
         Width           =   255
      End
      Begin VB.Label AUX_1 
         Caption         =   "# 1"
         Height          =   195
         Left            =   12540
         TabIndex        =   103
         Top             =   9900
         Width           =   255
      End
      Begin VB.Label AUX_LO_label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Feed Pump Aux LO"
         Height          =   195
         Left            =   12540
         TabIndex        =   102
         Top             =   11040
         Width           =   1575
      End
      Begin VB.Label VV_feed_aux_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   6720
         TabIndex        =   89
         Top             =   10020
         Width           =   555
      End
      Begin VB.Label VV_feed_aux_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   88
         Top             =   10020
         Width           =   555
      End
      Begin VB.Label RECIRC_label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reciculation Valve # 2"
         Height          =   195
         Index           =   2
         Left            =   12360
         TabIndex        =   71
         Top             =   8400
         Width           =   1755
      End
      Begin VB.Label RECIRC_label 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reciculation Valve # 1"
         Height          =   195
         Index           =   1
         Left            =   12300
         TabIndex        =   69
         Top             =   7740
         Width           =   1755
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "DFT Water level"
         Height          =   615
         Left            =   13020
         TabIndex        =   61
         Top             =   4740
         Width           =   555
      End
      Begin VB.Image Image24 
         Height          =   255
         Left            =   6960
         Picture         =   "Feed_Water_Alone.frx":D826
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   31
         X1              =   1260
         X2              =   1260
         Y1              =   3600
         Y2              =   8310
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   29
         X1              =   1260
         X2              =   7680
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   19
         X1              =   900
         X2              =   7380
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   10860
         Picture         =   "Feed_Water_Alone.frx":DB30
         Stretch         =   -1  'True
         Top             =   9180
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   7
         X1              =   1260
         X2              =   1260
         Y1              =   10380
         Y2              =   8520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   900
         X2              =   900
         Y1              =   8400
         Y2              =   3420
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "H.P. Heater"
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
         Left            =   1320
         TabIndex        =   60
         Top             =   8820
         Width           =   1035
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No 1"
         Height          =   255
         Left            =   9360
         TabIndex        =   10
         Top             =   8160
         Width           =   495
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   255
         Left            =   7920
         TabIndex        =   52
         Top             =   6300
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   11400
         X2              =   8340
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   8580
         X2              =   8580
         Y1              =   7620
         Y2              =   7440
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   11700
         X2              =   12000
         Y1              =   7440
         Y2              =   7320
      End
      Begin VB.Line Line2 
         X1              =   11700
         X2              =   12000
         Y1              =   7200
         Y2              =   7080
      End
      Begin VB.Image Image6 
         Height          =   255
         Index           =   0
         Left            =   9780
         Picture         =   "Feed_Water_Alone.frx":DE3A
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   255
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   1
         X1              =   5340
         X2              =   8160
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   10860
         Picture         =   "Feed_Water_Alone.frx":E144
         Stretch         =   -1  'True
         Top             =   7740
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   17
         X1              =   8580
         X2              =   10620
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   6600
         X2              =   6600
         Y1              =   10380
         Y2              =   9210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   26
         X1              =   4860
         X2              =   4860
         Y1              =   10380
         Y2              =   9060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   25
         X1              =   6600
         X2              =   6600
         Y1              =   8940
         Y2              =   7620
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To ADT"
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         Top             =   10740
         Width           =   675
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   3360
         Picture         =   "Feed_Water_Alone.frx":E44E
         Stretch         =   -1  'True
         Top             =   10740
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   24
         X1              =   3000
         X2              =   3480
         Y1              =   10860
         Y2              =   10860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   23
         X1              =   2940
         X2              =   2940
         Y1              =   10680
         Y2              =   10500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   22
         X1              =   2940
         X2              =   2940
         Y1              =   10260
         Y2              =   9420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   11
         X1              =   12120
         X2              =   12780
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   12360
         X2              =   12840
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   8220
         Picture         =   "Feed_Water_Alone.frx":E758
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   16
         X1              =   6480
         X2              =   6480
         Y1              =   6840
         Y2              =   6120
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   6
         X1              =   4020
         X2              =   6480
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   14
         X1              =   6480
         X2              =   6480
         Y1              =   5280
         Y2              =   4500
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   12
         X1              =   4020
         X2              =   4020
         Y1              =   5280
         Y2              =   4500
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Feed pump reciculation"
         Height          =   435
         Left            =   14220
         TabIndex        =   28
         Top             =   4320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   5040
         X2              =   4020
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Main Feed"
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Aux  Feed"
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   10500
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   2760
         X2              =   3360
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   6
         X1              =   3360
         X2              =   3360
         Y1              =   9060
         Y2              =   7620
      End
      Begin VB.Label Label6 
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
         Height          =   315
         Left            =   11340
         TabIndex        =   21
         Top             =   4710
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   8160
         X2              =   11820
         Y1              =   9300
         Y2              =   9300
      End
      Begin VB.Image Image16 
         Height          =   255
         Index           =   0
         Left            =   11820
         Picture         =   "Feed_Water_Alone.frx":EA62
         Stretch         =   -1  'True
         Top             =   3780
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   5
         X1              =   9720
         X2              =   11820
         Y1              =   7860
         Y2              =   7860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   11820
         X2              =   11820
         Y1              =   7140
         Y2              =   5940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   11910
         X2              =   12420
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   11280
         Picture         =   "Feed_Water_Alone.frx":ED6C
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   360
         Left            =   11820
         Picture         =   "Feed_Water_Alone.frx":F076
         Stretch         =   -1  'True
         Top             =   3420
         Width           =   285
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   11400
         X2              =   11400
         Y1              =   4140
         Y2              =   3420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   43
         X1              =   11820
         X2              =   12720
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   61
         X1              =   11820
         X2              =   11820
         Y1              =   9300
         Y2              =   7380
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   0
         Left            =   1320
         Picture         =   "Feed_Water_Alone.frx":F380
         Stretch         =   -1  'True
         Top             =   10260
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "No 2"
         Height          =   255
         Left            =   7860
         TabIndex        =   11
         Top             =   9600
         Width           =   555
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Turbo Feed Pumps"
         Height          =   735
         Index           =   0
         Left            =   8700
         TabIndex        =   9
         Top             =   5640
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   0
         Left            =   3540
         Picture         =   "Feed_Water_Alone.frx":F68A
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   0
         X1              =   4020
         X2              =   4020
         Y1              =   6840
         Y2              =   5280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   21
         X1              =   10620
         X2              =   10620
         Y1              =   7440
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   20
         X1              =   7620
         X2              =   7620
         Y1              =   7500
         Y2              =   7140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   18
         X1              =   3360
         X2              =   7920
         Y1              =   9060
         Y2              =   9060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   16
         X1              =   10860
         X2              =   11820
         Y1              =   10680
         Y2              =   10680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   14
         X1              =   1260
         X2              =   9840
         Y1              =   10380
         Y2              =   10380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   13
         X1              =   3360
         X2              =   7980
         Y1              =   7620
         Y2              =   7620
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   12
         X1              =   7620
         X2              =   7920
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   3360
         X2              =   4020
         Y1              =   5700
         Y2              =   5700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   10620
         X2              =   11580
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Electric Motor"
         Height          =   255
         Left            =   10980
         TabIndex        =   5
         Top             =   10980
         Width           =   1095
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   13
         X1              =   4020
         X2              =   5040
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   11280
         Top             =   4140
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   5
         Left            =   11280
         Shape           =   3  'Circle
         Top             =   4860
         Width           =   1095
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   11460
         Shape           =   3  'Circle
         Top             =   3735
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   15
         X1              =   7620
         X2              =   7620
         Y1              =   7740
         Y2              =   9060
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":F994
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   58
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":17BAC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   57
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   6120
         OleObjectBlob   =   "Feed_Water_Alone.frx":205C4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   62
         Top             =   8100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   6120
         OleObjectBlob   =   "Feed_Water_Alone.frx":28DDC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   63
         Top             =   8100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   11220
         OleObjectBlob   =   "Feed_Water_Alone.frx":30BF4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   64
         Top             =   6900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   11220
         OleObjectBlob   =   "Feed_Water_Alone.frx":3940C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   65
         Top             =   7260
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label VV_pos_feed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   5295
         TabIndex        =   107
         Top             =   4920
         Width           =   555
      End
      Begin VB.Label VV_pos_feed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   5340
         TabIndex        =   108
         Top             =   6480
         Width           =   555
      End
      Begin VB.Label Tripped_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Tripped"
         Height          =   255
         Index           =   2
         Left            =   7740
         TabIndex        =   51
         Top             =   5640
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":41224
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   50
         Top             =   4920
         Width           =   315
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":4983C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   59
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":51654
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   53
         Top             =   6480
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":59C6C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   56
         Top             =   6480
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":61A84
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   55
         Top             =   6480
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   5040
         OleObjectBlob   =   "Feed_Water_Alone.frx":69C9C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   54
         Top             =   6480
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 psi DSH Steam"
         Height          =   615
         Left            =   2700
         TabIndex        =   24
         Top             =   5070
         Width           =   855
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   9180
         OleObjectBlob   =   "Feed_Water_Alone.frx":726B4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   140
         Top             =   600
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   2550
         Top             =   8520
         Width           =   195
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   8880
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   50
         X1              =   2520
         X2              =   2520
         Y1              =   8400
         Y2              =   9060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   51
         X1              =   2760
         X2              =   2760
         Y1              =   8400
         Y2              =   9060
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   2400
         Top             =   8160
         Width           =   555
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   2
         X1              =   5340
         X2              =   7800
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1155
         Left            =   7620
         OleObjectBlob   =   "Feed_Water_Alone.frx":7ACCC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisFdPmp.vsd"
         TabIndex        =   45
         Top             =   5940
         Width           =   915
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   8340
         X2              =   8340
         Y1              =   3420
         Y2              =   6180
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1155
         Left            =   9180
         OleObjectBlob   =   "Feed_Water_Alone.frx":826E4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisFdPmp.vsd"
         TabIndex        =   175
         Top             =   4440
         Width           =   915
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   9900
         X2              =   9900
         Y1              =   3420
         Y2              =   4920
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   9180
         OleObjectBlob   =   "Feed_Water_Alone.frx":8A0FC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   143
         Top             =   600
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   9180
         OleObjectBlob   =   "Feed_Water_Alone.frx":92B14
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   142
         Top             =   600
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   9180
         OleObjectBlob   =   "Feed_Water_Alone.frx":9AD2C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   141
         Top             =   600
         Width           =   315
      End
      Begin VB.OLE Feed_reg_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   5220
         OleObjectBlob   =   "Feed_Water_Alone.frx":A2B44
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   163
         Top             =   600
         Width           =   315
      End
      Begin VB.OLE Feed_reg_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   5220
         OleObjectBlob   =   "Feed_Water_Alone.frx":AB55C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   162
         Top             =   600
         Width           =   315
      End
      Begin VB.OLE Feed_reg_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   5220
         OleObjectBlob   =   "Feed_Water_Alone.frx":B3774
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   161
         Top             =   600
         Width           =   315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   45
         X1              =   2220
         X2              =   7380
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   2625
         Index           =   1
         Left            =   1140
         Shape           =   4  'Rounded Rectangle
         Top             =   660
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   32
         X1              =   3300
         X2              =   12180
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   2595
         Index           =   0
         Left            =   11100
         Shape           =   4  'Rounded Rectangle
         Top             =   660
         Width           =   2175
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2100
      TabIndex        =   16
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Feed_Water_Alone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub AUTO_feed_pump_Click(Index As Integer)



' 18/7/00   ADDED this:

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        
    Exit Sub
    
End If


' Try to put the feed pump into AUTO....

If Index = 1 Then
  Call setPCDI(23)
ElseIf Index = 2 Then
  Call setPCDI(120)
End If


End Sub

Private Sub AUTO_VV_Click(Index As Integer)


set_dummy_focus = True


' 17/10/99
Select Case Index

  Case 1
    Call setPCDI(11)
  Case 2
    Call setPCDI(99)
    
End Select


End Sub


Private Sub CLOSE_VV_Click(Index As Integer)

' === 13/11/02 ===
set_dummy_focus = True

' 2/11/99   CLOSE
Select Case Index

  Case 1
    Call setPCDI(12)
  Case 2
    Call setPCDI(100)
    
End Select


End Sub

Private Sub Command1_Click()

Condensate.Show

End Sub


Private Sub Command3_Click()


' 13/07/09
'If LSS_reals(80) > 25# Then
'  Gland_Steam_Full_away.Show
'Else
'  Gland_Steam.Show
'End If


' 13/07/09
Aux_Steam_Alone.ZOrder


End Sub





Private Sub Command2_Click()

Aux_Steam_Alone.ZOrder

End Sub

Private Sub Command4_Click()

Aux_Steam_Alone.ZOrder

End Sub

Private Sub Command5_Click()


' 13/07/09
Condensate.ZOrder

End Sub

Private Sub Command7_Click()

Stbd_Boiler.ZOrder

End Sub

Private Sub Command8_Click()

Port_Boiler.ZOrder

End Sub

Private Sub Command9_Click()

Aux_Steam_Alone.ZOrder

End Sub

Private Sub Drum_water_level_port_Click()

    Unload Alarm_set_point   ' Do this first
    
    nAlarm_HI = 52
    nAlarm_LO = 60  ' 6/1/00

    Alarm_set_point.Show


End Sub

Private Sub Drum_water_level_stbd_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 5
    
    nAlarm_LO = 13  ' 6/1/00
    
    Alarm_set_point.Show


End Sub

Private Sub Form_Load()

' 14/07/09 Removed so that objects which do NOT exist can be detected....
'On Error Resume Next

' 09/06/09
'Call JYAScreenInit(Me, idFeed_water)
Call JYAScreenInit(Me, idFeed_water_Alone)


' === 18/02/03 TESTING NO CRASH...===
'    RECIRC_label(55).Visible = False
' === 18/02/03 TESTING method or data member not found....===
'Valve(44).fred = 1
' === 18/02/03 TESTiNG Worked ?...
'Valve(44).OpenBtmRight = "Hello"
' === 18/02/03 TESTiNG Type Mis-match...
'Command1.Picture = "Hello"

' === 18/02/03 TESTiNG Worked OK ( also tested Steam_SimV22_test.exe and was OK )...
'VScroll1.Value = 255 * 4


' === 18/11/02 ===
update_VSCROLL(23) = True
update_VSCROLL(25) = True
update_VSCROLL(28) = True




' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' === 27/09/02 14.114 ===
    RECIRC_label(1).Visible = False
    RECIRC_label(2).Visible = False
    OPEN_VV(1).Visible = False
    OPEN_VV(2).Visible = False
    CLOSE_VV(1).Visible = False
    CLOSE_VV(2).Visible = False
    AUTO_VV(1).Visible = False
    AUTO_VV(2).Visible = False
    ' =======================
    
        ' === 10/04/02 S Green e-mail 30/3/02 point 2b ===
        REMOTE_feed_control(1).Visible = False
        REMOTE_feed_control(2).Visible = False
    
    
     ' 11/12/00
     ' 19/07/09 REMOVED...
     'BYPASS_IP.Visible = False
     'BYPASS_875_35.Visible = False
     'BYPASS_HP.Visible = False
     'BYPASS_atom.Visible = False
     'BYPASS_875_150.Visible = False
     'BYPASS_35_dump.Visible = False
     
     BYPASS_feed_1.Visible = False
     BYPASS_feed_2.Visible = False
   
   
    ' 11/12/00
    ' 14/07/09 HScroll1.Visible = False
    
    
    ' === 18/9/00  Point 8.126 ===
    
    ' === 01/03/01 Point 10.2b NOT the Elec Pump ===
    'ONpump(0).Visible = False
    'OFFpump(0).Visible = False
    'sbyPUMP(0).Visible = False

    ONpump(3).Visible = False
    OFFpump(3).Visible = False
    sbyPUMP(3).Visible = False
    
    ONpump(4).Visible = False
    OFFpump(4).Visible = False
    sbyPUMP(4).Visible = False
    
    AUX_LO_label.Visible = False
    AUX_1.Visible = False
    AUX_2.Visible = False
    
    ' 7/9/00
    TLO_main_feed(1).Visible = False
    TLO_main_feed(2).Visible = False
    
    ' 10/12/00
    PLO_total_feed_pump(1).Visible = False
    PLO_total_feed_pump(2).Visible = False
    
    ' 4/5/00
    ' 19/07/09 REMOVED...
    'SP_HP.Visible = False
    'SP_875_150.Visible = False
    'SP_IP.Visible = False
    'SP_875_35.Visible = False
    
    
    ' === 21/9/00 ===
    ' 19/07/09 REMOVED...
    '     VV_pos_875_35_reg.Visible = False
    '    VV_pos_875_150_reg.Visible = False
    '         VV_pos_IP_reg.Visible = False
    '         VV_pos_HP_reg.Visible = False
    '       VV_pos_dump_reg.Visible = False
        
        VV_pos_feed(1).Visible = False
        VV_pos_feed(2).Visible = False
        
        VV_feed_aux_pos(1).Visible = False
        VV_feed_aux_pos(2).Visible = False
        
    
Else

    ' === 27/09/02 14.114 ===
    RECIRC_label(1).Visible = True
    RECIRC_label(2).Visible = True
    OPEN_VV(1).Visible = True
    OPEN_VV(2).Visible = True
    CLOSE_VV(1).Visible = True
    CLOSE_VV(2).Visible = True
    AUTO_VV(1).Visible = True
    AUTO_VV(2).Visible = True
    ' =======================

        ' === 10/04/02 S Green e-mail 30/3/02 point 2b ===
        REMOTE_feed_control(1).Visible = True
        REMOTE_feed_control(2).Visible = True

     ' 11/12/00
     ' 19/07/09 REMOVED...
     'BYPASS_IP.Visible = True
     'BYPASS_875_35.Visible = True
     'BYPASS_HP.Visible = True
     'BYPASS_atom.Visible = True
     'BYPASS_875_150.Visible = True
     'BYPASS_35_dump.Visible = True
     
     BYPASS_feed_1.Visible = True
     BYPASS_feed_2.Visible = True



    ' 11/12/00
   ' 14/07/09 HScroll1.Visible = True

    ' === 18/9/00  Point 8.126 ===
    ONpump(0).Visible = True
    OFFpump(0).Visible = True
    sbyPUMP(0).Visible = True


    ONpump(3).Visible = True
    OFFpump(3).Visible = True
    sbyPUMP(3).Visible = True
    
    ONpump(4).Visible = True
    OFFpump(4).Visible = True
    sbyPUMP(4).Visible = True
    
    AUX_LO_label.Visible = True
    AUX_1.Visible = True
    AUX_2.Visible = True
    
    ' 7/9/00
    TLO_main_feed(1).Visible = True
    TLO_main_feed(2).Visible = True

    ' 10/12/00
    PLO_total_feed_pump(1).Visible = True
    PLO_total_feed_pump(2).Visible = True

    
    ' 4/5/00
    ' 19/07/09 REMOVED...
    'SP_HP.Visible = True
    'SP_875_150.Visible = True
    'SP_IP.Visible = True
    'SP_875_35.Visible = True
    
    
    ' === 21/9/00 ===
    ' 19/07/09 REMOVED...
    '     VV_pos_875_35_reg.Visible = True
    '    VV_pos_875_150_reg.Visible = True
    '         VV_pos_IP_reg.Visible = True
    '         VV_pos_HP_reg.Visible = True
    '       VV_pos_dump_reg.Visible = True
        
        
        VV_pos_feed(1).Visible = True
        VV_pos_feed(2).Visible = True
        
        VV_feed_aux_pos(1).Visible = True
        VV_feed_aux_pos(2).Visible = True


End If


' 19/07/09 REMOVED...
'VScroll1.Visible = PTT

' 14/07/09 HScroll2.Visible = PTT

' 19/07/09 REMOVED...
'HScroll3.Visible = PTT

HScroll5.Visible = PTT
HScroll4.Visible = PTT

HScroll9.Visible = PTT

HScroll7.Visible = PTT

HScroll6.Visible = PTT

'update_VBAI(6) = True


' === 13/11/02 ===
update_VBAI(6) = True
update_VBAI(7) = True
update_VBAI(8) = True

' === 13/11/02 ===
update_VBAI(12) = True
update_VBAI(13) = True
update_VBAI(14) = True
update_VBAI(15) = True
update_VBAI(16) = True

' === 14/07/09 STBD Boiler Aux Feed valve ===
update_VBAI(4) = True
HScroll_Aux_feed_STBD.Visible = PTT

' === 14/07/09 PORT Boiler Aux Feed valve ===
update_VBAI(5) = True
HScroll_Aux_feed_PORT.Visible = PTT


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then

    HavenStretch1.Visible = False
    
    
    ' 09/06/09
   ' Call JYAScreenResize(Me, idFeed_water)
    Call JYAScreenResize(Me, idFeed_water_Alone)
    
    
    HavenStretch1.Visible = True
    
    
    
   
End If

End Sub

Private Sub HScroll_Aux_feed_PORT_Change()

Dim AIvalue As Integer

' 5/1/00   Port Aux Feed Valve(479)
' VALVE is now variable ( used to be clickable )

' 14/07/09
AIvalue = HScroll_Aux_feed_PORT.Value / 4
Call setAI(5, AIvalue)


End Sub

Private Sub HScroll_Aux_feed_STBD_Change()

Dim AIvalue As Integer

' 5/1/00   Stbd Aux Feed Valve(449)
' VALVE is now variable ( used to be clickable )


' 14/07/09
AIvalue = HScroll_Aux_feed_STBD.Value / 4
Call setAI(4, AIvalue)


End Sub

Private Sub HScroll3_Change()
Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  Eight75_35_bypass
' 1/12/99
'Call set_VSCROLL(22, HScroll3.Value)

' 1/12/99
AIvalue = HScroll3.Value / 4
Call setAI(8, AIvalue)


End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub HScroll2_Change()

End Sub

Private Sub HScroll4_Change()
Dim AIvalue As Integer


' 3/11/99  Aux feed from Main pump # 2

' 1/12/99
'Call set_VSCROLL(20, HScroll4.Value)
AIvalue = HScroll4.Value / 4
Call setAI(13, AIvalue)


End Sub

Private Sub HScroll5_Change()
Dim AIvalue As Integer

' 3/11/99  Aux feed from Main pump # 1

' 1/12/99
'Call set_VSCROLL(19, HScroll5.Value)
AIvalue = HScroll5.Value / 4
Call setAI(12, AIvalue)


End Sub

Private Sub HScroll6_Change()
Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)

' 6/11/99  Feed_pump_bypass   # 1

' 1/12/99
'Call set_VSCROLL(26, HScroll6.Value)

AIvalue = HScroll6.Value / 4
Call setAI(16, AIvalue)



End Sub

Private Sub HScroll7_Change()
Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)

' 6/11/99  Feed_pump_bypass   # 2

' 1/12/99
'Call set_VSCROLL(27, HScroll7.Value)
AIvalue = HScroll7.Value / 4
Call setAI(15, AIvalue)


End Sub

Private Sub HScroll8_Change()


'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)

' 6/11/99  Atom_bypass

' 7/12/00
'Call set_VSCROLL(28, HScroll8.Value)
Call setVscroll(28, HScroll8.Value)

End Sub

Private Sub HScroll9_Change()
Dim AIvalue As Integer

' 1/12/99  VALVE is now variable ( used to be clickable )

AIvalue = HScroll9.Value / 4
Call setAI(14, AIvalue)


End Sub

Private Sub LO_feed_remote_Click(Index As Integer)

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)
If Index = 0 Then
'  11/10/99
'   Call setPCDI(PumpIO_LSS(4, 3))
    Call setPCDI(PumpIO_LSS(4, 9))
    
ElseIf Index = 3 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(4, 6))       ' Feed pump AUX LO # 1
        
ElseIf Index = 4 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(4, 18))      ' Feed pump AUX LO  # 2

End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)
If Index = 0 Then
    '  11/10/99
    '   Call setPCDI(PumpIO_LSS(2, 3))
        Call setPCDI(PumpIO_LSS(2, 9))      ' Feed Water Elec pump
    
ElseIf Index = 3 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(2, 6))       ' Feed pump AUX LO # 1
        
ElseIf Index = 4 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(2, 18))      ' Feed pump AUX LO  # 2

End If

End Sub

Private Sub OPEN_VV_Click(Index As Integer)

set_dummy_focus = True


' 2/11/99   OPEN
Select Case Index

  Case 1
    Call setPCDI(10)
  Case 2
    Call setPCDI(98)
    
End Select


End Sub

Private Sub P_MF_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 75
Alarm_set_point.Show

End Sub

Private Sub PLO_total_feed_pump_Click(Index As Integer)

' 27/8/00
Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_LO = 78
Else
    nAlarm_LO = 79
End If

Alarm_set_point.Show

End Sub

Private Sub Psteam1_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 4
    nAlarm_LO = 12  ' 6/1/00
    Alarm_set_point.Show

End Sub

Private Sub Psteam101_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 53
    nAlarm_LO = 61  ' 6/1/99
    Alarm_set_point.Show

End Sub

Private Sub REMOTE_feed_control_Click(Index As Integer)

set_dummy_focus = True

' 17/8/00  If on the REAL sim then this is a HARD switch in the ER

' 17/8/00 If PTT then do the following...
If PTT Then
    If Index = 1 Then
        Call setPCDI(356)
    Else
        Call setPCDI(357)
    End If
End If

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(5=Stand-bye,nPump)
If Index = 0 Then
'  11/10/99
'   Call setPCDI(PumpIO_LSS(5, 3))
    Call setPCDI(PumpIO_LSS(5, 9))
    
ElseIf Index = 3 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(5, 6))       ' Feed pump AUX LO # 1
        
ElseIf Index = 4 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(5, 18))      ' Feed pump AUX LO  # 2

End If

End Sub


Private Sub SP_875_150_Click()

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 257

Alarm_set_point.Show

End Sub

Private Sub SP_875_35_Click()

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 259

Alarm_set_point.Show

End Sub

Private Sub SP_HP_Click()


'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 256

Alarm_set_point.Show

End Sub

Private Sub SP_IP_Click()

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 258

Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim DFT_level As Integer
Dim green_colour As Integer
Dim VV_pos As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' === 14/07/09 STBD Boiler Aux Feed valve ===
If update_VBAI(4) Then
     update_VBAI(4) = False
    Call update_the_slider(HScroll_Aux_feed_STBD, VB_AI(4) * 4)
End If
' === 14/07/09 PORT Boiler Aux Feed valve ===
If update_VBAI(5) Then
     update_VBAI(5) = False
    Call update_the_slider(HScroll_Aux_feed_PORT, VB_AI(5) * 4)
End If
' =======================

' === 22/06/09 Port Boiler  ===
Call update_vvalve_LSS(Me, 466, 466, -1)
Call update_vvalve_LSS(Me, 469, 469, -1)
Call update_vvalve_LSS(Me, 470, 470, -1)

For i = 477 To 478
    Call update_vvalve_LSS(Me, i, i, -1)
Next i
Call update_vvalve_LSS(Me, 480, 480, -1)
Call update_vvalve_LSS(Me, 482, 482, -1)
Call update_green_VV(484, valve(479))

V_feed_pos(2).Caption = Format(LSS_reals(341), "###0") & " %"
If LSS_reals(341) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(341) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(341) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Feed_reg(i).Visible = True
    Else
        Feed_reg(i).Visible = False
    End If
Next i

Call paint_alarm(Me.Psteam101, 53, 61, 0, 0)
Call paint_alarm(Me.Drum_water_level_port, 52, 60, 0, 0)


' === 22/06/09 Stbd Boiler ===
Call update_vvalve_LSS(Me, 436, 436, -1)
Call update_vvalve_LSS(Me, 439, 439, -1)
Call update_vvalve_LSS(Me, 440, 440, -1)

For i = 447 To 448
    Call update_vvalve_LSS(Me, i, i, -1)
Next i
Call update_vvalve_LSS(Me, 450, 450, -1)
Call update_vvalve_LSS(Me, 452, 452, -1)
Call update_green_VV(483, valve(449))

V_feed_pos(1).Caption = Format(LSS_reals(340), "###0") & " %"
If LSS_reals(340) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(340) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(340) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Feed_reg_stbd(i).Visible = True
    Else
        Feed_reg_stbd(i).Visible = False
    End If
Next i

Call paint_alarm(Me.Psteam1, 4, 12, 0, 0)
Call paint_alarm(Me.Drum_water_level_stbd, 5, 13, 0, 0)

    ' STBD
T_econ_IN(1).text = Format(LSS_reals(386) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_main_feed(1).text = Format(LSS_reals(390) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

    ' PORT
T_econ_IN(2).text = Format(LSS_reals(387) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_main_feed(2).text = Format(LSS_reals(391) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char



If SI_units Then
    ' STBD
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult, "###0.0") & Flow_chars
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    Drum_water_level_stbd.text = Format(LSS_reals(115) * 25.4, "###0") & " mm"
    ' PORT
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult, "###0.0") & Flow_chars
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    Drum_water_level_port.text = Format(LSS_reals(116) * 25.4, "###0") & " mm"

   
Else

     ' STBD
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult / 3600#, "###0.0") & " lb/s"
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    Drum_water_level_stbd.text = Format(LSS_reals(115), "##0.0") & " in"
    
    ' PORT
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult / 3600#, "###0.0") & " lb/s"
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    Drum_water_level_port.text = Format(LSS_reals(116), "##0.0") & " in"

End If





' === 18/11/02 ===
' === 18/11/02 ==============
'equivalence(STEAM_integers(46),Vscroll_value(1))

If update_VSCROLL(25) Then
   update_VSCROLL(25) = False
   
   ' === 19/02/03 ===
   'HScroll1.Value = VB_steam_integers(70)
   ' 09/06/09 Call update_the_slider(HScroll1, VB_steam_integers(70))
   
End If
If update_VSCROLL(23) Then
   update_VSCROLL(23) = False
   
   ' === 19/02/03 ===
   'VScroll5.Value = VB_steam_integers(68)
   ' 09/06/09 Call update_the_slider(VScroll5, VB_steam_integers(68))
   
End If
If update_VSCROLL(28) Then
   update_VSCROLL(28) = False
   
   ' === 19/02/03 ===
   'HScroll8.Value = VB_steam_integers(73)
   ' 09/06/09 Call update_the_slider(HScroll8, VB_steam_integers(73))
   
End If


' === 13/11/02 =======
If update_VBAI(6) Then
     update_VBAI(6) = False
     
     ' === 19/02/03 ===
     'VScroll1.Value = VB_AI(6) * 4
     ' 09/06/09 Call update_the_slider(VScroll1, VB_AI(6) * 4)
     
End If
If update_VBAI(7) Then
     update_VBAI(7) = False
     
     'HScroll2.Value = VB_AI(7) * 4
     ' 09/06/09 Call update_the_slider(HScroll2, VB_AI(7) * 4)
     
End If
If update_VBAI(8) Then
     update_VBAI(8) = False
     
     'HScroll3.Value = VB_AI(8) * 4
     ' 09/06/09 Call update_the_slider(HScroll3, VB_AI(8) * 4)
     
End If
If update_VBAI(12) Then
     update_VBAI(12) = False
     
     'HScroll5.Value = VB_AI(12) * 4
     Call update_the_slider(HScroll5, VB_AI(12) * 4)
     
End If
If update_VBAI(13) Then
     update_VBAI(13) = False
     
     'HScroll4.Value = VB_AI(13) * 4
     Call update_the_slider(HScroll4, VB_AI(13) * 4)
     
End If
If update_VBAI(14) Then
     update_VBAI(14) = False
     
     'HScroll9.Value = VB_AI(14) * 4
     Call update_the_slider(HScroll9, VB_AI(14) * 4)
     
End If
If update_VBAI(15) Then
     update_VBAI(15) = False
     
     'HScroll7.Value = VB_AI(15) * 4
     Call update_the_slider(HScroll7, VB_AI(15) * 4)
     
End If
If update_VBAI(16) Then
     update_VBAI(16) = False
     
     ' === 19/02/03 ===
     'HScroll6.Value = VB_AI(16) * 4
     Call update_the_slider(HScroll6, VB_AI(16) * 4)
     ' TESTING...
     'i = HScroll6.Value
  
End If
' === 13/11/02 END ===


' 17/8/00
If VB_steam_logicals(123) Then
    REMOTE_feed_control(1).BackColor = &HFFFF&          ' Yellow
Else
    REMOTE_feed_control(1).BackColor = &HC0C0C0         ' Grey
End If
If VB_steam_logicals(124) Then
    REMOTE_feed_control(2).BackColor = &HFFFF&          ' Yellow
Else
    REMOTE_feed_control(2).BackColor = &HC0C0C0         ' Grey
End If


' 18/9?00 Point 8.126
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    Tripped_lbl(1).Visible = False
    Tripped_lbl(2).Visible = False
Else

  If VB_steam_logicals(125) Then
    Tripped_lbl(1).Visible = True
  Else
    Tripped_lbl(1).Visible = False
  End If
  If VB_steam_logicals(126) Then
    Tripped_lbl(2).Visible = True
  Else
    Tripped_lbl(2).Visible = False
  End If

End If



' 15/8/00
P_main_out(1).text = Format(LSS_reals(344) * Press_mult, "##0.0") & Press_chars

' 5/9/00
P_main_out(2).text = Format(LSS_reals(345) * Press_mult, "##0.0") & Press_chars


' 19/7/00....

If LSS_LOGICALS(61) Then
    AUTO_feed_pump(1).BackColor = &HFFFF&         '  YELLOW
Else
    AUTO_feed_pump(1).BackColor = &HC0C0C0     'GREY
End If

If LSS_LOGICALS(62) Then
    AUTO_feed_pump(2).BackColor = &HFFFF&         '  YELLOW
Else
    AUTO_feed_pump(2).BackColor = &HC0C0C0     'GREY
End If






'   2/5/00
' LO pump for Feed pump # 1
Call update_pump_LSS(Me, 3, 6)      ' Feed pump AUX LO  # 1
' LO pump for Feed pump # 1
Call update_pump_LSS(Me, 4, 18)     ' Feed pump AUX LO  # 2


' 27/8/00
Call paint_alarm(Me.PLO_total_feed_pump(1), 78, 0, 0, 0)
Call paint_alarm(Me.PLO_total_feed_pump(2), 79, 0, 0, 0)


' 7/1/00
Call paint_alarm(Me.P_MF, 75, 0, 0, 0)
Call paint_alarm(Me.Z_DFT, 76, 81, 0, 0)

' 10/1/00
'    PC_REALS(398) = F_feed_pump(1)
'    PC_REALS(399) = F_feed_pump(2)
If (LSS_reals(398) > 0.02) Then
  HavenPump1(1).PumpOn = True
Else
  HavenPump1(1).PumpOn = False
End If
If (LSS_reals(399) > 0.02) Then
  HavenPump1(2).PumpOn = True
Else
  HavenPump1(2).PumpOn = False
End If


'  1/12/99
'If Not PTT Then
'
'    If update_VBAI(6) Then
'       update_VBAI(6) = False
'        VScroll1.Value = VB_AI(6) * 4  ' VB_AI(6) = AI(6) = HP bypass
'    End If
'End If



'          PC_REALS(342) = P_MF
'          PC_REALS(343) = P_AF
'          PC_REALS(344) = P_main_out(1)
'          PC_REALS(345) = P_main_out(2)
'          PC_REALS(346) = P_MAIN(1)
'          PC_REALS(347) = P_MAIN(2)
'          PC_REALS(348) = F_MAIN(1)
'          PC_REALS(349) = F_MAIN(2)
'          PC_REALS(350) = F_MF(1)
'          PC_REALS(351) = F_MF(2)
'          PC_REALS(352) = F_MF_recirc(1)
'          PC_REALS(353) = F_MF_recirc(2)
'          PC_REALS(354) = F_MF_aux(1)
'          PC_REALS(355) = F_MF_aux(2)
'          PC_REALS(356) = VV_feed_aux_pos(1) * 100#
'          PC_REALS(357) = VV_feed_aux_pos(2) * 100#
'          PC_REALS(358) = F_main_feed(1)
'          PC_REALS(359) = F_main_feed(2)
'          PC_REALS(360) = F_aux_feed(1)
'          PC_REALS(361) = F_aux_feed(2)

'          PC_REALS(362) =  P_aux_feed_OUT

'          PC_REALS(309) = T_DFT
'          PC_REALS(310) = T_HP_htr
'          PC_REALS(311) = T_DWT
'          PC_REALS(312) = T_ADT

T_DFT.text = Format(LSS_reals(309) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_HP_htr.text = Format(LSS_reals(310) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
    
    
'          PC_REALS(155) = TLO_main_feed(1)
'          PC_REALS(156) = TLO_main_feed(2)
' 7/9/00
TLO_main_feed(1).text = Format(LSS_reals(155) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Call paint_alarm(Me.TLO_main_feed(1), 83, 0, 0, 0)
TLO_main_feed(2).text = Format(LSS_reals(156) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Call paint_alarm(Me.TLO_main_feed(2), 84, 0, 0, 0)
    
' 12/9/00
'09/06/09 Tsteam11.text = Format(LSS_reals(71) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
    

 '09/06/09    P_150.Visible = Advanced_mode
 '09/06/09    P_875_desup_OUT.Visible = Advanced_mode
 '09/06/09    P_HP_reg_IN.Visible = Advanced_mode
 '09/06/09    P_IP_reg_IN.Visible = Advanced_mode
 '09/06/09    P_35.Visible = Advanced_mode
 '09/06/09    P_Atom.Visible = Advanced_mode
    
   ' 16/8/00
   
   ' 1/10/00
    P_main(1).Visible = Advanced_mode
    P_main(2).Visible = Advanced_mode
   ' P_main(1).Visible = False
   ' P_main(2).Visible = False
    
    ' 15/8/00
   ' F_main(1).Visible = Advanced_mode
   ' F_main(2).Visible = Advanced_mode
   ' F_MF(1).Visible = Advanced_mode
   ' F_MF(2).Visible = Advanced_mode
    F_main(1).Visible = False
    F_main(2).Visible = False
    F_MF(1).Visible = False
    F_MF(2).Visible = False
    
    F_MF_recirc(1).Visible = Advanced_mode
    F_MF_recirc(2).Visible = Advanced_mode
    
    F_MF_aux(1).Visible = Advanced_mode
    F_MF_aux(2).Visible = Advanced_mode
    
    F_main_feed(1).Visible = Advanced_mode
    F_main_feed(2).Visible = Advanced_mode
    
    F_aux_feed(1).Visible = Advanced_mode
    F_aux_feed(2).Visible = Advanced_mode
    

'  === 6/11/99 Regulator and Bypass Valves ===

'    PC_REALS(364) = Eight75_150_reg * 100#
'    PC_REALS(365) = Eight75_150_bypass * 100#
'
'    PC_REALS(366) = Eight75_35_reg * 100#
'    PC_REALS(367) = Eight75_35_bypass * 100#
'
'    PC_REALS(368) = Dump_35_reg * 100#
'    PC_REALS(369) = Dump_35_bypass * 100#
'
'    PC_REALS(370) = HP_reg * 100#
'    PC_REALS(371) = HP_bypass * 100#
'
'    PC_REALS(372) = IP_reg * 100#
'    PC_REALS(373) = IP_bypass * 100#
'
'''' PC_REALS(374) = Feed_pump_reg(1) * 100#
'    PC_REALS(374) =  VV_pos_feed_pump(1) * 100.0

'    PC_REALS(375) = Feed_pump_bypass(1) * 100#
'
'''' PC_REALS(376) = Feed_pump_reg(2) * 100#
'    PC_REALS(376) =  VV_pos_feed_pump(2) * 100.0

'    PC_REALS(377) = Feed_pump_bypass(2) * 100#
'
'    PC_REALS(378) = Atom_reg * 100#
'    PC_REALS(379) = Atom_bypass * 100#

'    PC_REALS(380) = P_150
'    PC_REALS(381) = P_875_desup_OUT
'    PC_REALS(382) = P_HP_reg_IN
'    PC_REALS(383) = P_IP_reg_IN
'    PC_REALS(384) = P_35
'    PC_REALS(385) = P_Atom


'09/06/09 VV_pos_875_150_reg.Caption = Format(LSS_reals(364), "##0") & " %"

'09/06/09 VV_pos_875_35_reg.Caption = Format(LSS_reals(366), "##0") & " %"

'09/06/09 VV_pos_dump_reg.Caption = Format(LSS_reals(368), "##0") & " %"

'09/06/09 VV_pos_HP_reg.Caption = Format(LSS_reals(370), "##0") & " %"

'09/06/09 VV_pos_IP_reg.Caption = Format(LSS_reals(372), "##0") & " %"


' 5/6/00
VV_pos_feed(1).Caption = Format(LSS_reals(374), "##0") & " %"
VV_pos_feed(2).Caption = Format(LSS_reals(376), "##0") & " %"


'  === 8/11/99 Regulator  Valves ===
' 09/06/09 Call update_regulator(364, Reg_875_150(0), Reg_875_150(1), Reg_875_150(2), Reg_875_150(3))
' 09/06/09 Call update_regulator(366, Reg_875_35(0), Reg_875_35(1), Reg_875_35(2), Reg_875_35(3))
' 09/06/09 Call update_regulator(368, Dump_35(0), Dump_35(1), Dump_35(2), Dump_35(3))
' 09/06/09 Call update_regulator(370, Reg_HP(0), Reg_HP(1), Reg_HP(2), Reg_HP(3))
' 09/06/09 Call update_regulator(372, Reg_IP(0), Reg_IP(1), Reg_IP(2), Reg_IP(3))
Call update_regulator(374, Reg_pump_1(0), Reg_pump_1(1), Reg_pump_1(2), Reg_pump_1(3))
Call update_regulator(376, Reg_pump_2(0), Reg_pump_2(1), Reg_pump_2(2), Reg_pump_2(3))
' 09/06/09 Call update_regulator(378, Reg_Atomising(0), Reg_Atomising(1), Reg_Atomising(2), Reg_Atomising(3))


' 26/8/00
'       PC_REALS(117) = VV_main_feed_pos(2) * 100#
Call update_green_VV(117, valve(5))


'  === 6/11/99 Bypass  Valves ===

'    PC_REALS(365) = Eight75_150_bypass * 100#
' 09/06/09 Call update_green_VV(365, Valve(20))
' 11/12/00
' 09/06/09 BYPASS_875_150.Caption = Format(LSS_reals(365), "##0") & " %"


'    PC_REALS(367) = Eight75_35_bypass * 100#
' 09/06/09 Call update_green_VV(367, Valve(39))
' 11/12/00
' 09/06/09 BYPASS_875_35.Caption = Format(LSS_reals(367), "##0") & " %"

' PC_REALS(369) = Dump_35_bypass * 100#
' 09/06/09 Call update_green_VV(369, Valve(21))
' 11/12/00
' 09/06/09 BYPASS_35_dump.Caption = Format(LSS_reals(369), "##0") & " %"


'    PC_REALS(371) = HP_bypass * 100#
' 09/06/09 Call update_green_VV(371, Valve(23))
' 11/12/00
' 09/06/09 BYPASS_HP.Caption = Format(LSS_reals(371), "##0") & " %"

'    PC_REALS(373) = IP_bypass * 100#
' 09/06/09 Call update_green_VV(373, Valve(43))
' 11/12/00
' 09/06/09 BYPASS_IP.Caption = Format(LSS_reals(373), "##0") & " %"



'        PC_REALS(375) = Feed_pump_bypass(1) * 100#
Call update_green_VV(375, valve(32))
' 11/12/00
BYPASS_feed_1.Caption = Format(LSS_reals(375), "##0") & " %"

'        PC_REALS(377) = Feed_pump_bypass(2) * 100#
Call update_green_VV(377, valve(33))
' 11/12/00
BYPASS_feed_2.Caption = Format(LSS_reals(377), "##0") & " %"



'        PC_REALS(379) = Atom_bypass * 100#
' 09/06/09 Call update_green_VV(379, Valve(47))
' 11/12/00
' 09/06/09 BYPASS_atom.Caption = Format(LSS_reals(379), "##0") & " %"


' Main Feed #1 to AUX feed valve...
VV_feed_aux_pos(1).Caption = Format(LSS_reals(356), "##0") & " %"

Call update_green_VV(356, valve(12))

'If LSS_reals(356) > 99# Then
'    green_colour = 255
'ElseIf LSS_reals(356) < 1# Then
'    green_colour = 0
'Else
'    green_colour = LSS_reals(356) * 2.55
'End If
'Valve(12).ColourClosed = RGB(0, green_colour, 0)


' Main Feed #2 to AUX feed valve...
VV_feed_aux_pos(2).Caption = Format(LSS_reals(357), "##0") & " %"

Call update_green_VV(357, valve(51))

'If LSS_reals(357) > 99# Then
'    green_colour = 255
'ElseIf LSS_reals(357) < 1# Then
'    green_colour = 0
'Else
'    green_colour = LSS_reals(357) * 2.55
'End If
'Valve(51).ColourClosed = RGB(0, green_colour, 0)



If SI_units Then
    
    ' 17/8/00
    PLO_total_feed_pump(1).text = Format(LSS_reals(486) * Press_mult, "##0.0") & Press_chars
    PLO_total_feed_pump(2).text = Format(LSS_reals(487) * Press_mult, "##0.0") & Press_chars

    Z_DFT.text = Format(LSS_reals(451) * 25.4, "###0.0") & " mm"

    P_aux_feed_out.text = Format(LSS_reals(362) * Press_mult, "##0.0") & Press_chars
  
    ' 22/06/09 Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0.0") & Press_chars
    ' 22/06/09 Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0.0") & Press_chars
    
    P_MF.text = Format(LSS_reals(342) * Press_mult, "##0.0") & Press_chars
    P_AF.text = Format(LSS_reals(343) * Press_mult, "##0.0") & Press_chars

  If (Advanced_mode) Then
  
    ' 22/06/09 P_150.text = Format(LSS_reals(380) * Press_mult, "##0.0") & Press_chars
    ' 22/06/09 P_875_desup_OUT.text = Format(LSS_reals(381) * Press_mult, "##0.0") & Press_chars
    ' 22/06/09 P_HP_reg_IN.text = Format(LSS_reals(382) * Press_mult, "##0.0") & Press_chars
    ' 22/06/09 P_IP_reg_IN.text = Format(LSS_reals(383) * Press_mult, "##0.0") & Press_chars
    ' 22/06/09 P_35.text = Format(LSS_reals(384) * Press_mult, "##0.0") & Press_chars
    ' 22/06/09 P_Atom.text = Format(LSS_reals(385) * Press_mult, "##0.0") & Press_chars
    
    P_main(1).text = Format(LSS_reals(346) * Press_mult, "##0.0") & Press_chars
    P_main(2).text = Format(LSS_reals(347) * Press_mult, "##0.0") & Press_chars
    
    F_main(1).text = Format(LSS_reals(348) * Flow_mult, "###0.0") & Flow_chars
    F_main(2).text = Format(LSS_reals(349) * Flow_mult, "###0.0") & Flow_chars
    
    F_MF(1).text = Format(LSS_reals(350) * Flow_mult, "###0.0") & Flow_chars
    F_MF(2).text = Format(LSS_reals(351) * Flow_mult, "###0.0") & Flow_chars
    
    F_MF_recirc(1).text = Format(LSS_reals(352) * Flow_mult, "###0.0") & Flow_chars
    F_MF_recirc(2).text = Format(LSS_reals(353) * Flow_mult, "###0.0") & Flow_chars
    
    F_MF_aux(1).text = Format(LSS_reals(354) * Flow_mult, "###0.0") & Flow_chars
    F_MF_aux(2).text = Format(LSS_reals(355) * Flow_mult, "###0.0") & Flow_chars
    
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult, "###0.0") & Flow_chars
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult, "###0.0") & Flow_chars
    
    F_aux_feed(1).text = Format(LSS_reals(360) * Flow_mult, "###0.0") & Flow_chars
    F_aux_feed(2).text = Format(LSS_reals(361) * Flow_mult, "###0.0") & Flow_chars
  End If
  
Else

    ' 17/8/00
    PLO_total_feed_pump(1).text = Format(LSS_reals(486) * Press_mult, "##0") & Press_chars
    PLO_total_feed_pump(2).text = Format(LSS_reals(487) * Press_mult, "##0") & Press_chars


    Z_DFT.text = Format(LSS_reals(451), "###0.0") & " in"
    
    P_aux_feed_out.text = Format(LSS_reals(362) * Press_mult, "##0") & Press_chars
    
' 09/06/09     Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0") & Press_chars
' 09/06/09     Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0") & Press_chars
    
    P_MF.text = Format(LSS_reals(342) * Press_mult, "##0") & Press_chars
    P_AF.text = Format(LSS_reals(343) * Press_mult, "##0") & Press_chars
    
  If (Advanced_mode) Then
  
    ' 10/06/09 P_150.text = Format(LSS_reals(380) * Press_mult, "##0") & Press_chars
    ' 10/06/09 P_875_desup_OUT.text = Format(LSS_reals(381) * Press_mult, "##0") & Press_chars
    ' 10/06/09 P_HP_reg_IN.text = Format(LSS_reals(382) * Press_mult, "##0") & Press_chars
    ' 10/06/09 P_IP_reg_IN.text = Format(LSS_reals(383) * Press_mult, "##0") & Press_chars
    ' 10/06/09 P_35.text = Format(LSS_reals(384) * Press_mult, "##0") & Press_chars
    ' 10/06/09 P_Atom.text = Format(LSS_reals(385) * Press_mult, "##0") & Press_chars
    
    P_main(1).text = Format(LSS_reals(346) * Press_mult, "##0") & Press_chars
    P_main(2).text = Format(LSS_reals(347) * Press_mult, "##0") & Press_chars
    
    F_main(1).text = Format(LSS_reals(348) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_main(2).text = Format(LSS_reals(349) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MF(1).text = Format(LSS_reals(350) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_MF(2).text = Format(LSS_reals(351) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MF_recirc(1).text = Format(LSS_reals(352) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_MF_recirc(2).text = Format(LSS_reals(353) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MF_aux(1).text = Format(LSS_reals(354) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_MF_aux(2).text = Format(LSS_reals(355) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_aux_feed(1).text = Format(LSS_reals(360) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_aux_feed(2).text = Format(LSS_reals(361) * Flow_mult / 3600#, "###0.0") & " lb/s"
  End If
  
End If

'  DFT Tank  ( -35 inches to +35 inches )
DFT_level = (LSS_reals(451) + 35#) * 100# / 70#
If DFT_level > 99 Then
  valve(44).LineWidth = 100
ElseIf DFT_level < 1 Then
  valve(44).LineWidth = 0
Else
  valve(44).LineWidth = DFT_level
End If


' === 21/11/00  Feed Pump re-circ valves===
   Call update_vvalve_LSS(Me, 85, 85, -1)
   Call update_vvalve_LSS(Me, 86, 86, -1)

' OPEN re-circ valves
' ==== 24/11/00 ====
'If (D_OUTPUT(145)) Then
If vv_status_LSS(85) = 3 Then


  OPEN_VV(1).BackColor = &HFF00&            ' GREEN
  
  ' 5/5/00
  'VV_feed_recirc_1(0).Visible = False
  'VV_feed_recirc_1(3).Visible = True
    valve(54).OpenBtmRight = True
    valve(54).OpenTopLeft = True
    
Else
  OPEN_VV(1).BackColor = &HC0C0C0           ' GREY
End If


' ==== 24/11/00 ====
'If (D_OUTPUT(493)) Then
If vv_status_LSS(86) = 3 Then

  OPEN_VV(2).BackColor = &HFF00&            ' GREEN
  
  ' 5/5/00
 ' VV_feed_recirc_2(0).Visible = False
 ' VV_feed_recirc_2(3).Visible = True
    valve(55).OpenBtmRight = True
    valve(55).OpenTopLeft = True
 
Else
  OPEN_VV(2).BackColor = &HC0C0C0           ' GREY
End If

' CLOSE re-circ valves
' ==== 24/11/00 ====
'If (D_OUTPUT(149)) Then
If vv_status_LSS(85) = 0 Then

  CLOSE_VV(1).BackColor = &H8080FF           ' PINK
  
  ' 5/5/00
 ' VV_feed_recirc_1(0).Visible = True
 ' VV_feed_recirc_1(3).Visible = False
    valve(54).OpenBtmRight = False
    valve(54).OpenTopLeft = False
 
Else
  CLOSE_VV(1).BackColor = &HC0C0C0           ' GREY
End If


' ==== 24/11/00 ====
'If (D_OUTPUT(497)) Then
If vv_status_LSS(86) = 0 Then

  CLOSE_VV(2).BackColor = &H8080FF           ' PINK
  
  ' 5/5/00
 ' VV_feed_recirc_2(0).Visible = True
 ' VV_feed_recirc_2(3).Visible = False
    valve(55).OpenBtmRight = False
    valve(55).OpenTopLeft = False

Else
  CLOSE_VV(2).BackColor = &HC0C0C0           ' GREY
End If


' Auto re-circ valves
If (D_OUTPUT(147)) Then
  AUTO_VV(1).BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_VV(1).BackColor = &HC0C0C0           ' GREY
End If
If (D_OUTPUT(495)) Then
  AUTO_VV(2).BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_VV(2).BackColor = &HC0C0C0           ' GREY
End If


' Note: update_pump_LSS(syxxx,index,nPump)
Call update_pump_LSS(Me, 0, 9)

' update_vvalve_LSS(syxxx,nValve,index1,index2)


' 26/8/00  Valve(5) is a POT

'For i = 0 To 7
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i

' === 10/06/09 ===
'For i = 0 To 4
' === 19/06/09 ===
'For i = 0 To 3
For i = 70 To 73
   'Call update_vvalve_LSS(Me, 70 + i, i, -1)
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

' === 10/06/09 ===
'For i = 6 To 7
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i


' === 10/06/09 ===
'For i = 9 To 11
'For i = 10 To 10
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
   Call update_vvalve_LSS(Me, 80, 80, -1)

' === 10/06/09 ===
'For i = 13 To 19
For i = 83 To 87
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

' === 10/06/09 ===
'   Call update_vvalve_LSS(Me, 70 + 22, 22, -1)
   
' === 10/06/09 ===
'For i = 24 To 27
For i = 95 To 97
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

For i = 99 To 101
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

' === 10/06/09 ===
'For i = 34 To 38
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i

' === 10/06/09 ===
'For i = 41 To 42
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i

' === 10/06/09 ===
'For i = 45 To 46
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
   
'  6/9/00   Point Hotel Services Valve
' === 10/06/09 ===
'   Call update_vvalve_LSS(Me, 126, 56, -1)




End Sub

Private Sub TLO_main_feed_Click(Index As Integer)


' 7/9/00
Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_HI = 83
Else
    nAlarm_HI = 84
End If

Alarm_set_point.Show

End Sub

Private Sub Tripped_lbl_Click(Index As Integer)
Dim nlog As Integer


' 17/8/00
If PTT Then
    nlog = 124 + Index
    status_flag = False
    
   ' === 22/11/02 ===
   'Call set_STEAM_logicals(nlog, status_flag)
    Call setSTEAMlogicals(nlog, status_flag)
    
End If

End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
  
 ' 7/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        
    '    IP            Cascade         HP
    
   ' === 10/06/02 Point 14.502 ===
   'If Index = 4 Or Index = 9 Or Index = 24 Then
   
   ' === 23/09/02 Point 14.502 MUST BE CLICKABLE FROM ALL STATIONS ===
   ' If Index = 4 Or Index = 9 Or Index = 24 Or Index = 27 Or Index = 29 Then
   
   ' === 18/06/09 These are AUX STEAM vales and don't exist here BUT...should be 74,79,94
    'If Index = 4 Or Index = 9 Or Index = 24 Then
    If Index = 74 Or Index = 79 Or Index = 94 Then
    
        Exit Sub
        
    End If
    
End If
  
  
' 1/12/99
If Not PTT Then    ' NON-Clickable Valves...


   ' === 21/03/02 Point 14.502 ===
   
   ' === 10/06/02 PUT BACK ORIGINAL LINE... ===
   ' If Index = 13 Or Index = 14 Or Index = 30 Or Index = 31 Then
   ' If Index = 13 Or Index = 14 Or Index = 30 Or Index = 31 Or Index = 27 Or Index = 29 Then
    
    ' === 18/06/09 ADDED 70...===
    'If Index = 13 Or Index = 14 Or Index = 30 Or Index = 31 Then
    If Index = 83 Or Index = 84 Or Index = 100 Or Index = 101 Then
    
        Exit Sub
    End If
End If


' 26/8/00  Valve is a pot
If Index = 5 Then
        Exit Sub
End If



'==== 19/06/09 from PORT BOILER valve_click ==========
If Index = 477 Or Index = 482 Or Index = 478 Then
  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
End If

' 22/06/09
'If Index = 472 Or Index = 476 Or Index = 481 Then
If Index = 472 Or Index = 476 Or Index = 481 Or Index = 479 Then
        'skip safety valves (can't click these !)
      Exit Sub
'Else
'        Call valve_pcdi_LSS(Index)
End If
'==== 19/06/09 from STBD BOILER valve_click ==========
' 22/06/09
'If Index = 447 Or Index = 448 Or Index = 452 Then
If Index = 447 Or Index = 448 Or Index = 452 Or Index = 449 Then
  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
End If
If Index = 442 Or Index = 446 Or Index = 451 Then
      ' Do NOT click safety valves
Else
            
          Call valve_pcdi_LSS(Index)
End If
'=====================================================

    ' === 18/06/09 subtracted 70...===
    ' Call valve_pcdi_LSS(Index + 70)
    ' 19/06/09 REMOVED...
    '  Call valve_pcdi_LSS(Index)




End Sub

Private Sub VScroll1_Change()
Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  HP_bypass

' 1/12/99    Call set_VSCROLL(24, VScroll1.Value)

' 1/12/99
AIvalue = VScroll1.Value / 4
Call setAI(6, AIvalue)

End Sub

Private Sub VScroll2_Change()


End Sub

Private Sub VScroll3_Change()

' 6/11/99  P_875_desup_OUT
Call set_VSCROLL(30, VScroll3.Value)

End Sub

Private Sub VScroll4_Change()

' 6/11/99  P_HP_reg_IN
Call set_VSCROLL(31, VScroll4.Value)


End Sub

Private Sub VScroll5_Change()

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  Dump_35_bypass
Call set_VSCROLL(23, VScroll5.Value)

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

Private Sub VScroll6_Change()

' 8/11/99  P_iP_reg_IN
Call set_VSCROLL(32, VScroll6.Value)

End Sub

Private Sub Z_DFT_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 81
nAlarm_LO = 76
Alarm_set_point.Show

End Sub
Sub update_the_slider(Slider As Object, Pot_Value As Integer)

Dim S_max As Integer
Dim S_min As Integer

' TESTING...
S_max = Slider.max
S_min = Slider.Min

' === 19/02/03 ===
'VScroll1.Value = VB_AI(6) * 4
If Pot_Value >= Slider.max Then
    Slider.Value = Slider.max
ElseIf Pot_Value <= Slider.Min Then
    Slider.Value = Slider.Min
Else
    Slider.Value = Pot_Value
End If


End Sub
