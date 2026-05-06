VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Condensate 
   Caption         =   "Condensate"
   ClientHeight    =   10845
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   16650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10845
   ScaleWidth      =   16650
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   10635
      Left            =   0
      Top             =   0
      Width           =   15795
      _ExtentX        =   27861
      _ExtentY        =   18759
      Begin VB.CommandButton Command6 
         Caption         =   "SSTG"
         Height          =   465
         Left            =   8700
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   11340
         TabIndex        =   197
         Text            =   "L"
         Top             =   4710
         Width           =   255
      End
      Begin VB.TextBox Text2 
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
         Index           =   1
         Left            =   3630
         TabIndex        =   196
         Text            =   "L"
         Top             =   3270
         Width           =   255
      End
      Begin VB.TextBox Text2 
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
         Height          =   255
         Index           =   0
         Left            =   3630
         TabIndex        =   195
         Text            =   "L"
         Top             =   2550
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         Height          =   315
         Left            =   8040
         TabIndex        =   194
         Top             =   2310
         Width           =   285
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Contaminated Steam Evaporator"
         Height          =   735
         Left            =   14460
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   9240
         Width           =   1095
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
         Left            =   12840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   9720
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Index           =   4
         Left            =   13380
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   9720
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   1335
         Left            =   7200
         Max             =   1000
         TabIndex        =   178
         Top             =   660
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
         Left            =   1245
         TabIndex        =   177
         Text            =   "P"
         Top             =   1560
         Width           =   195
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Distiller"
         Height          =   675
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   8940
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Main Condenser"
         Height          =   675
         Left            =   13860
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   4020
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
         Left            =   13740
         TabIndex        =   174
         Text            =   "20 %"
         Top             =   8100
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
         Left            =   2040
         TabIndex        =   166
         Text            =   "xx F"
         Top             =   6780
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
         Left            =   4680
         TabIndex        =   165
         Text            =   "xx F"
         Top             =   3300
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
         Left            =   4620
         TabIndex        =   164
         Text            =   "xx F"
         Top             =   6600
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
         Left            =   6000
         TabIndex        =   163
         Text            =   "xx F"
         Top             =   7080
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
         Left            =   7080
         TabIndex        =   162
         Text            =   "xx F"
         Top             =   7380
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
         Left            =   8220
         TabIndex        =   161
         Text            =   "xx F"
         Top             =   7380
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
         Left            =   8760
         TabIndex        =   160
         Text            =   "xx F"
         Top             =   9600
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
         Left            =   12660
         TabIndex        =   159
         Text            =   "xx F"
         Top             =   5640
         Width           =   855
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   1335
         Left            =   1800
         Max             =   1000
         TabIndex        =   158
         Top             =   2160
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
         Left            =   3000
         TabIndex        =   156
         Text            =   "xx.x kg/s"
         Top             =   7560
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
         Left            =   4230
         TabIndex        =   155
         Text            =   "xx.x kg/s"
         Top             =   4470
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
         Left            =   4440
         TabIndex        =   154
         Text            =   "xx.x kg/s"
         Top             =   1800
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
         Left            =   8160
         TabIndex        =   153
         Text            =   "xx.x kg/s"
         Top             =   6960
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
         Left            =   10260
         TabIndex        =   152
         Text            =   "xx.x kg/s"
         Top             =   5160
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
         Left            =   9300
         TabIndex        =   151
         Text            =   "xx.xkg/s"
         Top             =   6180
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
         Left            =   9000
         TabIndex        =   150
         Text            =   "xx.x kg/s"
         Top             =   4920
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
         Left            =   6600
         TabIndex        =   149
         Text            =   "xx.x kg/s"
         Top             =   3540
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
         Left            =   12420
         TabIndex        =   148
         Text            =   "xx.x kg/s"
         Top             =   6000
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
         Left            =   4440
         TabIndex        =   147
         Text            =   "xx.x kg/s"
         Top             =   8400
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
         Left            =   5760
         TabIndex        =   146
         Text            =   "x.x Bar"
         Top             =   4200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1335
         Left            =   13560
         Max             =   1000
         TabIndex        =   144
         Top             =   5040
         Width           =   135
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   6660
         Max             =   1000
         TabIndex        =   130
         Top             =   4200
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   4140
         Max             =   100
         TabIndex        =   129
         Top             =   1140
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   1620
         Max             =   100
         TabIndex        =   128
         Top             =   1200
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
         Left            =   4020
         TabIndex        =   122
         Text            =   "x.x Bar"
         Top             =   9780
         Width           =   795
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1335
         Left            =   5640
         Max             =   1000
         TabIndex        =   120
         Top             =   2160
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
         Left            =   9540
         TabIndex        =   99
         Text            =   "x.x Bar"
         Top             =   6720
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
         Left            =   2040
         TabIndex        =   98
         Text            =   "20 %"
         Top             =   6480
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
         Left            =   4680
         TabIndex        =   97
         Text            =   "20 %"
         Top             =   3000
         Width           =   795
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Left            =   600
         Max             =   1000
         TabIndex        =   96
         Top             =   6900
         Width           =   135
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1035
         Left            =   1080
         Max             =   1000
         TabIndex        =   95
         Top             =   4440
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
         Left            =   11280
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   8160
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
         Left            =   11400
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   6600
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   9480
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   8880
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
         Left            =   10800
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   8160
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
         Left            =   10920
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   6600
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
         Left            =   2700
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   9480
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
         Left            =   2700
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   8880
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
         Left            =   10320
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   8160
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
         Left            =   2160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   9480
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
         Left            =   2160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   8880
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
         Left            =   10440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   6600
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   14280
         Top             =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Gland Seal"
         Height          =   495
         Left            =   7320
         TabIndex        =   54
         Top             =   5580
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Hotwell_tank 
         Height          =   705
         Left            =   12420
         TabIndex        =   9
         Top             =   5115
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   13
         Left            =   10260
         TabIndex        =   34
         Top             =   5580
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4920
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
         Left            =   12660
         TabIndex        =   1
         Text            =   "20 %"
         Top             =   5340
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
         Left            =   8760
         TabIndex        =   0
         Text            =   "20 %"
         Top             =   9240
         Width           =   855
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   3
         Left            =   10980
         TabIndex        =   3
         Top             =   7500
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   273
         Left            =   11880
         TabIndex        =   4
         Top             =   7440
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
         Index           =   272
         Left            =   10560
         TabIndex        =   5
         Top             =   7200
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
         Left            =   10920
         TabIndex        =   6
         Top             =   6060
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   271
         Left            =   11880
         TabIndex        =   7
         Top             =   6000
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
         Index           =   270
         Left            =   10560
         TabIndex        =   8
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   0
         Left            =   4380
         TabIndex        =   17
         Top             =   8870
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
         Left            =   4380
         TabIndex        =   18
         Top             =   9410
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   80
         Left            =   4860
         TabIndex        =   19
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
         Height          =   315
         Index           =   2
         Left            =   8160
         TabIndex        =   24
         Top             =   4500
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   3
         Left            =   8700
         TabIndex        =   25
         Top             =   5520
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   8
         Left            =   9060
         TabIndex        =   26
         Top             =   5520
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   9
         Left            =   8820
         TabIndex        =   27
         Top             =   4500
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   120
         Left            =   660
         TabIndex        =   40
         Top             =   1020
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   111
         Left            =   2340
         TabIndex        =   41
         Top             =   195
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
         Index           =   112
         Left            =   1620
         TabIndex        =   42
         Top             =   195
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
         Index           =   113
         Left            =   2040
         TabIndex        =   43
         Top             =   720
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
         Index           =   98
         Left            =   5100
         TabIndex        =   44
         Top             =   600
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   107
         Left            =   4140
         TabIndex        =   45
         Top             =   195
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
         Index           =   108
         Left            =   4890
         TabIndex        =   46
         Top             =   195
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
         Index           =   109
         Left            =   4560
         TabIndex        =   47
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   1275
         Index           =   302
         Left            =   4440
         TabIndex        =   48
         Top             =   2400
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   11
         Left            =   3960
         TabIndex        =   52
         Top             =   8580
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   14
         Left            =   3960
         TabIndex        =   53
         Top             =   9120
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   77
         Left            =   780
         TabIndex        =   60
         Top             =   1980
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
         Index           =   81
         Left            =   780
         TabIndex        =   61
         Top             =   2640
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
         Index           =   91
         Left            =   1440
         TabIndex        =   62
         Top             =   2400
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
         Height          =   315
         Index           =   22
         Left            =   9840
         TabIndex        =   64
         Top             =   7800
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   274
         Left            =   11715
         TabIndex        =   65
         Top             =   7080
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   275
         Left            =   11340
         TabIndex        =   66
         Top             =   5760
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   15
         Left            =   1500
         TabIndex        =   67
         Top             =   6840
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
         Index           =   16
         Left            =   1500
         TabIndex        =   68
         Top             =   7500
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
         Index           =   17
         Left            =   840
         TabIndex        =   69
         Top             =   7140
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   10
         Left            =   4740
         TabIndex        =   71
         Top             =   8700
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   12
         Left            =   4740
         TabIndex        =   72
         Top             =   9240
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
         Left            =   5460
         TabIndex        =   2
         Top             =   8940
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   18
         Left            =   1920
         TabIndex        =   75
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   20
         Left            =   1920
         TabIndex        =   76
         Top             =   5160
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
         Index           =   21
         Left            =   1260
         TabIndex        =   77
         Top             =   4800
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
         Height          =   675
         Left            =   600
         TabIndex        =   78
         Top             =   5760
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1191
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   110
         Left            =   2400
         TabIndex        =   100
         Top             =   600
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   19
         Left            =   3780
         TabIndex        =   118
         Top             =   5700
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   24
         Left            =   2160
         TabIndex        =   119
         Top             =   7380
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   23
         Left            =   6420
         TabIndex        =   123
         Top             =   4620
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   1
         Left            =   6660
         TabIndex        =   124
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
         Index           =   5
         Left            =   7080
         TabIndex        =   126
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   4
         Left            =   7425
         TabIndex        =   125
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
         Left            =   10620
         TabIndex        =   145
         Text            =   "x.x Bar"
         Top             =   9600
         Visible         =   0   'False
         Width           =   795
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 CDT_tank 
         Height          =   915
         Left            =   12780
         TabIndex        =   168
         Top             =   7500
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   583
         Left            =   12420
         TabIndex        =   172
         Top             =   7965
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
         Index           =   4
         Left            =   13140
         TabIndex        =   182
         Top             =   9375
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   428
         Left            =   12840
         TabIndex        =   187
         Top             =   8400
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   429
         Left            =   13680
         TabIndex        =   188
         Top             =   9120
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   487
         Left            =   12810
         TabIndex        =   189
         Top             =   7140
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   5
         Left            =   7620
         Picture         =   "Condensate2.frx":0000
         Stretch         =   -1  'True
         Top             =   8070
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   14
         X1              =   7740
         X2              =   7740
         Y1              =   7770
         Y2              =   9030
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   50
         X1              =   5760
         X2              =   5760
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   3840
         Picture         =   "Condensate2.frx":030A
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   3
         Left            =   1860
         Picture         =   "Condensate2.frx":0614
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   2
         Left            =   3570
         Picture         =   "Condensate2.frx":091E
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   85
         X1              =   3180
         X2              =   3300
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   76
         X1              =   3480
         X2              =   3720
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Image Image19 
         Height          =   255
         Index           =   2
         Left            =   8850
         Picture         =   "Condensate2.frx":0C28
         Stretch         =   -1  'True
         Top             =   2670
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   64
         X1              =   8910
         X2              =   10920
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   61
         X1              =   10920
         X2              =   10920
         Y1              =   3630
         Y2              =   2790
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "From 35 # Dump"
         Height          =   255
         Left            =   7590
         TabIndex        =   193
         Top             =   2670
         Width           =   1425
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   2
         Left            =   3060
         Picture         =   "Condensate2.frx":0F32
         Stretch         =   -1  'True
         Top             =   3750
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   1
         Left            =   4260
         Picture         =   "Condensate2.frx":123C
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   15
         Left            =   12780
         Picture         =   "Condensate2.frx":1546
         Stretch         =   -1  'True
         Top             =   6840
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   9
         Left            =   1500
         Picture         =   "Condensate2.frx":1850
         Stretch         =   -1  'True
         Top             =   6420
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   1
         Left            =   5400
         Picture         =   "Condensate2.frx":1B5A
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Index           =   2
         Left            =   4560
         Picture         =   "Condensate2.frx":1E64
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Index           =   1
         Left            =   2040
         Picture         =   "Condensate2.frx":216E
         Stretch         =   -1  'True
         Top             =   8340
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   4800
         Picture         =   "Condensate2.frx":2478
         Stretch         =   -1  'True
         Top             =   7860
         Width           =   255
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From SSTG # 1"
         Height          =   195
         Index           =   1
         Left            =   8820
         TabIndex        =   192
         Top             =   2460
         Width           =   1710
      End
      Begin VB.Image Image19 
         Height          =   255
         Index           =   1
         Left            =   10620
         Picture         =   "Condensate2.frx":2782
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   375
      End
      Begin VB.Image Image27 
         Height          =   255
         Left            =   10620
         Picture         =   "Condensate2.frx":2A8C
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   84
         X1              =   11040
         X2              =   11040
         Y1              =   2580
         Y2              =   3600
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From SSTG # 2"
         Height          =   195
         Index           =   0
         Left            =   8820
         TabIndex        =   191
         Top             =   2160
         Width           =   1710
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   83
         X1              =   11220
         X2              =   11220
         Y1              =   2280
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   82
         X1              =   10620
         X2              =   11220
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   13
         X1              =   10620
         X2              =   11040
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   12900
         X2              =   12900
         Y1              =   7500
         Y2              =   6900
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "From DWT"
         Height          =   435
         Left            =   12600
         TabIndex        =   190
         Top             =   6480
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "CDT Pump"
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
         Left            =   12900
         TabIndex        =   186
         Top             =   9960
         Width           =   975
      End
      Begin VB.Image Image26 
         Height          =   255
         Left            =   14100
         Picture         =   "Condensate2.frx":2D96
         Stretch         =   -1  'True
         Top             =   9255
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   81
         X1              =   13440
         X2              =   14340
         Y1              =   9390
         Y2              =   9390
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   80
         X1              =   12960
         X2              =   13200
         Y1              =   9540
         Y2              =   9540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   79
         X1              =   12960
         X2              =   12960
         Y1              =   8400
         Y2              =   9540
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   45
         X1              =   1620
         X2              =   1620
         Y1              =   6240
         Y2              =   6900
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   44
         X1              =   2040
         X2              =   2040
         Y1              =   4380
         Y2              =   4560
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   5790
         Picture         =   "Condensate2.frx":30A0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   43
         X1              =   6180
         X2              =   5160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   42
         X1              =   4140
         X2              =   2640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line PY_line 
         Index           =   1
         X1              =   6120
         X2              =   6360
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line PY_line 
         Index           =   0
         X1              =   6120
         X2              =   6120
         Y1              =   4440
         Y2              =   4800
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Caption         =   "De-Aerator Flow Rate"
         Height          =   435
         Left            =   6900
         TabIndex        =   180
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label38 
         Caption         =   "0"
         Height          =   315
         Left            =   7440
         TabIndex        =   179
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   8880
         X2              =   8880
         Y1              =   5280
         Y2              =   5160
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   8880
         X2              =   8880
         Y1              =   5030
         Y2              =   4320
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   8520
         Picture         =   "Condensate2.frx":33AA
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   41
         X1              =   11520
         X2              =   7680
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   9120
         X2              =   9120
         Y1              =   5940
         Y2              =   5280
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   29
         X1              =   8760
         X2              =   8040
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   40
         X1              =   8340
         X2              =   8340
         Y1              =   5390
         Y2              =   5280
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   35
         X1              =   8220
         X2              =   8220
         Y1              =   5460
         Y2              =   5160
      End
      Begin VB.Line Line26 
         BorderStyle     =   3  'Dot
         X1              =   1320
         X2              =   1140
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line19 
         BorderStyle     =   3  'Dot
         X1              =   1320
         X2              =   1320
         Y1              =   1860
         Y2              =   2460
      End
      Begin VB.Image Image25 
         Height          =   255
         Left            =   120
         Picture         =   "Condensate2.frx":36B4
         Stretch         =   -1  'True
         Top             =   7200
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   78
         X1              =   4140
         X2              =   4140
         Y1              =   3660
         Y2              =   3960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   77
         X1              =   240
         X2              =   3150
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   15
         X1              =   8340
         X2              =   8340
         Y1              =   6120
         Y2              =   5580
      End
      Begin VB.Shape Sight_Glass 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   13530
         Shape           =   3  'Circle
         Top             =   7560
         Width           =   135
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To Slop Tank"
         Height          =   495
         Left            =   14280
         TabIndex        =   173
         Top             =   8580
         Width           =   855
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   14
         Left            =   14640
         Picture         =   "Condensate2.frx":39BE
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   75
         X1              =   14760
         X2              =   14760
         Y1              =   7680
         Y2              =   8580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   74
         X1              =   13680
         X2              =   14760
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   73
         X1              =   12240
         X2              =   12840
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   70
         X1              =   12240
         X2              =   12240
         Y1              =   8280
         Y2              =   8760
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sight Glass"
         Height          =   255
         Left            =   13695
         TabIndex        =   171
         Top             =   7680
         Width           =   855
      End
      Begin VB.Shape Shape41 
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   915
         Left            =   13500
         Shape           =   4  'Rounded Rectangle
         Top             =   7500
         Width           =   195
      End
      Begin VB.Image Image24 
         Height          =   255
         Left            =   9780
         Picture         =   "Condensate2.frx":3CC8
         Stretch         =   -1  'True
         Top             =   8640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   72
         X1              =   8580
         X2              =   8580
         Y1              =   9120
         Y2              =   8760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   71
         X1              =   8580
         X2              =   12240
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Image Image23 
         Height          =   255
         Left            =   13800
         Picture         =   "Condensate2.frx":3FD2
         Stretch         =   -1  'True
         Top             =   7200
         Width           =   255
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   14160
         TabIndex        =   170
         Top             =   6960
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   69
         X1              =   13200
         X2              =   13200
         Y1              =   7740
         Y2              =   7320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   68
         X1              =   13200
         X2              =   14100
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contaminated Drain Tank"
         Height          =   435
         Left            =   13020
         TabIndex        =   169
         Top             =   6780
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   67
         X1              =   9120
         X2              =   8400
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Label Label32 
         Caption         =   "Misc returns"
         Height          =   255
         Left            =   8640
         TabIndex        =   167
         Top             =   8400
         Width           =   855
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   13
         Left            =   8280
         Picture         =   "Condensate2.frx":42DC
         Stretch         =   -1  'True
         Top             =   8520
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   63
         X1              =   8400
         X2              =   8400
         Y1              =   8400
         Y2              =   9000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   66
         X1              =   3180
         X2              =   3180
         Y1              =   4080
         Y2              =   1740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   65
         X1              =   240
         X2              =   240
         Y1              =   4080
         Y2              =   8460
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Hotwell"
         Height          =   255
         Left            =   11580
         TabIndex        =   157
         Top             =   5160
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   32
         X1              =   8040
         X2              =   6360
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label COND_recirc_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   143
         Top             =   4620
         Width           =   555
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   9780
         TabIndex        =   137
         Top             =   1050
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape40 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   9450
         Top             =   1440
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   33
         X1              =   900
         X2              =   900
         Y1              =   1560
         Y2              =   2040
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   6240
         Picture         =   "Condensate2.frx":45E6
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   34
         X1              =   8040
         X2              =   8040
         Y1              =   5100
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   46
         X1              =   8760
         X2              =   10620
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line18 
         X1              =   10260
         X2              =   10380
         Y1              =   6900
         Y2              =   6900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   40
         X1              =   240
         X2              =   3900
         Y1              =   8460
         Y2              =   8460
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   12
         Left            =   5880
         Picture         =   "Condensate2.frx":48F0
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   62
         X1              =   6000
         X2              =   6000
         Y1              =   7800
         Y2              =   9000
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   9120
         X2              =   8340
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label28 
         Caption         =   "Air Ejector"
         Height          =   255
         Left            =   8280
         TabIndex        =   82
         Top             =   7875
         Width           =   855
      End
      Begin VB.Image Image20 
         Height          =   360
         Left            =   720
         Picture         =   "Condensate2.frx":4BFA
         Stretch         =   -1  'True
         Top             =   5460
         Width           =   285
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   0
         Left            =   960
         Picture         =   "Condensate2.frx":4F04
         Stretch         =   -1  'True
         Top             =   360
         Width           =   375
      End
      Begin VB.Image Image12 
         Height          =   255
         Index           =   0
         Left            =   2520
         Picture         =   "Condensate2.frx":520E
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image19 
         Height          =   255
         Index           =   0
         Left            =   9960
         Picture         =   "Condensate2.frx":5518
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   30
         X1              =   10800
         X2              =   10800
         Y1              =   3600
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   17
         X1              =   9780
         X2              =   10800
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line9 
         X1              =   5400
         X2              =   5640
         Y1              =   5280
         Y2              =   5160
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   11
         Left            =   3840
         Picture         =   "Condensate2.frx":5822
         Stretch         =   -1  'True
         Top             =   3780
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
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   79
         Top             =   5640
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   29
         X1              =   4020
         X2              =   5520
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line29 
         BorderStyle     =   3  'Dot
         X1              =   3060
         X2              =   3060
         Y1              =   2670
         Y2              =   4980
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   30
         X1              =   1380
         X2              =   1380
         Y1              =   4380
         Y2              =   5640
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3840
         X2              =   1380
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   2040
         X2              =   1380
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line23 
         BorderStyle     =   3  'Dot
         X1              =   3720
         X2              =   3000
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Line Line22 
         BorderStyle     =   3  'Dot
         X1              =   3660
         X2              =   3480
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line21 
         BorderStyle     =   3  'Dot
         X1              =   3480
         X2              =   1860
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line17 
         BorderStyle     =   3  'Dot
         X1              =   3480
         X2              =   3480
         Y1              =   3360
         Y2              =   7320
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   8970
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image Image18 
         Height          =   360
         Left            =   7200
         Picture         =   "Condensate2.frx":5B2C
         Stretch         =   -1  'True
         Top             =   8640
         Width           =   285
      End
      Begin VB.Line Line16 
         BorderStyle     =   3  'Dot
         X1              =   4620
         X2              =   4500
         Y1              =   9180
         Y2              =   9060
      End
      Begin VB.Line Line15 
         BorderStyle     =   3  'Dot
         X1              =   4620
         X2              =   4560
         Y1              =   9180
         Y2              =   9420
      End
      Begin VB.Line Line14 
         BorderStyle     =   3  'Dot
         X1              =   5220
         X2              =   4680
         Y1              =   9180
         Y2              =   9180
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   5280
         TabIndex        =   73
         Top             =   9120
         Width           =   195
      End
      Begin VB.Shape Shape31 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   5220
         Top             =   9060
         Width           =   315
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   3360
         Picture         =   "Condensate2.frx":5E36
         Stretch         =   -1  'True
         Top             =   7860
         Width           =   375
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   10
         Left            =   5400
         Picture         =   "Condensate2.frx":6140
         Stretch         =   -1  'True
         Top             =   8340
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   5520
         X2              =   5520
         Y1              =   7980
         Y2              =   9060
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   960
         X2              =   5520
         Y1              =   7980
         Y2              =   7980
      End
      Begin VB.Label Label7 
         Caption         =   "Make up Feed Water to ADT"
         Height          =   255
         Left            =   3120
         TabIndex        =   70
         Top             =   8040
         Width           =   2295
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1620
         X2              =   960
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   960
         X2              =   960
         Y1              =   6720
         Y2              =   7980
      End
      Begin VB.Line Line13 
         BorderStyle     =   3  'Dot
         X1              =   11820
         X2              =   11460
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Shape Shape19 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   11460
         Shape           =   3  'Circle
         Top             =   5040
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   6360
         X2              =   6720
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   8
         Left            =   8520
         Picture         =   "Condensate2.frx":644A
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   8640
         X2              =   8640
         Y1              =   4320
         Y2              =   3660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   9480
         X2              =   10380
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   24
         X1              =   10380
         X2              =   10380
         Y1              =   7380
         Y2              =   8040
      End
      Begin VB.Line Line8 
         X1              =   5400
         X2              =   5640
         Y1              =   5040
         Y2              =   4920
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
         Left            =   600
         TabIndex        =   63
         Top             =   3420
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   780
         Picture         =   "Condensate2.frx":6754
         Stretch         =   -1  'True
         Top             =   3180
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   38
         X1              =   1560
         X2              =   1560
         Y1              =   1560
         Y2              =   3180
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   39
         X1              =   900
         X2              =   1560
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   1800
         Picture         =   "Condensate2.frx":6A5E
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   36
         X1              =   1920
         X2              =   1920
         Y1              =   1560
         Y2              =   1980
      End
      Begin VB.Label Label37 
         Caption         =   "To Evap."
         Height          =   435
         Left            =   2040
         TabIndex        =   59
         Top             =   1800
         Width           =   435
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   8220
         OleObjectBlob   =   "Condensate2.frx":6D68
         SizeMode        =   1  'Stretch
         TabIndex        =   58
         Top             =   6210
         Width           =   195
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   7950
         OleObjectBlob   =   "Condensate2.frx":8980
         SizeMode        =   1  'Stretch
         TabIndex        =   57
         Top             =   6240
         Width           =   195
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   8640
         OleObjectBlob   =   "Condensate2.frx":A598
         SizeMode        =   1  'Stretch
         TabIndex        =   56
         Top             =   6030
         Width           =   255
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   9000
         OleObjectBlob   =   "Condensate2.frx":C1B0
         SizeMode        =   1  'Stretch
         TabIndex        =   55
         Top             =   6030
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   3
         Left            =   5820
         Picture         =   "Condensate2.frx":DDC8
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Astern Turbine"
         Height          =   495
         Left            =   10860
         TabIndex        =   32
         Top             =   360
         Width           =   555
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   11340
         Picture         =   "Condensate2.frx":E0D2
         Stretch         =   -1  'True
         Top             =   2100
         Width           =   255
      End
      Begin VB.Line Line38 
         BorderStyle     =   3  'Dot
         X1              =   5940
         X2              =   5940
         Y1              =   2940
         Y2              =   5460
      End
      Begin VB.Line Line37 
         BorderStyle     =   3  'Dot
         X1              =   12120
         X2              =   5940
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line36 
         BorderStyle     =   3  'Dot
         X1              =   12120
         X2              =   12120
         Y1              =   2160
         Y2              =   2940
      End
      Begin VB.Line Line10 
         BorderStyle     =   3  'Dot
         X1              =   11460
         X2              =   11460
         Y1              =   6300
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   58
         X1              =   4380
         X2              =   4560
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   57
         X1              =   4140
         X2              =   4560
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "DFT Water level"
         Height          =   615
         Left            =   4680
         TabIndex        =   49
         Top             =   2340
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   23
         X1              =   5400
         X2              =   3960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   25
         X1              =   1440
         X2              =   1440
         Y1              =   1020
         Y2              =   480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   24
         X1              =   2880
         X2              =   2880
         Y1              =   1020
         Y2              =   480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   20
         X1              =   2880
         X2              =   1440
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   21
         X1              =   3960
         X2              =   3960
         Y1              =   960
         Y2              =   480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   22
         X1              =   5400
         X2              =   5400
         Y1              =   960
         Y2              =   480
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   3300
         Picture         =   "Condensate2.frx":E3DC
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "No 2"
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         Top             =   9240
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "No 1"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   8520
         Width           =   375
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   7
         Left            =   12000
         Picture         =   "Condensate2.frx":E6E6
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "L.P. Extraction"
         Height          =   495
         Left            =   12360
         TabIndex        =   33
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   6
         Left            =   7980
         Picture         =   "Condensate2.frx":E9F0
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   6780
         Picture         =   "Condensate2.frx":ECFA
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   9540
         Picture         =   "Condensate2.frx":F004
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   0
         Left            =   7140
         Picture         =   "Condensate2.frx":F30E
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   0
         Left            =   9900
         Picture         =   "Condensate2.frx":F618
         Stretch         =   -1  'True
         Top             =   3720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "L.P Ahead Turbine"
         Height          =   375
         Left            =   11820
         TabIndex        =   31
         Top             =   420
         Width           =   975
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   11460
         X2              =   11460
         Y1              =   1500
         Y2              =   3600
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   12120
         X2              =   12120
         Y1              =   2100
         Y2              =   1320
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   8580
         X2              =   8580
         Y1              =   6540
         Y2              =   6180
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   7980
         X2              =   8580
         Y1              =   6180
         Y2              =   6180
      End
      Begin VB.Label Label4 
         Caption         =   "L.P. Extraction"
         Height          =   495
         Left            =   6600
         TabIndex        =   30
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "150 psi."
         Height          =   255
         Left            =   8400
         TabIndex        =   29
         Top             =   3240
         Width           =   615
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   19
         X1              =   8880
         X2              =   8220
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line20 
         BorderStyle     =   3  'Dot
         X1              =   8700
         X2              =   9480
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   9480
         X2              =   10920
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   9480
         X2              =   9480
         Y1              =   6000
         Y2              =   3840
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   2400
         Picture         =   "Condensate2.frx":F922
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   2760
         Picture         =   "Condensate2.frx":FC2C
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape Shape38 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8835
         Top             =   6900
         Width           =   285
      End
      Begin VB.Shape Shape18 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8295
         Top             =   6900
         Width           =   285
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   8820
         Shape           =   3  'Circle
         Top             =   6660
         Width           =   315
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   8280
         Shape           =   3  'Circle
         Top             =   6660
         Width           =   315
      End
      Begin VB.Shape Shape37 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   7515
         Top             =   6840
         Width           =   225
      End
      Begin VB.Shape Shape36 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   7500
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   60
         X1              =   7740
         X2              =   7740
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   59
         X1              =   7500
         X2              =   7500
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Shape Shape15 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8595
         Top             =   6840
         Width           =   225
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   8580
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   54
         X1              =   8820
         X2              =   8820
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   53
         X1              =   8580
         X2              =   8580
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Shape Shape34 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8055
         Top             =   6840
         Width           =   225
      End
      Begin VB.Shape Shape33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   8040
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   56
         X1              =   8280
         X2              =   8280
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   55
         X1              =   8040
         X2              =   8040
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   6735
         Top             =   6840
         Width           =   225
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   6720
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   52
         X1              =   6960
         X2              =   6960
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   41
         X1              =   6720
         X2              =   6720
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Shape Shape30 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   5760
         Top             =   6840
         Width           =   225
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   5760
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   51
         X1              =   6000
         X2              =   6000
         Y1              =   6840
         Y2              =   7500
      End
      Begin VB.Shape Loop_Seal_shape 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   9255
         Top             =   7860
         Width           =   210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   49
         X1              =   9480
         X2              =   9480
         Y1              =   7680
         Y2              =   8160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   48
         X1              =   9240
         X2              =   9240
         Y1              =   7800
         Y2              =   8160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   47
         X1              =   9120
         X2              =   9120
         Y1              =   7140
         Y2              =   6900
      End
      Begin VB.Shape Shape26 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   11940
         Shape           =   3  'Circle
         Top             =   5400
         Width           =   135
      End
      Begin VB.Shape Shape25 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   10740
         Shape           =   3  'Circle
         Top             =   4740
         Width           =   135
      End
      Begin VB.Shape Shape20 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   10560
         Shape           =   3  'Circle
         Top             =   4500
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   9
         X1              =   9960
         X2              =   9960
         Y1              =   4860
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   6
         X1              =   9960
         X2              =   10800
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   9960
         X2              =   9960
         Y1              =   7080
         Y2              =   5160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   9960
         X2              =   9960
         Y1              =   7680
         Y2              =   7320
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   8760
         X2              =   8760
         Y1              =   6000
         Y2              =   5460
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   6780
         Picture         =   "Condensate2.frx":FF36
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   7500
         Picture         =   "Condensate2.frx":10240
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   14
         X1              =   8040
         X2              =   8040
         Y1              =   6120
         Y2              =   5460
      End
      Begin VB.Shape Shape24 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   8640
         Top             =   5940
         Width           =   255
      End
      Begin VB.Shape Shape23 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   9000
         Top             =   5940
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   8220
         Top             =   6120
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Atmospheric Drain Tank"
         Height          =   255
         Left            =   6000
         TabIndex        =   23
         Top             =   9840
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "875#/35#"
         Height          =   195
         Left            =   5640
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   0
         Left            =   5400
         Picture         =   "Condensate2.frx":1054A
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "IP Bleed "
         Height          =   435
         Left            =   660
         TabIndex        =   20
         Top             =   540
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   7620
         X2              =   7620
         Y1              =   5820
         Y2              =   6540
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   8220
         X2              =   8220
         Y1              =   4320
         Y2              =   5040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   5940
         X2              =   5940
         Y1              =   5460
         Y2              =   6540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   45
         X1              =   10620
         X2              =   10620
         Y1              =   3960
         Y2              =   4560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   44
         X1              =   5640
         X2              =   8520
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   43
         X1              =   4140
         X2              =   5400
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   42
         X1              =   8100
         X2              =   8100
         Y1              =   7800
         Y2              =   9000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   39
         X1              =   3900
         X2              =   4500
         Y1              =   9420
         Y2              =   9420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   38
         X1              =   3900
         X2              =   4500
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   26
         X1              =   3900
         X2              =   3900
         Y1              =   8460
         Y2              =   9420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   25
         X1              =   5160
         X2              =   5460
         Y1              =   9660
         Y2              =   9660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   21
         X1              =   5160
         X2              =   5160
         Y1              =   9000
         Y2              =   9660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   20
         X1              =   4560
         X2              =   5160
         Y1              =   9540
         Y2              =   9540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   19
         X1              =   4560
         X2              =   5160
         Y1              =   9000
         Y2              =   9000
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   10740
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   16
         X1              =   5520
         X2              =   5520
         Y1              =   5220
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   15
         X1              =   6885
         X2              =   6885
         Y1              =   7800
         Y2              =   9000
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   2520
         X2              =   2520
         Y1              =   1560
         Y2              =   2220
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   2880
         X2              =   2880
         Y1              =   1560
         Y2              =   2280
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   13
         X1              =   3420
         X2              =   3420
         Y1              =   1920
         Y2              =   480
      End
      Begin VB.Line R_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   10380
         X2              =   12060
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line R_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   10380
         X2              =   10380
         Y1              =   5520
         Y2              =   6060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   10380
         X2              =   10620
         Y1              =   7500
         Y2              =   7500
      End
      Begin VB.Line Line12 
         BorderStyle     =   3  'Dot
         X1              =   11835
         X2              =   11835
         Y1              =   7740
         Y2              =   5640
      End
      Begin VB.Line Line11 
         BorderStyle     =   3  'Dot
         X1              =   11460
         X2              =   11580
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   10620
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   22
         X1              =   7740
         X2              =   8040
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   6900
         X2              =   6900
         Y1              =   6540
         Y2              =   6300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   5
         X1              =   5520
         X2              =   5760
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   11
         X1              =   6960
         X2              =   7500
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   6000
         X2              =   6720
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Image Image9 
         Height          =   360
         Left            =   3840
         Picture         =   "Condensate2.frx":10854
         Stretch         =   -1  'True
         Top             =   1155
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
         Left            =   3360
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   3960
         X2              =   5520
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Evap"
         Height          =   255
         Left            =   6660
         TabIndex        =   15
         Top             =   5940
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   7
         X1              =   10380
         X2              =   11160
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1290
         Left            =   7380
         Top             =   6540
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   12
         X1              =   3960
         X2              =   3960
         Y1              =   5040
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   23
         X1              =   9120
         X2              =   10380
         Y1              =   7155
         Y2              =   7155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   9480
         X2              =   9960
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   32
         X1              =   10380
         X2              =   11220
         Y1              =   7500
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   33
         X1              =   10380
         X2              =   10380
         Y1              =   6060
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   34
         X1              =   11280
         X2              =   12360
         Y1              =   6300
         Y2              =   6300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   35
         X1              =   11280
         X2              =   12360
         Y1              =   7740
         Y2              =   7740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   36
         X1              =   12360
         X2              =   12360
         Y1              =   7740
         Y2              =   5760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   37
         X1              =   6360
         X2              =   6360
         Y1              =   4560
         Y2              =   6840
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
         Left            =   4800
         TabIndex        =   13
         Top             =   7200
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
         Left            =   10080
         TabIndex        =   12
         Top             =   8400
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         Caption         =   "No 1"
         Height          =   255
         Left            =   10980
         TabIndex        =   11
         Top             =   5700
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "No 2"
         Height          =   255
         Left            =   10980
         TabIndex        =   10
         Top             =   7200
         Width           =   495
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   10200
         Shape           =   3  'Circle
         Top             =   1755
         Width           =   15
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   0
         Left            =   10140
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   15
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   12540
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1515
         Left            =   10440
         Top             =   3600
         Width           =   2595
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1755
         Index           =   2
         Left            =   9660
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   1515
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1980
         Left            =   12240
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   1515
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   3420
         X2              =   900
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   7920
         Top             =   6120
         Width           =   255
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   7980
         Width           =   255
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   7860
         Top             =   6540
         Width           =   1515
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   5640
         Top             =   6540
         Width           =   555
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   6600
         Top             =   6540
         Width           =   555
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   6600
         Top             =   5880
         Width           =   615
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   885
         Left            =   11295
         OleObjectBlob   =   "Condensate2.frx":10B5E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hotwell.vsd"
         TabIndex        =   50
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3345
         Left            =   10740
         OleObjectBlob   =   "Condensate2.frx":18376
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   51
         Top             =   -1260
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
         Left            =   11460
         TabIndex        =   14
         Top             =   3180
         Width           =   1575
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   6900
         X2              =   6900
         Y1              =   6000
         Y2              =   5580
      End
      Begin VB.Shape Shape32 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   8940
         Top             =   930
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   3300
         Top             =   1935
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   5
         Left            =   3300
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   3480
         Shape           =   3  'Circle
         Top             =   1515
         Width           =   735
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   37
         X1              =   900
         X2              =   900
         Y1              =   2940
         Y2              =   3480
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "From Aux Exhaust"
         Height          =   255
         Left            =   9240
         TabIndex        =   81
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label26 
         Caption         =   "DFT Dump Regulator"
         Height          =   435
         Left            =   2460
         TabIndex        =   80
         Top             =   5160
         Width           =   855
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   2040
         X2              =   2040
         Y1              =   5280
         Y2              =   5820
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   1620
         X2              =   1620
         Y1              =   7740
         Y2              =   7980
      End
      Begin VB.Label DWT_dump_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   139
         Top             =   6840
         Width           =   555
      End
      Begin VB.Label DFT_dump_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   140
         Top             =   4680
         Width           =   555
      End
      Begin VB.Label DWT_dump_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1860
         TabIndex        =   138
         Top             =   7080
         Width           =   555
      End
      Begin VB.Label DFT_dump_bypass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1380
         TabIndex        =   141
         Top             =   4500
         Width           =   555
      End
      Begin VB.Label COND_recirc_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7020
         TabIndex        =   142
         Top             =   5280
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   18
         X1              =   5520
         X2              =   5520
         Y1              =   1740
         Y2              =   4980
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   780
         OleObjectBlob   =   "Condensate2.frx":1FD8E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   105
         Top             =   2340
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   780
         OleObjectBlob   =   "Condensate2.frx":27DA6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   108
         Top             =   2340
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   780
         OleObjectBlob   =   "Condensate2.frx":2FBBE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   107
         Top             =   2340
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   780
         OleObjectBlob   =   "Condensate2.frx":383D6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   106
         Top             =   2340
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   26
         X1              =   1680
         X2              =   900
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Auxiliary Exhaust from Turbo feed and Cargo Pumps"
         Height          =   1215
         Left            =   2070
         TabIndex        =   22
         Top             =   2280
         Width           =   1005
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   1980
         OleObjectBlob   =   "Condensate2.frx":403EE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   113
         Top             =   150
         Width           =   405
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   1980
         OleObjectBlob   =   "Condensate2.frx":48406
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   133
         Top             =   120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   1980
         OleObjectBlob   =   "Condensate2.frx":5021E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   132
         Top             =   120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   1980
         OleObjectBlob   =   "Condensate2.frx":58436
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   131
         Top             =   120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   4500
         OleObjectBlob   =   "Condensate2.frx":60E4E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   114
         Top             =   150
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   4500
         OleObjectBlob   =   "Condensate2.frx":68E66
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   136
         Top             =   120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   4500
         OleObjectBlob   =   "Condensate2.frx":70C7E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   135
         Top             =   120
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   4500
         OleObjectBlob   =   "Condensate2.frx":78E96
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   134
         Top             =   120
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
         Index           =   0
         Left            =   1920
         OleObjectBlob   =   "Condensate2.frx":818AE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   101
         Top             =   4860
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
         Left            =   1920
         OleObjectBlob   =   "Condensate2.frx":898C6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   104
         Top             =   4860
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
         Left            =   1920
         OleObjectBlob   =   "Condensate2.frx":916DE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   103
         Top             =   4860
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
         Index           =   1
         Left            =   1920
         OleObjectBlob   =   "Condensate2.frx":996F6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   102
         Top             =   4860
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
         Left            =   1500
         OleObjectBlob   =   "Condensate2.frx":A1F0E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   109
         Top             =   7200
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
         Left            =   1500
         OleObjectBlob   =   "Condensate2.frx":A9F26
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   112
         Top             =   7200
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
         Left            =   1500
         OleObjectBlob   =   "Condensate2.frx":B1D3E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   111
         Top             =   7200
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
         Index           =   2
         Left            =   1500
         OleObjectBlob   =   "Condensate2.frx":BA556
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   110
         Top             =   7200
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   7020
         OleObjectBlob   =   "Condensate2.frx":C256E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   127
         Top             =   4740
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   7020
         OleObjectBlob   =   "Condensate2.frx":CA586
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   117
         Top             =   4740
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   7020
         OleObjectBlob   =   "Condensate2.frx":D259E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   116
         Top             =   4740
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_cond_recirc 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   7020
         OleObjectBlob   =   "Condensate2.frx":DA7B6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   115
         Top             =   4740
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Line Line28 
         BorderStyle     =   3  'Dot
         X1              =   3060
         X2              =   2220
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gland Exhaust Condenser"
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   6930
         TabIndex        =   181
         Top             =   7920
         Width           =   795
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
      Left            =   10500
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
      Left            =   10020
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
      Left            =   9540
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
      Left            =   780
      TabIndex        =   121
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Condensate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Command1_Click()


' 6/4/00  set_dummy_focus = True


' 13/07/09
'Feed_Water.Show
Feed_Water_Alone.Show


End Sub

Private Sub Command12_Click()

Aux_Steam_Alone.ZOrder

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

Private Sub Command5_Click()

' === 16/08/02 ===
Stm_Stm_Generator.ZOrder


End Sub

Private Sub Command6_Click()


' 13/07/09
SSTG.ZOrder

' === 20/11/02 17.011 ===
SSTG.WindowState = 0 ' Normal


End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idCondensate)


' === 13/11/02 ===
update_VBAI(9) = True
update_VBAI(10) = True
update_VBAI(11) = True



' === 13/12/02 19.126 ===



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


' === 16/12/02 19.126 ===
'Valve(13).Visible = True
valve(13).Visible = False
R_line(1).Visible = False
R_line(2).Visible = False
Shape26.Visible = False

' 1/12/99
VScroll1.Visible = PTT
VScroll4.Visible = PTT

HScroll3.Visible = PTT

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


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
    
    
  Case 4
    Call setPCDI(4405) ' CDT pump

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
    
  Case 4
    Call setPCDI(4404) ' CDT pump
    
    
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

' === 13/11/02 ===========
If update_VBAI(9) Then
     update_VBAI(9) = False
     
     ' === 19/02/03 ===
     'VScroll1.Value = VB_AI(9) * 4
      Call update_the_slider(VScroll1, VB_AI(9) * 4)

     
End If
If update_VBAI(10) Then
     update_VBAI(10) = False
     
     'VScroll4.Value = VB_AI(10) * 4
      Call update_the_slider(VScroll4, VB_AI(10) * 4)
     
End If
If update_VBAI(11) Then
     update_VBAI(11) = False
     
     'HScroll3.Value = VB_AI(11) * 4
      Call update_the_slider(HScroll3, VB_AI(11) * 4)
     
End If
' === 13/11/02 END =======


' === 21/08/02 ===
If LSS_LOGICALS(100) Then
    HavenPump1(4).PumpOn = True
    ONpump(4).BackColor = &HFF00&          ' GREEN
    OFFpump(4).BackColor = &HC0C0C0        ' GREY
Else
    HavenPump1(4).PumpOn = False
    ONpump(4).BackColor = &HC0C0C0        ' GREY
    OFFpump(4).BackColor = &HFFFFFF      ' WHITE
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
   
' === 20/08/02 ===
'CDT Outlet Valve to CDT Pump
   Call update_vvalve_LSS(Me, 428, 428, -1)
'CDT Pump Out
   Call update_vvalve_LSS(Me, 429, 429, -1)
   
  
'=== 07/03/07 ADDED  DWT to CDT Tank valve for TEXAS...
   Call update_vvalve_LSS(Me, 487, 487, -1)
 
   

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

