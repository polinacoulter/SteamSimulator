VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form SSTG 
   Caption         =   "SSTG"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   12750
   Icon            =   "sstg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   12750
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7995
      Left            =   0
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   14102
      Begin VB.TextBox TEXSTEAM 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10620
         TabIndex        =   114
         Top             =   360
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox PEXSTEAM 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10620
         TabIndex        =   113
         Top             =   60
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.VScrollBar VScroll_Local 
         Height          =   1155
         Index           =   2
         Left            =   8100
         Max             =   255
         TabIndex        =   109
         Top             =   480
         Width           =   135
      End
      Begin VB.VScrollBar VScroll_Local 
         Height          =   1155
         Index           =   1
         Left            =   2700
         Max             =   255
         TabIndex        =   107
         Top             =   480
         Width           =   135
      End
      Begin VB.TextBox VAC_SSTG 
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
         Index           =   1
         Left            =   1920
         TabIndex        =   102
         Top             =   4860
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox VAC_SSTG 
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
         Index           =   2
         Left            =   7380
         TabIndex        =   101
         Top             =   4800
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000C0&
         Height          =   585
         Left            =   4140
         MultiLine       =   -1  'True
         TabIndex        =   100
         Text            =   "sstg.frx":030A
         Top             =   7140
         Width           =   3450
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
         Index           =   1
         Left            =   12000
         TabIndex        =   99
         Text            =   "*"
         Top             =   4560
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
         Left            =   6540
         TabIndex        =   98
         Text            =   "*"
         Top             =   4500
         Width           =   210
      End
      Begin VB.TextBox Psea_SSTG2 
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
         Left            =   6660
         MousePointer    =   1  'Arrow
         TabIndex        =   97
         Top             =   5940
         Width           =   615
      End
      Begin VB.TextBox TD2 
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
         Left            =   8400
         MousePointer    =   1  'Arrow
         TabIndex        =   96
         Text            =   " "
         Top             =   1980
         Width           =   855
      End
      Begin VB.CommandButton Reset_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESET"
         Height          =   375
         Index           =   2
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   7020
         Width           =   675
      End
      Begin VB.CommandButton Reset_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESET"
         Height          =   375
         Index           =   1
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   7020
         Width           =   675
      End
      Begin VB.CommandButton Trip_cmd 
         BackColor       =   &H008080FF&
         Caption         =   "TRIP"
         Height          =   375
         Index           =   2
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   7020
         Width           =   675
      End
      Begin VB.CommandButton Trip_cmd 
         BackColor       =   &H008080FF&
         Caption         =   "TRIP"
         Height          =   375
         Index           =   1
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   7020
         Width           =   675
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
         Left            =   8160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   6420
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
         Left            =   9120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4800
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
         Left            =   8640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   4800
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
         Height          =   195
         Index           =   0
         Left            =   3600
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   6480
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
         Left            =   3720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4800
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   4800
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox PCONDVAC 
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
         Left            =   660
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Top             =   7200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TLUB3 
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
         Left            =   5820
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Top             =   4320
         Width           =   675
      End
      Begin VB.TextBox plub3 
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
         Left            =   5820
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Top             =   4620
         Width           =   675
      End
      Begin VB.TextBox Zsump3 
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
         Left            =   1260
         MousePointer    =   1  'Arrow
         TabIndex        =   62
         Top             =   4500
         Width           =   800
      End
      Begin VB.TextBox Psea_TA 
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
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Top             =   5940
         Width           =   615
      End
      Begin VB.TextBox psea 
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
         Height          =   285
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   60
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox P_SSTG_stm 
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
         Index           =   2
         Left            =   6480
         TabIndex        =   59
         Top             =   540
         Width           =   800
      End
      Begin VB.TextBox TD3 
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
         MousePointer    =   1  'Arrow
         TabIndex        =   58
         Text            =   " "
         Top             =   1980
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1155
         Index           =   2
         Left            =   7920
         Max             =   1000
         TabIndex        =   57
         Top             =   480
         Width           =   135
      End
      Begin VB.TextBox PID_status 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox PID_status 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3540
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         Top             =   180
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Auto_cmd 
         BackColor       =   &H0000FF00&
         Caption         =   "A"
         Height          =   255
         Index           =   2
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1260
         Width           =   315
      End
      Begin VB.CommandButton Manual_cmd 
         BackColor       =   &H008080FF&
         Caption         =   "M"
         Height          =   255
         Index           =   2
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1020
         Width           =   315
      End
      Begin VB.CommandButton Remote_cmd 
         BackColor       =   &H0000FFFF&
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   780
         Width           =   315
      End
      Begin VB.CommandButton Auto_cmd 
         BackColor       =   &H0000FF00&
         Caption         =   "A"
         Height          =   255
         Index           =   1
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1260
         Width           =   315
      End
      Begin VB.CommandButton Manual_cmd 
         BackColor       =   &H008080FF&
         Caption         =   "M"
         Height          =   255
         Index           =   1
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1020
         Width           =   315
      End
      Begin VB.CommandButton Remote_cmd 
         BackColor       =   &H0000FFFF&
         Caption         =   "R"
         Height          =   255
         Index           =   1
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   780
         Width           =   315
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   60
         Top             =   2280
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SW PUMP"
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   5550
         Width           =   645
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1155
         Index           =   1
         Left            =   2520
         Max             =   1000
         TabIndex        =   12
         Top             =   480
         Width           =   135
      End
      Begin VB.TextBox P_SSTG_stm 
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
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   540
         Width           =   800
      End
      Begin VB.TextBox PLUB2 
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
         Left            =   11280
         TabIndex        =   10
         Top             =   4620
         Width           =   675
      End
      Begin VB.TextBox Tair2 
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
         Left            =   9720
         TabIndex        =   9
         Top             =   1920
         Width           =   800
      End
      Begin VB.TextBox Text5 
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
         Left            =   3840
         TabIndex        =   8
         Text            =   "TEMP"
         Top             =   480
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.TextBox TLUB2 
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
         Left            =   11280
         TabIndex        =   7
         Top             =   4320
         Width           =   675
      End
      Begin VB.TextBox Tair3 
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
         Left            =   4260
         TabIndex        =   6
         Top             =   1890
         Width           =   800
      End
      Begin VB.TextBox Zsump2 
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
         Left            =   6780
         TabIndex        =   5
         Top             =   4440
         Width           =   800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Port Boiler"
         Height          =   495
         Left            =   11100
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stbd. Boiler"
         Height          =   495
         Left            =   9300
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1200
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   12
         Left            =   10095
         TabIndex        =   2
         Top             =   4020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
      End
      Begin PumpPrj.Pump SSTG_1_LO_pump 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Colour3         =   49152
         PumpOpen        =   -1  'True
         State           =   4
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin PumpPrj.Pump SSTG_2_LO_pump 
         Height          =   375
         Left            =   9060
         TabIndex        =   4
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Colour3         =   49152
         PumpOpen        =   -1  'True
         State           =   4
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   0
         Left            =   840
         TabIndex        =   14
         Top             =   1140
         Width           =   495
         _ExtentX        =   873
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
         Index           =   1
         Left            =   1800
         TabIndex        =   15
         Top             =   1140
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   1140
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   3
         Left            =   1140
         TabIndex        =   17
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
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
         Index           =   4
         Left            =   1500
         TabIndex        =   18
         Top             =   6060
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   5
         Left            =   11370
         TabIndex        =   19
         Top             =   6600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   6
         Left            =   6300
         TabIndex        =   20
         Top             =   1140
         Width           =   495
         _ExtentX        =   873
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
         Index           =   7
         Left            =   6780
         TabIndex        =   21
         Top             =   1140
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   8
         Left            =   7260
         TabIndex        =   22
         Top             =   1140
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   9
         Left            =   2760
         TabIndex        =   23
         Top             =   4560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   10
         Left            =   8220
         TabIndex        =   24
         Top             =   4560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   11
         Left            =   4665
         TabIndex        =   25
         Top             =   4020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   13
         Left            =   6600
         TabIndex        =   26
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump Havenpump1 
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   69
         Top             =   4500
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   16711680
      End
      Begin PumpPrj.Pump Havenpump1 
         Height          =   315
         Index           =   1
         Left            =   9000
         TabIndex        =   73
         Top             =   4500
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   16711680
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   14
         Left            =   4160
         TabIndex        =   80
         Top             =   4020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   15
         Left            =   9645
         TabIndex        =   81
         Top             =   4020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   16
         Left            =   5220
         TabIndex        =   82
         Top             =   4920
         Width           =   495
         _ExtentX        =   873
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
         Index           =   17
         Left            =   5640
         TabIndex        =   83
         Top             =   5340
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   18
         Left            =   10680
         TabIndex        =   84
         Top             =   4920
         Width           =   495
         _ExtentX        =   873
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
         Index           =   19
         Left            =   11100
         TabIndex        =   85
         Top             =   5340
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   20
         Left            =   420
         TabIndex        =   86
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   21
         Left            =   5880
         TabIndex        =   87
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   22
         Left            =   10020
         TabIndex        =   88
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   23
         Left            =   10620
         TabIndex        =   89
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label L_label 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         Height          =   195
         Index           =   2
         Left            =   8100
         TabIndex        =   111
         Top             =   240
         Width           =   135
      End
      Begin VB.Label L_label 
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   110
         Top             =   240
         Width           =   135
      End
      Begin VB.Label R_label 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         Height          =   195
         Index           =   2
         Left            =   7920
         TabIndex        =   108
         Top             =   240
         Width           =   135
      End
      Begin VB.Label R_label 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   106
         Top             =   240
         Width           =   135
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   8520
         X2              =   8520
         Y1              =   2520
         Y2              =   2640
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3060
         X2              =   3060
         Y1              =   2520
         Y2              =   2580
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6900
         X2              =   8520
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   1440
         X2              =   3060
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GLAND SEAL STEAM"
         Height          =   735
         Left            =   6240
         TabIndex        =   105
         Top             =   2180
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reg"
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
         Left            =   6600
         TabIndex        =   104
         Top             =   1800
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   1800
         Picture         =   "sstg.frx":0358
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reg"
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
         Left            =   1200
         TabIndex        =   103
         Top             =   1800
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line15 
         Visible         =   0   'False
         X1              =   7860
         X2              =   8280
         Y1              =   4860
         Y2              =   4320
      End
      Begin VB.Line Line9 
         Visible         =   0   'False
         X1              =   2400
         X2              =   2820
         Y1              =   4860
         Y2              =   4320
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   1920
         X2              =   1920
         Y1              =   960
         Y2              =   2400
      End
      Begin VB.Label MAN_SP 
         BackStyle       =   0  'Transparent
         Caption         =   "xxx %"
         Height          =   195
         Index           =   2
         Left            =   8760
         TabIndex        =   94
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label MAN_SP 
         BackStyle       =   0  'Transparent
         Caption         =   "xxx %"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   93
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label GL3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "xxx %"
         Height          =   195
         Left            =   1860
         TabIndex        =   92
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   8800
         Picture         =   "sstg.frx":0662
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   255
      End
      Begin VB.Shape PRE_lub_shape 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   1
         Left            =   8780
         Shape           =   3  'Circle
         Top             =   3300
         Width           =   315
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   3330
         Picture         =   "sstg.frx":096C
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   255
      End
      Begin VB.Shape PRE_lub_shape 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   0
         Left            =   3300
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   315
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BULKHEAD STOP"
         Height          =   405
         Left            =   10560
         TabIndex        =   91
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BULKHEAD STOP"
         Height          =   405
         Left            =   9540
         TabIndex        =   90
         Top             =   780
         Width           =   915
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   9420
         X2              =   11160
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   2760
         Picture         =   "sstg.frx":0C76
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   7080
         Picture         =   "sstg.frx":0F80
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   11220
         X2              =   10920
         Y1              =   6420
         Y2              =   6420
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   10920
         X2              =   10920
         Y1              =   6420
         Y2              =   6840
      End
      Begin VB.Label PRE_lub_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRE LO PUMP"
         Height          =   435
         Index           =   1
         Left            =   8880
         TabIndex        =   75
         Top             =   5100
         Width           =   645
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "PRE LO PUMP"
         Height          =   675
         Left            =   2880
         TabIndex        =   74
         Top             =   3060
         Width           =   465
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   1080
         Picture         =   "sstg.frx":128A
         Stretch         =   -1  'True
         Top             =   3060
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OLE OLE10 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   375
         Left            =   11580
         OleObjectBlob   =   "sstg.frx":1594
         SizeMode        =   1  'Stretch
         TabIndex        =   46
         Top             =   2940
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   2730
         Picture         =   "sstg.frx":35AC
         Stretch         =   -1  'True
         Top             =   5010
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   8190
         Picture         =   "sstg.frx":38B6
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   11070
         Picture         =   "sstg.frx":3BC0
         Stretch         =   -1  'True
         Top             =   6750
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   1140
         Picture         =   "sstg.frx":3ECA
         Stretch         =   -1  'True
         Top             =   6180
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   1650
         Picture         =   "sstg.frx":41D4
         Stretch         =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   7050
         Picture         =   "sstg.frx":44DE
         Stretch         =   -1  'True
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   6600
         Picture         =   "sstg.frx":47E8
         Stretch         =   -1  'True
         Top             =   3030
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line75 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         X1              =   6150
         X2              =   8520
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   5
         Left            =   6300
         Picture         =   "sstg.frx":4AF2
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   6
         Left            =   7260
         Picture         =   "sstg.frx":4DFC
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   570
         Picture         =   "sstg.frx":5106
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   6030
         Picture         =   "sstg.frx":5410
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   10380
         Picture         =   "sstg.frx":571A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   3
         Left            =   840
         Picture         =   "sstg.frx":5A24
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   255
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   7380
         X2              =   7380
         Y1              =   960
         Y2              =   2400
      End
      Begin VB.Line Line38 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   4950
         X2              =   5250
         Y1              =   3810
         Y2              =   3810
      End
      Begin VB.Line Line32 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   4170
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   330
         X2              =   10800
         Y1              =   6300
         Y2              =   6300
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   4770
         X2              =   11460
         Y1              =   6870
         Y2              =   6870
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   4770
         X2              =   4770
         Y1              =   3900
         Y2              =   6120
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   4770
         X2              =   4770
         Y1              =   6390
         Y2              =   6840
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   2880
         Y2              =   6300
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   2580
         Y2              =   6180
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   5190
         X2              =   5760
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   5130
         X2              =   5340
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   5760
         X2              =   5760
         Y1              =   6390
         Y2              =   6840
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   540
         X2              =   1920
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   960
         X2              =   960
         Y1              =   960
         Y2              =   1950
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   540
         X2              =   540
         Y1              =   120
         Y2              =   960
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   2850
         X2              =   2850
         Y1              =   3810
         Y2              =   5220
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   2010
         X2              =   2760
         Y1              =   3720
         Y2              =   3480
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00FFFFFF&
         X1              =   1440
         X2              =   3000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   1020
         X2              =   2040
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line28 
         BorderStyle     =   2  'Dash
         X1              =   4530
         X2              =   4530
         Y1              =   2220
         Y2              =   2610
      End
      Begin VB.Line Line30 
         BorderStyle     =   2  'Dash
         X1              =   5250
         X2              =   5820
         Y1              =   3780
         Y2              =   4620
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillStyle       =   4  'Upward Diagonal
         Height          =   255
         Left            =   2940
         Top             =   4200
         Width           =   285
      End
      Begin VB.Line Line31 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3090
         X2              =   3480
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line33 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   5220
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line35 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   1650
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Line Line36 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   3480
         Y1              =   4320
         Y2              =   2790
      End
      Begin VB.Line Line37 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   5250
         X2              =   5250
         Y1              =   3810
         Y2              =   3180
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   10800
         X2              =   10800
         Y1              =   2880
         Y2              =   6300
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   10230
         X2              =   10230
         Y1              =   6390
         Y2              =   6840
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   11220
         X2              =   11220
         Y1              =   2580
         Y2              =   6420
      End
      Begin VB.Line Line47 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   10650
         X2              =   11220
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line48 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   10590
         X2              =   10800
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   6000
         X2              =   7380
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6420
         X2              =   6420
         Y1              =   960
         Y2              =   1890
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   990
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   8310
         X2              =   8310
         Y1              =   3810
         Y2              =   5220
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   7470
         X2              =   8220
         Y1              =   3720
         Y2              =   3480
      End
      Begin VB.Line Line57 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6510
         X2              =   7500
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line59 
         BorderStyle     =   2  'Dash
         X1              =   10740
         X2              =   11340
         Y1              =   3780
         Y2              =   4620
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         FillStyle       =   4  'Upward Diagonal
         Height          =   255
         Left            =   8400
         Top             =   4170
         Width           =   285
      End
      Begin VB.Line Line61 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8550
         X2              =   8940
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line62 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8940
         X2              =   9630
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line63 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8940
         X2              =   10680
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line65 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8940
         X2              =   7110
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Line Line66 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8940
         X2              =   8940
         Y1              =   4290
         Y2              =   2790
      End
      Begin VB.Line Line67 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   10710
         X2              =   10710
         Y1              =   3810
         Y2              =   3180
      End
      Begin VB.Line Line68 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   10410
         X2              =   10710
         Y1              =   3810
         Y2              =   3810
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00FFFFFF&
         X1              =   1740
         X2              =   1740
         Y1              =   2700
         Y2              =   2520
      End
      Begin VB.Line Line69 
         BorderColor     =   &H00FFFFFF&
         X1              =   3000
         X2              =   3000
         Y1              =   2700
         Y2              =   2520
      End
      Begin VB.Line Line70 
         BorderColor     =   &H00FFFFFF&
         X1              =   7200
         X2              =   7200
         Y1              =   2700
         Y2              =   2520
      End
      Begin VB.Line Line71 
         BorderColor     =   &H00FFFFFF&
         X1              =   8460
         X2              =   8460
         Y1              =   2700
         Y2              =   2520
      End
      Begin VB.Line Line72 
         BorderColor     =   &H00FFFFFF&
         X1              =   6900
         X2              =   8460
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line73 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         X1              =   480
         X2              =   3090
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line76 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         X1              =   1620
         X2              =   1620
         Y1              =   3150
         Y2              =   2850
      End
      Begin VB.Line Line77 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         X1              =   3060
         X2              =   3060
         Y1              =   3180
         Y2              =   2880
      End
      Begin VB.Line Line78 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         X1              =   7020
         X2              =   7020
         Y1              =   3150
         Y2              =   2850
      End
      Begin VB.Line Line79 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         X1              =   8490
         X2              =   8490
         Y1              =   3150
         Y2              =   2850
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OVERBOARD"
         Height          =   375
         Left            =   10440
         TabIndex        =   45
         Top             =   7050
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "FROM SW SYSTEM"
         Height          =   225
         Left            =   120
         TabIndex        =   44
         Top             =   6480
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALTERNATOR"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4050
         TabIndex        =   43
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ALTERNATOR"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9540
         TabIndex        =   42
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GLAND SEAL STEAM"
         Height          =   735
         Left            =   840
         TabIndex        =   41
         Top             =   2180
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DRAIN"
         Height          =   285
         Left            =   330
         TabIndex        =   40
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "DRAIN"
         Height          =   285
         Left            =   5760
         TabIndex        =   39
         Top             =   1260
         Width           =   555
      End
      Begin VB.Line Line5 
         BorderStyle     =   2  'Dash
         X1              =   10020
         X2              =   10020
         Y1              =   2160
         Y2              =   2550
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SSTG # 1"
         Height          =   315
         Left            =   3300
         TabIndex        =   38
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SSTG # 2"
         Height          =   315
         Left            =   8790
         TabIndex        =   37
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TO CONDENSER"
         Height          =   195
         Left            =   2040
         TabIndex        =   36
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "TO CONDENSER"
         Height          =   195
         Left            =   7440
         TabIndex        =   35
         Top             =   5340
         Width           =   1335
      End
      Begin VB.Label PRE_lub_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRE LO PUMP"
         Height          =   435
         Index           =   0
         Left            =   3480
         TabIndex        =   34
         Top             =   5100
         Width           =   645
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "PRE LO PUMP"
         Height          =   675
         Left            =   8400
         TabIndex        =   33
         Top             =   3150
         Width           =   465
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   9750
         X2              =   9750
         Y1              =   3930
         Y2              =   6300
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   4290
         X2              =   4290
         Y1              =   3930
         Y2              =   6300
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   11760
         X2              =   11760
         Y1              =   5310
         Y2              =   7380
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "LO SUMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1860
         TabIndex        =   32
         Top             =   4110
         Width           =   765
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "LO SUMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7290
         TabIndex        =   31
         Top             =   4110
         Width           =   765
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "CASING DRAIN"
         Height          =   465
         Left            =   840
         TabIndex        =   30
         Top             =   3870
         Width           =   645
      End
      Begin VB.Label Label18 
         Caption         =   "CASING DRAIN"
         Height          =   465
         Left            =   6300
         TabIndex        =   29
         Top             =   3870
         Width           =   645
      End
      Begin VB.Label Label23 
         Caption         =   "LEVEL"
         Height          =   225
         Left            =   1260
         TabIndex        =   28
         Top             =   4860
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "LEVEL"
         Height          =   225
         Left            =   6840
         TabIndex        =   27
         Top             =   4770
         Width           =   495
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         X1              =   10230
         X2              =   10230
         Y1              =   3900
         Y2              =   6120
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   540
         X2              =   10500
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   10500
         X2              =   10500
         Y1              =   120
         Y2              =   1440
      End
      Begin VB.OLE OLE8 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   2565
         Left            =   6960
         OleObjectBlob   =   "sstg.frx":5D2E
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\sstgvis.vsd"
         TabIndex        =   47
         Top             =   1980
         Width           =   3765
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   2565
         Left            =   1500
         OleObjectBlob   =   "sstg.frx":A746
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\sstgvis.vsd"
         TabIndex        =   48
         Top             =   1980
         Width           =   3765
      End
      Begin VB.Label GL2 
         Alignment       =   2  'Center
         Caption         =   "xxx %"
         Height          =   195
         Left            =   7320
         TabIndex        =   95
         Top             =   1500
         Width           =   615
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1200
         Top             =   1680
         Width           =   555
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   1440
         X2              =   1440
         Y1              =   960
         Y2              =   2520
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   6600
         Top             =   1680
         Width           =   555
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   6900
         X2              =   6900
         Y1              =   960
         Y2              =   2520
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   540
      TabIndex        =   112
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "SSTG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer
Dim rma_text(2) As String
Dim set_point(2) As Single

Private Sub Auto_cmd_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 16/8/00
' S Green want Remote Auto to be UN-AVAILABLE on the real simulator
If Not PTT Then
    If rma_status_LSS(i) = 2 Then Exit Sub
End If


' Select AUTO
'  Call setPCDI(RMAIO(3,nRMA)      nRMA=1 for SSTG # 1 RMA
'  Call setPCDI(RMAIO(3,nRMA)      nRMA=2 for SSTG # 2 RMA
   Call setPCDI(RMAIO_LSS(3, Index))

End Sub

Private Sub Command1_Click()

Salt_Water.Show

End Sub

Private Sub Command2_Click()

Port_Boiler.Show

End Sub

Private Sub Command3_Click()

Stbd_Boiler.Show

End Sub

Private Sub Form_Load()

' 23/11/99
On Error Resume Next
Call JYAScreenInit(Me, id_SSTG)

' === 15/11/02 ===
update_VBAI(40) = True
update_VBAI(41) = True

' === 18/11/02 ===
update_VSCROLL(33) = True
update_VSCROLL(34) = True

' === 30/09/02 17.008 ===
If Not I_AM_INSTRUCTOR And Not PTT Then
    ' I MUST be an ACT Terminal (Full Mission) so
    ONpump(0).Visible = False
    ONpump(1).Visible = False
    OFFpump(0).Visible = False
    OFFpump(1).Visible = False
    PRE_lub_label(0).Visible = False
    PRE_lub_label(1).Visible = False
Else
    ONpump(0).Visible = True
    ONpump(1).Visible = True
    OFFpump(0).Visible = True
    OFFpump(1).Visible = True
    PRE_lub_label(0).Visible = True
    PRE_lub_label(1).Visible = True
End If
' =======================

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

' 23/11/99
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_SSTG)
    HavenStretch1.Visible = True
End If


End Sub

Private Sub Manual_cmd_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' Select MANUAL
'  Call setPCDI(RMAIO(2,nRMA)      nRMA=1 for SSTG # 1 RMA
'  Call setPCDI(RMAIO(2,nRMA)      nRMA=2 for SSTG # 2 RMA
   Call setPCDI(RMAIO_LSS(2, Index))

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 20))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 21))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

' 18/11/99
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 20))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 21))
End If

End Sub


Private Sub P_SSTG_stm_Click(Index As Integer)

If Index = 1 Then

    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 356
    Alarm_set_point.Show
    
ElseIf Index = 2 Then
  ' 7/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 389
    Alarm_set_point.Show

End If


End Sub

Private Sub PLUB2_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 386
    Alarm_set_point.Show

End Sub

Private Sub plub3_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 353
    Alarm_set_point.Show

End Sub

Private Sub psea_Click()

' 10/7/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 98
Alarm_set_point.Show

End Sub

Private Sub Psea_SSTG2_Click()

  ' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 394
    Alarm_set_point.Show


End Sub

Private Sub Psea_TA_Click()

  ' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 361
    Alarm_set_point.Show

End Sub

Private Sub Remote_cmd_Click(Index As Integer)


set_dummy_focus = True

' 16/1/98
Control_PC_check
If This_Control_status = False Then Exit Sub


' 6/9/00
If Not I_AM_INSTRUCTOR And Not PTT Then
    ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If



' 18/11/99 Removed If statement
'''If In_a_group Then

' Select REMOTE
'  Call setPCDI(RMAIO(1,nRMA)     nRMA=1 for SSTG # 1 RMA
'  Call setPCDI(RMAIO(1,nRMA)     nRMA=2 for SSTG # 2 RMA
   Call setPCDI(RMAIO_LSS(1, Index))
   
   
'''Else
'''            response = MsgBox("OFF LINE Simulation. Only REMOTE Control available", vbOK, "Information")
'''End If

End Sub

Private Sub Reset_cmd_Click(Index As Integer)

' 19/11/99   Operator Trip RESET
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

' 6/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


If Index = 1 Then
   Call setPCDI(705)
ElseIf Index = 2 Then
   Call setPCDI(707)
End If

End Sub

Private Sub Tair2_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 397
    Alarm_set_point.Show

End Sub

Private Sub Tair3_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 364
    Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim green_colour As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

'  SSTG_1_manual_remote  ( Vscroll_value(33) )
'  SSTG_2_manual_remote  ( Vscroll_value(34) )

' === 18/11/02 ===
'        equivalence(STEAM_integers(46),Vscroll_value(1))
'        integer*2                      Vscroll_value(50)
' === 18/11/02 ==============
'equivalence(STEAM_integers(46),Vscroll_value(1))
If update_VSCROLL(33) Then
   update_VSCROLL(33) = False
   
   ' ===26/03/03 ===
   'VScroll1(1).Value = VB_steam_integers(78)
   Call update_the_slider(VScroll1(1), VB_steam_integers(78))
   
End If
If update_VSCROLL(34) Then
   update_VSCROLL(34) = False
   
   ' ===26/03/03 ===
   'VScroll1(2).Value = VB_steam_integers(79)
   Call update_the_slider(VScroll1(2), VB_steam_integers(79))
   
End If

' === 13/11/02 =======
If update_VBAI(40) Then
     update_VBAI(40) = False
     
    ' === 18/11/02 ===
    ' VScroll_Local(1).Value = VB_AI(40) * 4
    
    ' === 26/03/03 ===
    ' VScroll_Local(1).Value = 255 - VB_AI(40)
    Call update_the_slider(VScroll_Local(1), 255 - VB_AI(40))
      
End If
If update_VBAI(41) Then
     update_VBAI(41) = False
   ' === 18/11/02 ===
   ' VScroll_Local(2).Value = VB_AI(41) * 4
   
    ' === 26/03/03 ===
    ' VScroll_Local(2).Value = 255 - VB_AI(41)
    Call update_the_slider(VScroll_Local(2), 255 - VB_AI(41))

End If



'c   15/9/00  SSTG Vacuum inHg
'         PC_REALS(802) = VAC_SSTG(1)
'         PC_REALS(803) = VAC_SSTG(2)
VAC_SSTG(1).text = Format(LSS_reals(802), "###0.0") & " inHg"
VAC_SSTG(2).text = Format(LSS_reals(803), "###0.0") & " inHg"
Call paint_elec_alarm(9, VAC_SSTG(1))
Call paint_elec_alarm(41, VAC_SSTG(2))



'c 21 / 11 / 99
'        PC_REALS(394) = Vscroll_value(33) * 0.1

' === 15/11/02 ===
'MAN_SP(1).Caption = Format(LSS_reals(394), "###0.0") & " %"


' 7/7/00
'    if( GL(2).gt.1.000 ) then
'      PC_REALS(297) = 100#
'    Else
'      PC_REALS(297) = GL(2) * 100#
'    End If
'
'c   7/7/00  SSTG #2  Steam Pressure
'    PC_REALS(298) = P_SSTG_stm(2) / (14.5 * 0.2057)
''
'
'c   7/7/00 SSTG # 2  Manual Set-point
'    PC_REALS(299) = Vscroll_value(34) * 0.1

' === 15/11/02 ===
'MAN_SP(2).Caption = Format(LSS_reals(299), "###0.0") & " %"

' === 18/11/02 ===
'MAN_SP(1).Caption = Format(LSS_reals(394), "##0") & " %"
'MAN_SP(2).Caption = Format(LSS_reals(299), "##0") & " %"




'cc          ETATRIPLP     ( SSTG # 1)
'             LSS_LOGICALS(26) = DO(1409)
'cc          ETATRIPLP_2   ( SSTG # 2)
'             LSS_LOGICALS(27) = DO(1413)

For i = 1 To 2
     If LSS_LOGICALS(i + 25) Then
       Trip_cmd(i).BackColor = &H8080FF      ' LIGHT RED
     Else
       Trip_cmd(i).BackColor = &HC0C0C0      ' GREY
     End If
Next i


' === 18/11/99 ===

' 23/11/99
'        PC_REALS(393) = P_SSTG_stm(1) / (14.5 * 0.2057)

'TSW_LO_cooler.text = Format(LSS_reals(214) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char



TLUB3.text = Format(LSS_reals(514) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

' 7/7/00    SSTG # 2
TLUB2.text = Format(LSS_reals(707) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

TD3.text = Format(LSS_reals(47), "##0") & " rpm"

' 7/7/00    SSTG # 2
TD2.text = Format(LSS_reals(46), "##0") & " rpm"

'           === 06/01/03  ====
'           PC_REALS(867) = PEXSTEAM / (14.5 * 0.2057)
'           PC_REALS(868) = TEXSTEAM
TEXSTEAM.text = Format(LSS_reals(868) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


If SI_units Then

    ' === 06/01/03  ====
    PEXSTEAM.text = Format(LSS_reals(867) * Press_mult, "##0.0") & Press_chars    ' Bar
    
    P_SSTG_stm(1).text = Format(LSS_reals(393) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    ' 7/7/00
    P_SSTG_stm(2).text = Format(LSS_reals(298) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    Psea_TA.text = Format(LSS_reals(735) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    PCONDVAC.text = Format(LSS_reals(503) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    psea.text = Format(LSS_reals(504) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    ' 10/7/00
    Psea_SSTG2.text = Format(LSS_reals(734) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    plub3.text = Format(LSS_reals(510) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    ' 7/7/00
    PLUB2.text = Format(LSS_reals(509) * Press_mult, "##0.0") & Press_chars     ' Bar

  '  FSW_LO_cooler.text = Format(LSS_reals(213) * Flow_mult, "###0.000") & Flow_chars
Else

    ' === 06/01/03  ====
    PEXSTEAM.text = Format(LSS_reals(867) * Press_mult, "##0") & Press_chars    ' psi
    
    P_SSTG_stm(1).text = Format(LSS_reals(393) * Press_mult, "###0") & Press_chars     ' psi
    
    ' 7/7/00
    P_SSTG_stm(2).text = Format(LSS_reals(298) * Press_mult, "###0") & Press_chars     ' psi
    
    Psea_TA.text = Format(LSS_reals(735) * Press_mult, "###0") & Press_chars     ' psi
    
    PCONDVAC.text = Format(LSS_reals(503) * Press_mult, "###0") & Press_chars     ' psi
   
    psea.text = Format(LSS_reals(504) * Press_mult, "###0") & Press_chars     ' psi
    
    ' 10/7/00
    Psea_SSTG2.text = Format(LSS_reals(734) * Press_mult, "###0") & Press_chars     ' psi
    
    plub3.text = Format(LSS_reals(510) * Press_mult, "###0") & Press_chars     ' psi
    
    ' 7/7/00
    PLUB2.text = Format(LSS_reals(509) * Press_mult, "###0") & Press_chars     ' psi
  
  '  FSW_LO_cooler.text = Format(LSS_reals(213) * Flow_mult / 3600#, "###0.000") & " lb/s"
End If

' 23/11/99
'Call paint_elec_alarm(6, STEAMP(1))
Call paint_elec_alarm(6, P_SSTG_stm(1))

' 7/7/00  SSTG # 2
Call paint_elec_alarm(39, P_SSTG_stm(2))



'              EALM(14) = TAIR(3).GT.66.0
'              EALM(32) = TAIR(1).GT.66.0
'              EALM(47) = TAIR(2).GT.66.0
'  ( See 5.47 S Green )
' 5/4/00
Tair3.text = Format(LSS_reals(533) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Call paint_elec_alarm(14, Tair3)
' 7/7/00    SSTG # 2     PC_REALS(532) = Tair(2)
Tair2.text = Format(LSS_reals(532) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Call paint_elec_alarm(47, Tair2)



' SSTG # 1
Call paint_elec_alarm(3, plub3)
Call paint_elec_alarm(8, TLUB3)

' SSTG # 2   7/7/00
Call paint_elec_alarm(36, PLUB2)
Call paint_elec_alarm(42, TLUB2)


' 25/8/98
Call paint_elec_alarm(13, Zsump3)
Zsump3.text = Format(LSS_reals(732), "##0") & "%"

' 7/7/00  SSTG # 2
Call paint_elec_alarm(46, Zsump2)
Zsump2.text = Format(LSS_reals(711), "##0") & "%"


' 25/8/98
Call paint_elec_alarm(11, Psea_TA)
' 6/4/00
'Psea_TA.text = Format(LSS_reals(735), "#0.0") & " bar"

' 10/7/00
Call paint_elec_alarm(44, Psea_SSTG2)

' 10/7/00
'Call paint_elec_alarm(98, psea)
Call paint_alarm(psea, 98, 0, 0, 0)



' 23/11/99
' SSTG # 1 Mechanical Pump
If LSS_reals(47) > 8500# Then
    green_colour = 255
    SSTG_1_LO_pump.Colour3 = RGB(0, green_colour, 0)
ElseIf LSS_reals(47) < 2500# Then
    'green_colour = 0
    SSTG_1_LO_pump.Colour3 = &HFFFFFF        ' WHITE
Else
    green_colour = LSS_reals(47) * 255# / 8500#
    SSTG_1_LO_pump.Colour3 = RGB(0, green_colour, 0)
End If


' 7/7/00
' SSTG # 2 Mechanical Pump
If LSS_reals(46) > 8500# Then
    green_colour = 255
    SSTG_2_LO_pump.Colour3 = RGB(0, green_colour, 0)
ElseIf LSS_reals(46) < 2500# Then
    'green_colour = 0
    SSTG_2_LO_pump.Colour3 = &HFFFFFF        ' WHITE
Else
    green_colour = LSS_reals(46) * 255# / 8500#
    SSTG_2_LO_pump.Colour3 = RGB(0, green_colour, 0)
End If


For i = 1 To 2

    Call update_rma_LSS(i, set_point(i), rma_text(i))    '  1 for SSTG # 1 , 2 for SSTG # 2
    PID_status(i).text = rma_text(i)
       
    If rma_status_LSS(i) = 0 Then
    
        ' LOCAL MANUAL
        Remote_cmd(i).BackColor = &HC0C0C0     ' GREY
        Manual_cmd(i).BackColor = &H8080FF     'PINKy
        Auto_cmd(i).BackColor = &HC0C0C0       ' GREY
        
        
        ' === 18/11/02 ===
        If i = 1 Then
            'MAN_SP(1).Caption = Format(LSS_reals(394), "##0") & " %"
             MAN_SP(1).Caption = Format((255 - VB_AI(40)) / 2.55, "##0") & " %"
        Else
            'MAN_SP(2).Caption = Format(LSS_reals(299), "##0") & " %"
             MAN_SP(2).Caption = Format((255 - VB_AI(41)) / 2.55, "##0") & " %"
        End If
        
        
        ' === 15/11/02 ===
        ' MAN_SP(i).Visible = PTT
        ' VScroll1(i).Visible = PTT
        If PTT Then
            If I_AM_INSTRUCTOR Then
                R_label(i).Visible = True
                VScroll1(i).Visible = True        ' REMOTE MANUAL control
                L_label(i).Visible = True
                VScroll_Local(i).Visible = True   ' LOCAL  MANUAL control
                 MAN_SP(i).Visible = True
                 
                 ' === 18/11/02 ===
                VScroll_Local(i).Enabled = True     ' LOCAL MANUAL control
                VScroll1(i).Enabled = False           ' REMOTE MANUAL control
            Else
                R_label(i).Visible = False
                VScroll1(i).Visible = False       ' REMOTE MANUAL control
                L_label(i).Visible = True
                VScroll_Local(i).Visible = True   ' LOCAL  MANUAL control
                 MAN_SP(i).Visible = False
            End If
        Else
                R_label(i).Visible = False
                VScroll1(i).Visible = False        ' REMOTE MANUAL control
                L_label(i).Visible = False
                VScroll_Local(i).Visible = False   ' LOCAL  MANUAL control
            If I_AM_INSTRUCTOR Then
                 MAN_SP(i).Visible = True
            Else
                 MAN_SP(i).Visible = False
            End If
        End If
                   
        ' === 14/11/02 ===
        ' DO NOT FORCE the RMA to REMOTE control
        'If PTT Then
        '    ' FORCE the RMA to REMOTE control
        '    Call setPCDI(RMAIO_LSS(1, i))
        'End If


    ElseIf rma_status_LSS(i) = 1 Then
    
        ' LOCAL AUTO
        Remote_cmd(i).BackColor = &HC0C0C0     ' GREY
        Manual_cmd(i).BackColor = &HC0C0C0     ' GREY
        Auto_cmd(i).BackColor = &HFF00&        'green
        
        
             ' === 15/11/02 ===
               'VScroll1(i).Visible = False
                R_label(i).Visible = False
                VScroll1(i).Visible = False       ' REMOTE MANUAL control
                L_label(i).Visible = False
                VScroll_Local(i).Visible = False   ' LOCAL  MANUAL control
                 MAN_SP(i).Visible = False

        
        ' === 14/11/02 REMOVE===
        'MAN_SP(i).Visible = False
        
        ' === 14/11/02 ===
        ' DO NOT FORCE the RMA to REMOTE control
        'If PTT Then
        '    ' FORCE the RMA to REMOTE control
        '    Call setPCDI(RMAIO_LSS(1, i))
        'End If
        
    ElseIf rma_status_LSS(i) = 2 Then
    
        ' REMOTE MANUAL
        Remote_cmd(i).BackColor = &HFFFF&         'yellow
        Manual_cmd(i).BackColor = &H8080FF        'PINKy
        Auto_cmd(i).BackColor = &HC0C0C0          'GREY
        
        
         ' === 18/11/02 ===
        If i = 1 Then
              MAN_SP(1).Caption = Format(LSS_reals(394), "##0") & " %"
            ' MAN_SP(1).Caption = Format(VB_AI(40), "##0") & " %"
        Else
              MAN_SP(2).Caption = Format(LSS_reals(299), "##0") & " %"
             'MAN_SP(2).Caption = Format(VB_AI(41), "##0") & " %"
        End If
       
        ' === 15/11/02 ===
        'If (PTT) Then
        '        VScroll1(i).Visible = True
        '        MAN_SP(i).Visible = True
        'Else
        '    If I_AM_INSTRUCTOR Then
        '        VScroll1(i).Visible = True
        '          MAN_SP(i).Visible = True
        '    Else
        '        VScroll1(i).Visible = False
        '          MAN_SP(i).Visible = False
        '    End If
        'End If
        
        ' === 15/11/02 ===
        If PTT Then
            If I_AM_INSTRUCTOR Then
                R_label(i).Visible = True
                VScroll1(i).Visible = True         ' REMOTE MANUAL control
                L_label(i).Visible = True
                VScroll_Local(i).Visible = True  ' LOCAL  MANUAL control
                 MAN_SP(i).Visible = True
                 ' === 18/11/02 ===
                VScroll_Local(i).Enabled = False     ' LOCAL MANUAL control
                VScroll1(i).Enabled = True           ' REMOTE MANUAL control

            Else
                R_label(i).Visible = False
                VScroll1(i).Visible = False       ' REMOTE MANUAL control
                L_label(i).Visible = False
                VScroll_Local(i).Visible = False   ' LOCAL  MANUAL control
                 MAN_SP(i).Visible = False
            End If
        Else
                R_label(i).Visible = True
                VScroll1(i).Visible = True       ' REMOTE MANUAL control
                L_label(i).Visible = False
                VScroll_Local(i).Visible = False   ' LOCAL  MANUAL control
            If I_AM_INSTRUCTOR Then
                 MAN_SP(i).Visible = True
            Else
                 MAN_SP(i).Visible = False
            End If
        End If
        
        
        
        
   ElseIf rma_status_LSS(i) = 3 Then
    
        ' REMOTE AUTO
        Remote_cmd(i).BackColor = &HFFFF&           'yellow
        Manual_cmd(i).BackColor = &HC0C0C0          'GREY
        Auto_cmd(i).BackColor = &HFF00&             'green
        
               ' === 15/11/02 ===
                'VScroll1(i).Visible = False
                'MAN_SP(i).Visible = False
                R_label(i).Visible = False
                VScroll1(i).Visible = False       ' REMOTE MANUAL control
                L_label(i).Visible = False
                VScroll_Local(i).Visible = False   ' LOCAL  MANUAL control
                 MAN_SP(i).Visible = False

        
    End If
    
Next i



' 18/11/99
   Call update_vvalve_LSS(Me, 385, 0, -1)
' 19/11/99
   Call update_vvalve_LSS(Me, 387, 2, -1)
   Call update_vvalve_LSS(Me, 388, 3, -1)
   Call update_vvalve_LSS(Me, 389, 4, -1)
   Call update_vvalve_LSS(Me, 354, 5, -1)
   Call update_vvalve_LSS(Me, 391, 6, -1)
   Call update_vvalve_LSS(Me, 392, 7, -1)

   Call update_vvalve_LSS(Me, 394, 9, -1)
   Call update_vvalve_LSS(Me, 395, 10, -1)
   Call update_vvalve_LSS(Me, 356, 11, -1)
   Call update_vvalve_LSS(Me, 337, 12, -1)
   Call update_vvalve_LSS(Me, 398, 13, -1)
   
   Call update_vvalve_LSS(Me, 355, 14, -1)
   Call update_vvalve_LSS(Me, 336, 15, -1)
   Call update_vvalve_LSS(Me, 357, 16, -1)
   Call update_vvalve_LSS(Me, 358, 17, -1)
   
   Call update_vvalve_LSS(Me, 338, 18, -1)
   Call update_vvalve_LSS(Me, 339, 19, -1)
   Call update_vvalve_LSS(Me, 390, 20, -1)
   Call update_vvalve_LSS(Me, 396, 21, -1)
   Call update_vvalve_LSS(Me, 397, 22, -1)
   Call update_vvalve_LSS(Me, 399, 23, -1)
   
   
' SSTG # 1 Throttling Valve
' PC_REALS(392) = GL(3)
GL3.Caption = Format(LSS_reals(392), "##0") & " %"
Call update_green_valve(392, Valve(1))


' SSTG # 2 Throttling Valve
' PC_REALS(297) = GL(2)
GL2.Caption = Format(LSS_reals(297), "##0") & " %"
Call update_green_valve(297, Valve(8))


' Note: update_pumps(syxxx,index,nPump)
Call update_pump_LSS(Me, 0, 20)      ' SSTG # 1 Pre-Lub

' 21/11/99
'Call update_pump_LSS(Me, 1, 20)     ' SSTG # 2 Pre-Lub
Call update_pump_LSS(Me, 1, 21)      ' SSTG # 2 Pre-Lub

' 21/11/99
If Havenpump1(0).PumpOn Then
    PRE_lub_shape(0).FillColor = &HFF00&   ' GREEN
Else
    PRE_lub_shape(0).FillColor = &HFFFFFF     ' WHITE
End If
    
If Havenpump1(1).PumpOn Then
    PRE_lub_shape(1).FillColor = &HFF00&   ' GREEN
Else
    PRE_lub_shape(1).FillColor = &HFFFFFF     ' WHITE
End If

End Sub

Private Sub TLUB2_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 392
    Alarm_set_point.Show

End Sub

Private Sub TLUB3_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 358
    Alarm_set_point.Show

End Sub

Private Sub Trip_cmd_Click(Index As Integer)

' 19/11/99   Operator Trip
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   


' 6/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

If Index = 1 Then
   Call setPCDI(704)
ElseIf Index = 2 Then
   Call setPCDI(706)
End If

End Sub

Private Sub VAC_SSTG_Click(Index As Integer)

' 10/7/00
    Unload Alarm_set_point   ' Do this first
If Index = 1 Then
    nAlarm_LO = 359
Else
    nAlarm_LO = 391
End If

    Alarm_set_point.Show

End Sub

Private Sub valve_Click(Index As Integer)


Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 18/11/99   SSTG Valves
If Index = 1 Or Index = 8 Then
    Exit Sub
ElseIf Index = 0 Then
    Call valve_pcdi_LSS(385)        ' 19/11/99
ElseIf Index = 2 Then
    Call valve_pcdi_LSS(387)        ' 19/11/99
ElseIf Index = 3 Then
    Call valve_pcdi_LSS(388)        ' 19/11/99
ElseIf Index = 4 Then
    Call valve_pcdi_LSS(389)        ' 19/11/99
    
ElseIf Index = 5 Then
    Call valve_pcdi_LSS(354)        ' Over-Board V/V

ElseIf Index = 6 Then
    Call valve_pcdi_LSS(391)        ' 19/11/99
ElseIf Index = 7 Then
    Call valve_pcdi_LSS(392)        ' 19/11/99
ElseIf Index = 9 Then
    Call valve_pcdi_LSS(394)        ' 19/11/99
ElseIf Index = 10 Then
    Call valve_pcdi_LSS(395)        ' 19/11/99
ElseIf Index = 11 Then
    Call valve_pcdi_LSS(356)        ' 19/11/99
ElseIf Index = 12 Then
    Call valve_pcdi_LSS(337)        ' 19/11/99
ElseIf Index = 13 Then
    Call valve_pcdi_LSS(398)        ' 19/11/99
    
ElseIf Index = 14 Then
    Call valve_pcdi_LSS(355)        ' 19/11/99
ElseIf Index = 15 Then
    Call valve_pcdi_LSS(336)        ' 19/11/99
    
ElseIf Index = 16 Then
    Call valve_pcdi_LSS(357)
    
ElseIf Index = 17 Then
    Call valve_pcdi_LSS(358)
    
ElseIf Index = 18 Then
    Call valve_pcdi_LSS(338)
    
ElseIf Index = 19 Then
    Call valve_pcdi_LSS(339)
    
ElseIf Index = 20 Then
    Call valve_pcdi_LSS(390)
ElseIf Index = 21 Then
    Call valve_pcdi_LSS(396)
ElseIf Index = 22 Then
    Call valve_pcdi_LSS(397)
ElseIf Index = 23 Then
    Call valve_pcdi_LSS(399)

    
    
End If


End Sub

Private Sub VScroll_Local_Change(Index As Integer)

Dim AIvalue As Integer

' === 15/11/02 ===

' === 18/11/02 ===
'AIvalue = VScroll_Local(Index).Value / 4
AIvalue = 255 - VScroll_Local(Index).Value

Call setAI(39 + Index, AIvalue)

End Sub

Private Sub VScroll1_Change(Index As Integer)

' 18/11/99   SSTG_1_manual_remote  ( Vscroll_value(33) )
' 18/11/99   SSTG_2_manual_remote  ( Vscroll_value(33) )
Call set_VSCROLL(32 + Index, VScroll1(Index).Value)

End Sub
Sub update_green_valve(nReals As Integer, obj As Object)
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
Private Sub paint_elec_alarm(nalarm As Integer, obj As Object)
Dim alarmno As Integer

' 19/8/98 SUB Copied from sy201_lss.frm
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

Private Sub Zsump2_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 396
    Alarm_set_point.Show

End Sub

Private Sub Zsump3_Click()

' 10/7/00
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 363
    Alarm_set_point.Show

End Sub
Sub update_the_slider(Slider As Object, Pot_Value As Integer)

Dim S_max As Integer
Dim S_min As Integer

' === 26/03/03 (ADDED) ===

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


