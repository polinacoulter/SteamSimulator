VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Begin VB.Form SwitchBoards 
   Caption         =   "SWITCHBOARDS"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   9930
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7035
      Left            =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12409
      Begin VB.CommandButton Reverse_power 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   4
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   3000
         Width           =   435
      End
      Begin VB.CommandButton Reverse_power 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   3
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   3000
         Width           =   435
      End
      Begin VB.CommandButton Reverse_power 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   2
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   3000
         Width           =   435
      End
      Begin VB.CommandButton Reverse_power 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "TRIP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   1
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   3000
         Width           =   435
      End
      Begin VB.CommandButton Lower 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   3
         Left            =   2400
         TabIndex        =   121
         Top             =   6360
         Width           =   495
      End
      Begin VB.CommandButton Raise 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   3
         Left            =   2880
         TabIndex        =   120
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1020
         MousePointer    =   1  'Arrow
         TabIndex        =   119
         Text            =   "VEx"
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   3
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   118
         Top             =   4260
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   3
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   117
         Top             =   4020
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   3
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   116
         Top             =   3780
         Width           =   675
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   3
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   115
         Top             =   3540
         Width           =   555
      End
      Begin VB.TextBox Text27 
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
         Height          =   250
         Left            =   2880
         MousePointer    =   1  'Arrow
         TabIndex        =   114
         Text            =   " Hz"
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox Text25 
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
         Height          =   250
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   113
         Text            =   " V"
         Top             =   3780
         Width           =   375
      End
      Begin VB.TextBox Text24 
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
         Height          =   250
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   112
         Text            =   " A"
         Top             =   4020
         Width           =   375
      End
      Begin VB.TextBox Text23 
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
         Height          =   250
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   111
         Text            =   " kW"
         Top             =   4260
         Width           =   375
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   3
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   110
         Top             =   4500
         Width           =   675
      End
      Begin VB.TextBox Text17 
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
         Height          =   250
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   109
         Text            =   " PF"
         Top             =   4500
         Width           =   375
      End
      Begin VB.CommandButton AVR_auto 
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
         Height          =   257
         Index           =   3
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton AVR_MAN 
         Caption         =   "MAN"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   3
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   5580
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         MousePointer    =   1  'Arrow
         TabIndex        =   106
         Text            =   "VEx"
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         MousePointer    =   1  'Arrow
         TabIndex        =   105
         Text            =   "VEx"
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Index           =   3
         Left            =   2400
         MousePointer    =   1  'Arrow
         TabIndex        =   104
         Top             =   5220
         Width           =   675
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   3
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   102
         Top             =   3300
         Width           =   555
      End
      Begin VB.TextBox Text6 
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
         Height          =   250
         Left            =   2880
         MousePointer    =   1  'Arrow
         TabIndex        =   101
         Text            =   " rpm"
         Top             =   3300
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   3
         Left            =   3120
         Max             =   255
         TabIndex        =   100
         Top             =   5040
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "S/S OFF"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   4260
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S/S ON"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   4260
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Synchroscope"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7380
         TabIndex        =   97
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton Lower 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   2
         Left            =   5160
         TabIndex        =   96
         Top             =   6360
         Width           =   495
      End
      Begin VB.CommandButton Lower 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   1
         Left            =   3720
         TabIndex        =   95
         Top             =   6360
         Width           =   495
      End
      Begin VB.CommandButton Raise 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   2
         Left            =   5640
         TabIndex        =   94
         Top             =   6360
         Width           =   495
      End
      Begin VB.CommandButton Raise 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   1
         Left            =   4200
         TabIndex        =   93
         Top             =   6360
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   4
         Left            =   1740
         Max             =   255
         TabIndex        =   92
         Top             =   5040
         Width           =   255
      End
      Begin VB.CommandButton AVR_auto 
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
         Height          =   257
         Index           =   2
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   5820
         Width           =   675
      End
      Begin VB.CommandButton AVR_auto 
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
         Height          =   257
         Index           =   1
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton AVR_MAN 
         Caption         =   "MAN"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   2
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   5580
         Width           =   675
      End
      Begin VB.CommandButton AVR_MAN 
         Caption         =   "MAN"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   1
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   5580
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   1
         Left            =   4440
         Max             =   255
         TabIndex        =   87
         Top             =   5040
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   2
         Left            =   5880
         Max             =   255
         TabIndex        =   86
         Top             =   5040
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ELECTRICAL CONTROL PANEL"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7380
         TabIndex        =   85
         Top             =   4980
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Left            =   9000
         TabIndex        =   84
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Left            =   3300
         TabIndex        =   83
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Left            =   3300
         TabIndex        =   82
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Index           =   4
         Left            =   1020
         MousePointer    =   1  'Arrow
         TabIndex        =   81
         Top             =   5220
         Width           =   675
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Index           =   2
         Left            =   5160
         MousePointer    =   1  'Arrow
         TabIndex        =   80
         Top             =   5220
         Width           =   675
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Index           =   1
         Left            =   3720
         MousePointer    =   1  'Arrow
         TabIndex        =   79
         Top             =   5220
         Width           =   675
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   4
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   78
         Top             =   3300
         Width           =   555
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   77
         Top             =   3300
         Width           =   555
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   76
         Top             =   3300
         Width           =   555
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   4
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   74
         Top             =   4500
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   4
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   73
         Top             =   4260
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   4
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   72
         Top             =   4020
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   4
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   71
         Top             =   3780
         Width           =   675
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   4
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   70
         Top             =   3540
         Width           =   555
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Top             =   4500
         Width           =   675
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   68
         Top             =   4500
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   67
         Top             =   4260
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   66
         Top             =   4260
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Top             =   4020
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Top             =   4020
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Top             =   3780
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   62
         Top             =   3780
         Width           =   675
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   2
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Top             =   3540
         Width           =   555
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   3660
         MousePointer    =   1  'Arrow
         TabIndex        =   60
         Top             =   3540
         Width           =   555
      End
      Begin VB.CommandButton AVR_auto 
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
         Height          =   257
         Index           =   4
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton AVR_MAN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MAN"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   4
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5580
         Width           =   615
      End
      Begin VB.TextBox Text28 
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
         Height          =   250
         Left            =   5760
         MousePointer    =   1  'Arrow
         TabIndex        =   46
         Text            =   " PF"
         Top             =   4500
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1320
         Top             =   6420
      End
      Begin VB.TextBox valt4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   6840
         MousePointer    =   1  'Arrow
         TabIndex        =   45
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox aalt4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   6840
         MousePointer    =   1  'Arrow
         TabIndex        =   44
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox walt4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   6840
         MousePointer    =   1  'Arrow
         TabIndex        =   43
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox vesrs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   7920
         MousePointer    =   1  'Arrow
         TabIndex        =   42
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox ashore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   7920
         MousePointer    =   1  'Arrow
         TabIndex        =   41
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox wshore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   7920
         MousePointer    =   1  'Arrow
         TabIndex        =   40
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text4 
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
         Height          =   250
         Left            =   1620
         MousePointer    =   1  'Arrow
         TabIndex        =   39
         Text            =   " V"
         Top             =   3780
         Width           =   375
      End
      Begin VB.TextBox Text8 
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
         Height          =   250
         Left            =   4320
         MousePointer    =   1  'Arrow
         TabIndex        =   38
         Text            =   " V"
         Top             =   3780
         Width           =   375
      End
      Begin VB.TextBox Text11 
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
         Height          =   250
         Left            =   7440
         MousePointer    =   1  'Arrow
         TabIndex        =   37
         Text            =   " V"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox Text12 
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
         Height          =   250
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   36
         Text            =   " V"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox Text9 
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
         Height          =   250
         Left            =   1620
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Text            =   " A"
         Top             =   4020
         Width           =   375
      End
      Begin VB.TextBox Text14 
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
         Height          =   250
         Left            =   4320
         MousePointer    =   1  'Arrow
         TabIndex        =   34
         Text            =   " A"
         Top             =   4020
         Width           =   375
      End
      Begin VB.TextBox Text1 
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
         Height          =   250
         Left            =   1620
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         Text            =   " kW"
         Top             =   4260
         Width           =   375
      End
      Begin VB.TextBox Text16 
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
         Height          =   250
         Left            =   4320
         MousePointer    =   1  'Arrow
         TabIndex        =   32
         Text            =   " kW"
         Top             =   4260
         Width           =   375
      End
      Begin VB.TextBox Text21 
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
         Height          =   250
         Left            =   7440
         MousePointer    =   1  'Arrow
         TabIndex        =   31
         Text            =   " A"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox Text22 
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
         Height          =   250
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   30
         Text            =   " A"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox Text19 
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
         Height          =   250
         Left            =   7440
         MousePointer    =   1  'Arrow
         TabIndex        =   29
         Text            =   " kW"
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox Text20 
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
         Height          =   250
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         Text            =   " kW"
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox Text29 
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
         Height          =   250
         Left            =   4320
         MousePointer    =   1  'Arrow
         TabIndex        =   27
         Text            =   " PF"
         Top             =   4500
         Width           =   375
      End
      Begin VB.TextBox Text31 
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
         Height          =   250
         Left            =   1620
         MousePointer    =   1  'Arrow
         TabIndex        =   26
         Text            =   " PF"
         Top             =   4500
         Width           =   375
      End
      Begin VB.TextBox Text10 
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
         Height          =   250
         Left            =   5760
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         Text            =   " V"
         Top             =   3780
         Width           =   375
      End
      Begin VB.TextBox Text15 
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
         Height          =   250
         Left            =   5760
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Text            =   " A"
         Top             =   4020
         Width           =   375
      End
      Begin VB.TextBox Text18 
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
         Height          =   250
         Left            =   5760
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Text            =   " kW"
         Top             =   4260
         Width           =   375
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5160
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Text            =   "VEx"
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox Text2 
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
         Height          =   287
         Left            =   6480
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Text            =   " V"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text61 
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
         Height          =   287
         Left            =   660
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         Text            =   " PF"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text59 
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
         Height          =   287
         Left            =   1680
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Text            =   " Hz"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text57 
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
         Height          =   287
         Left            =   2700
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Text            =   " V"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text55 
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
         Height          =   287
         Left            =   3720
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Text            =   " A"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text53 
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
         Height          =   287
         Left            =   4740
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Text            =   " kW"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text49 
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
         Height          =   287
         Left            =   7500
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Text            =   " A"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox Text39 
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
         Height          =   287
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Text            =   " kW"
         Top             =   180
         Width           =   375
      End
      Begin VB.TextBox webes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   8880
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox aebes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   7860
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox vebes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   6840
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox wbusb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   5100
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox abusb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   4080
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox vembs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   3060
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox powfact 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   1020
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox hembs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   287
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox Text37 
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
         Height          =   250
         Left            =   4200
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Text            =   " rpm"
         Top             =   3300
         Width           =   495
      End
      Begin VB.TextBox Text33 
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
         Height          =   250
         Left            =   1500
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Text            =   " rpm"
         Top             =   3300
         Width           =   495
      End
      Begin VB.TextBox Text5 
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
         Height          =   250
         Left            =   1500
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Text            =   " Hz"
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox Text32 
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
         Height          =   250
         Left            =   4200
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Text            =   " Hz"
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox Text36 
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
         Height          =   250
         Left            =   5640
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Text            =   " Hz"
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox Text40 
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
         Height          =   250
         Left            =   5640
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Text            =   " rpm"
         Top             =   3300
         Width           =   495
      End
      Begin VB.CommandButton Reverse_power 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   257
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   5520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SWBD #13"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   3120
         MousePointer    =   1  'Arrow
         TabIndex        =   131
         Top             =   1980
         Width           =   795
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SWBD #11"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2520
         MousePointer    =   1  'Arrow
         TabIndex        =   130
         Top             =   600
         Width           =   765
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   8
         Left            =   240
         Picture         =   "SwitchBoards.frx":0000
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   495
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   240
         X2              =   720
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line28 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   390
         X2              =   390
         Y1              =   2460
         Y2              =   1080
      End
      Begin VB.Line Line27 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   520
         X2              =   520
         Y1              =   2460
         Y2              =   1200
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   660
         X2              =   660
         Y1              =   2439
         Y2              =   1380
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "THRUSTER"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   0
         MousePointer    =   1  'Arrow
         TabIndex        =   129
         Top             =   3180
         Width           =   900
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BOW"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   300
         MousePointer    =   1  'Arrow
         TabIndex        =   128
         Top             =   3000
         Width           =   390
      End
      Begin VB.Label Label8 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   440
         TabIndex        =   127
         Top             =   2580
         Width           =   195
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   495
         Left            =   240
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   555
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   7
         Left            =   1800
         Picture         =   "SwitchBoards.frx":030A
         Stretch         =   -1  'True
         Top             =   915
         Width           =   495
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   3420
         Shape           =   3  'Circle
         Top             =   2220
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   3
         Left            =   2460
         Picture         =   "SwitchBoards.frx":0614
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   495
      End
      Begin VB.Line Line24 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   2610
         X2              =   2610
         Y1              =   2412
         Y2              =   1080
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2760
         X2              =   2760
         Y1              =   2395
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   2900
         X2              =   2900
         Y1              =   2439
         Y2              =   1380
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "T/A"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   103
         Top             =   3075
         Width           =   300
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   2100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   7740
         Shape           =   3  'Circle
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   4620
         Shape           =   3  'Circle
         Top             =   2100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   5700
         Shape           =   3  'Circle
         Top             =   1860
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   735
         Left            =   2220
         Top             =   6180
         Width           =   4935
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1275
         Left            =   780
         Top             =   4920
         Width           =   6375
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   3300
         X2              =   3300
         Y1              =   1560
         Y2              =   900
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   3420
         X2              =   3420
         Y1              =   1560
         Y2              =   900
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3540
         X2              =   3540
         Y1              =   1560
         Y2              =   900
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   620
         Index           =   5
         Left            =   8160
         Picture         =   "SwitchBoards.frx":091E
         Stretch         =   -1  'True
         Top             =   2421
         Width           =   615
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   4
         Left            =   1080
         Picture         =   "SwitchBoards.frx":0C28
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   615
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   3
         Left            =   2460
         Picture         =   "SwitchBoards.frx":0F32
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   615
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   2
         Left            =   5160
         Picture         =   "SwitchBoards.frx":123C
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   1
         Left            =   3780
         Picture         =   "SwitchBoards.frx":1546
         Stretch         =   -1  'True
         Top             =   2415
         Width           =   615
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   8880
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   3180
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   5880
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   255
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   499
         Index           =   0
         Left            =   7080
         Picture         =   "SwitchBoards.frx":1850
         Stretch         =   -1  'True
         Top             =   1574
         Width           =   495
      End
      Begin VB.Image HBKR_closed 
         Height          =   480
         Left            =   120
         Picture         =   "SwitchBoards.frx":1B5A
         Top             =   4500
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image HBKR_open 
         Height          =   480
         Left            =   120
         Picture         =   "SwitchBoards.frx":1E64
         Top             =   4080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image GEN_on 
         Height          =   480
         Left            =   120
         Picture         =   "SwitchBoards.frx":216E
         Top             =   6300
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image GEN_off 
         Height          =   480
         Left            =   120
         Picture         =   "SwitchBoards.frx":2478
         Top             =   5820
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image VBKR_open 
         Height          =   480
         Left            =   120
         Picture         =   "SwitchBoards.frx":2782
         Top             =   4980
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image VBKR_closed 
         Height          =   480
         Left            =   660
         Picture         =   "SwitchBoards.frx":2A8C
         Top             =   6360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   4
         Left            =   1080
         Picture         =   "SwitchBoards.frx":2D96
         Stretch         =   -1  'True
         Top             =   1575
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   6
         Left            =   6120
         Picture         =   "SwitchBoards.frx":30A0
         Stretch         =   -1  'True
         Top             =   915
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   3840
         Picture         =   "SwitchBoards.frx":33AA
         Stretch         =   -1  'True
         Top             =   1575
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   2
         Left            =   5160
         Picture         =   "SwitchBoards.frx":36B4
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   499
         Index           =   5
         Left            =   8160
         Picture         =   "SwitchBoards.frx":39BE
         Stretch         =   -1  'True
         Top             =   1574
         Width           =   495
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "D/A 1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   3900
         MousePointer    =   1  'Arrow
         TabIndex        =   57
         Top             =   3075
         Width           =   420
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SHAFT"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1080
         MousePointer    =   1  'Arrow
         TabIndex        =   56
         Top             =   3075
         Width           =   570
      End
      Begin VB.Line Line23 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   520
         X2              =   8880
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   660
         X2              =   8940
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   390
         X2              =   8880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   1380
         X2              =   1380
         Y1              =   2425
         Y2              =   1230
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1500
         X2              =   1500
         Y1              =   2424
         Y2              =   1365
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   1230
         X2              =   1230
         Y1              =   2427
         Y2              =   1095
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   4000
         X2              =   4000
         Y1              =   2427
         Y2              =   1095
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4260
         X2              =   4260
         Y1              =   2424
         Y2              =   1365
      End
      Begin VB.Line Line13 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   4140
         X2              =   4140
         Y1              =   2395
         Y2              =   1200
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5310
         X2              =   5310
         Y1              =   2412
         Y2              =   1080
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5580
         X2              =   5580
         Y1              =   2439
         Y2              =   1380
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   5460
         X2              =   5460
         Y1              =   2425
         Y2              =   1230
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   7230
         X2              =   7230
         Y1              =   2421
         Y2              =   1089
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   7515
         X2              =   7515
         Y1              =   2421
         Y2              =   1362
      End
      Begin VB.Line Line19 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   7380
         X2              =   7380
         Y1              =   2421
         Y2              =   1226
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   8310
         X2              =   8310
         Y1              =   2421
         Y2              =   1089
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   8595
         X2              =   8595
         Y1              =   2421
         Y2              =   1362
      End
      Begin VB.Line Line22 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8460
         X2              =   8460
         Y1              =   2421
         Y2              =   1226
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   620
         Index           =   0
         Left            =   7080
         Picture         =   "SwitchBoards.frx":3CC8
         Stretch         =   -1  'True
         Top             =   2421
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EDA"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   7200
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         Top             =   3147
         Width           =   360
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "D/A 2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   5280
         MousePointer    =   1  'Arrow
         TabIndex        =   54
         Top             =   3075
         Width           =   450
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SHORE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   8160
         MousePointer    =   1  'Arrow
         TabIndex        =   53
         Top             =   3147
         Width           =   585
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EMERGENCY BUS-BAR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6780
         MousePointer    =   1  'Arrow
         TabIndex        =   52
         Top             =   840
         Width           =   1860
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN BUS-BAR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   4080
         MousePointer    =   1  'Arrow
         TabIndex        =   51
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TIE"
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
         Left            =   6240
         MousePointer    =   1  'Arrow
         TabIndex        =   50
         Top             =   1620
         Width           =   210
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BREAKER"
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
         Left            =   6060
         MousePointer    =   1  'Arrow
         TabIndex        =   49
         Top             =   1860
         Width           =   690
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   7755
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   AVR      Controls"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6240
         TabIndex        =   48
         Top             =   5160
         Width           =   765
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Speed      Controls"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6180
         TabIndex        =   47
         Top             =   6300
         Width           =   855
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1860
      TabIndex        =   75
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "SwitchBoards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim set_dummy_focus As Integer

Private Sub ALT_image_Click(index As Integer)


' 21/8/98 MOVED to Reverse_power() command

'Control_PC_check
'If This_Control_status = False Then Exit Sub

'If Index > 0 And Index < 5 Then
'  Call setPCDI(REVRESET(Index))
'End If


End Sub

Private Sub AVR_auto_Click(index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 9/3/98
Call setDI(AVRMANSW(index), False)
Call setDI(AVRAUTOSW(index), True)

End Sub

Private Sub AVR_MAN_Click(index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 9/3/98
Call setDI(AVRMANSW(index), True)
Call setDI(AVRAUTOSW(index), False)

End Sub

Private Sub breaker_Click(index As Integer)
Dim IAI_PHASE As Integer
Dim single_char As String * 1

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
If index = 0 Then
' 7/6/98
'   Call setPCDI(1098)   '     Emerg ALT BREAKER   (Open/Close)
    Call setPCDI(777)    ' LSS Emerg ALT BREAKER   (Open/Close)
    
ElseIf index > 0 And index < 5 Then

' 7/6/98
'    If elec_lamps(Index) Then
     single_char = Mid$(PC_INT_STRING, 429 + index, 1)
     If single_char = "1" Then    ' ( Breaker r is closed so Open it )
         Call setPCDI(OPENMBKSW(index))    ' OPEN the Main Breaker
     Else
       If AIPHASE > 0 And AIPHASE < 256 Then
           IAI_PHASE = AIPHASE
'         Call setAI(59, IAI_PHASE)
          Call setAI(35, IAI_PHASE)  ' (for LSS)
       End If
         Call setPCDI(MBKSW(index))        ' CLOSE the Main Breaker
     End If
     
     
ElseIf index = 5 Then
     single_char = Mid$(PC_INT_STRING, 421, 1)
     If single_char = "1" Then
     
'       20/9/99  STEAM_SIM
'        Call setPCDI(782)    ' OPEN the LSS SHORE Breaker
         Call setPCDI(752)    ' OPEN the LSS SHORE Breaker
         
     Else
     
'        Call setPCDI(790)    ' CLOSE the SHORE Breaker
         Call setPCDI(751)    ' CLOSE the SHORE Breaker
         
     End If
     
ElseIf index = 6 Then
     single_char = Mid$(PC_INT_STRING, 422, 1)
     If single_char = "1" Then
     Else
     End If
     
' 3/9/98
ElseIf index = 7 Then   ' Bow Thruster/Shaft Alt Tie Breaker
     single_char = Mid$(PC_INT_STRING, 424, 1)
     If single_char = "1" Then
         Call setPCDI(477)    ' OPEN the Bow Thruster/Shaft Alt Tie Breaker
     Else
         Call setPCDI(479)    ' CLOSE the Bow Thruster/Shaft Alt Tie Breaker
     End If
     
' 3/9/98
ElseIf index = 8 Then   ' Bow Thruster Breaker
     single_char = Mid$(PC_INT_STRING, 423, 1)
     If single_char = "1" Then
         Call setPCDI(476)    ' OPEN the Bow Thruster/Shaft Alt Tie Breaker
     Else
         Call setPCDI(478)    ' CLOSE the Bow Thruster/Shaft Alt Tie Breaker
     End If
      
     
End If

End Sub

Private Sub Command1_Click()

frm_synch.Show


End Sub

Private Sub Command2_Click()

' 12/3/98
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

ss_supply = True

End Sub

Private Sub Command3_Click()



' 4/4/00
'sy205_LSS.Show

main_SB_1.Show

End Sub

Private Sub Command3D22_Click()

End Sub

Private Sub Command4_Click()


' 4/4/00
'sy206_LSS.Show

main_SB_2.Show


End Sub

Private Sub Command5_Click()

'  16/9/99
Emergency_swbd.Show

'sy207_LSS.Show

End Sub

Private Sub Command6_Click()


' 5/6/98
Elec_ctrl_LSS.Show

End Sub

Private Sub Command7_Click()

' 12/3/98
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

ss_supply = False

End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idsy200)

    
'   AIvalue = 255 - VScroll1(Index).Value
'   Call setAI(AVR_pot(Index), AIvalue)
  
' 22/7/98 THIS IS MEDIUM SPEED CODE:
'  VScroll1(1).Value = 255 - VB_pots(AVR_pot(1))  ' D/A 1
'  VScroll1(2).Value = 255 - VB_pots(AVR_pot(2))  ' D/A 2
'  VScroll1(4).Value = 255 - VB_pots(AVR_pot(4))  ' SHAFT
    
' 22/7/98
update_VBAI(13) = True
update_VBAI(14) = True
update_VBAI(15) = True
update_VBAI(16) = True
   
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy200)
    HavenStretch1.Visible = True
End If
End Sub

Private Sub Image8_Click()
End Sub


Private Sub Turbo_Change()

End Sub

Private Sub OFFpump_Click(index As Integer)

End Sub

Private Sub ONpump_Click(index As Integer)


End Sub

Private Sub sbyPUMP_Click(index As Integer)


End Sub

Private Sub Lower_Click(index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If index = 1 Then
' Call setPCDI(1630)   ' DECGOV1
  Call setPCDI(587)    ' DECGOV1
ElseIf index = 2 Then
' Call setPCDI(1631)   ' DECGOV2
  Call setPCDI(596)    ' DECGOV2
ElseIf index = 3 Then
  Call setPCDI(468)    ' DECGOV3 21/8/98
End If

End Sub

Private Sub Raise_Click(index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If index = 1 Then
' Call setPCDI(1628)   ' INCGOV1
  Call setPCDI(595)    ' INCGOV1
ElseIf index = 2 Then
' Call setPCDI(1629)   ' INCGOV2
  Call setPCDI(588)    ' INCGOV2
ElseIf index = 3 Then
 Call setPCDI(460)     ' INCGOV3  21/8/98
End If

End Sub

Private Sub Reverse_power_Click(index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If index > 0 And index < 5 Then
  Call setPCDI(REVRESET(index))
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim single_char As String * 1


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' 22/7/98 ( see TakeSnapForm and HOST.frm )
' D/A 1 AVR POT
If update_VBAI(15) Then
   update_VBAI(15) = False
   VScroll1(1).Value = 255 - VB_AI(15)  ' VB_AI(15) = AI(32) = AVR_POT(1)
End If
' D/A 2 AVR POT
If update_VBAI(16) Then
   update_VBAI(16) = False
   VScroll1(2).Value = 255 - VB_AI(16)  ' VB_AI(16) = AI(34) = AVR_POT(2)
End If

' 20/8/98
' Turbo Alt AVR POT
If update_VBAI(14) Then
   update_VBAI(14) = False
   VScroll1(3).Value = 255 - VB_AI(14)  ' VB_AI(14) = AI(26) = AVR_POT(3)
End If

' Shaft Alt AVR POT
If update_VBAI(13) Then
   update_VBAI(13) = False
   VScroll1(4).Value = 255 - VB_AI(13)  ' VB_AI(13) = AI(24) = AVR_POT(4)
End If



' Emergency Alternator
'valt4.Text = Format(elec_reals(1), "###0")
'aalt4.Text = Format(elec_reals(2), "###0")
'walt4.Text = Format(elec_reals(3), "###0")
valt4.text = Format(LSS_reals(39), "###0")
aalt4.text = Format(LSS_reals(40), "###0")
walt4.text = Format(LSS_reals(41), "###0")

' Shore Supply
vesrs.text = Format(LSS_reals(42), "###0")
ashore.text = Format(LSS_reals(43), "###0")
wshore.text = Format(LSS_reals(44), "###0")

' ALTERNATORS
For i = 1 To 4

   ' 20/8/98 ADDED T/A
 ' Skip the T/A
 'If i <> 3 Then
  
    HEALT(i).text = Format(LSS_reals(574 + i), "#0.0")
    
    vsense(i).text = Format(LSS_reals(9 + i), "###0")
    aealt(i).text = Format(LSS_reals(13 + i), "###0")
    wealt(i).text = Format(LSS_reals(17 + i), "###0")
    oalt(i).text = Format(LSS_reals(21 + i), "#0.00")
    ALT_exc(i).text = Format(LSS_reals(25 + i), "##0")
   
    ALT_rpm(i).text = Format(LSS_reals(44 + i), "##0")
    
    
  'End If
   
Next i

' Emergency Bus-Bars
webes.text = Format(LSS_reals(36), "###0")
aebes.text = Format(LSS_reals(35), "###0")
vebes.text = Format(LSS_reals(34), "###0")

' Main Bus-Bars
wbusb.text = Format(LSS_reals(32), "###0")
abusb.text = Format(LSS_reals(31), "###0")
vembs.text = Format(LSS_reals(30), "###0")
hembs.text = Format(LSS_reals(38), "#0.0")
powfact.text = Format(LSS_reals(33), "#0.00")


' Emergency ALT ON and OFF
'If elec_lamps(14) Then
single_char = Mid$(PC_INT_STRING, 394, 1)
If single_char = "1" Then
   ALT_image(0).Picture = GEN_on.Picture
Else
   ALT_image(0).Picture = GEN_off.Picture
End If

' Emergency ALT Power Available
'If elec_lamps(15) Then
single_char = Mid$(PC_INT_STRING, 410, 1)
If single_char = "1" Then
   Power_available(0).FillColor = &HFFFFFF     ' WHITE
Else
   Power_available(0).FillColor = &HC0C0C0     ' GREY
End If

' Emergency ALT Breaker
'If elec_lamps(13) Then
single_char = Mid$(PC_INT_STRING, 396, 1)
If single_char = "1" Then
   breaker(0).Picture = VBKR_closed.Picture
Else
   breaker(0).Picture = VBKR_open.Picture
End If


' SHORE SUPPLY ON and OFF
'If elec_lamps(18) Then
single_char = Mid$(PC_INT_STRING, 415, 1)
If single_char = "1" Then
   ALT_image(5).Picture = GEN_on.Picture
Else
   ALT_image(5).Picture = GEN_off.Picture
End If

' SHORE SUPPLY Power Available
'If elec_lamps(19) Then
single_char = Mid$(PC_INT_STRING, 409, 1)
If single_char = "1" Then
   Power_available(5).FillColor = &HFFFFFF     ' WHITE
Else
   Power_available(5).FillColor = &HC0C0C0     ' GREY
End If

' SHORE supply Breaker
'If elec_lamps(17) Then
single_char = Mid$(PC_INT_STRING, 421, 1)
If single_char = "1" Then
   breaker(5).Picture = VBKR_closed.Picture
Else
   breaker(5).Picture = VBKR_open.Picture
End If


' Tie-Breaker
'If elec_lamps(20) Then
single_char = Mid$(PC_INT_STRING, 422, 1)
If single_char = "1" Then
   breaker(6).Picture = HBKR_closed.Picture
Else
   breaker(6).Picture = HBKR_open.Picture
End If

' 3/9/98 Bow Thruster/Shaft Alt Tie-Breaker
single_char = Mid$(PC_INT_STRING, 424, 1)
If single_char = "1" Then
   breaker(7).Picture = HBKR_closed.Picture
Else
   breaker(7).Picture = HBKR_open.Picture
End If

' 3/9/98 Bow Thruster Breaker
single_char = Mid$(PC_INT_STRING, 423, 1)
If single_char = "1" Then
   breaker(8).Picture = VBKR_closed.Picture
Else
   breaker(8).Picture = VBKR_open.Picture
End If

' =======   MAIN ALTERNATORS   ============

For i = 1 To 4


 ' 20/8/98 ADDED T/A
 ' Skip the T/A
 ' If i <> 3 Then
  
'   ALTERNATOR Pictures
'    If elec_lamps(i + 4) Then
     single_char = Mid$(PC_INT_STRING, 425 + i, 1)
     If single_char = "2" Or single_char = "3" Then
           ALT_image(i).Picture = GEN_on.Picture
     Else
           ALT_image(i).Picture = GEN_off.Picture
     End If
     
'   ALTERNATOR MAIN BREAKERS
'     If elec_lamps(i) Then
      single_char = Mid$(PC_INT_STRING, 429 + i, 1)
     If single_char = "1" Then
       breaker(i).Picture = VBKR_closed.Picture
     Else
       breaker(i).Picture = VBKR_open.Picture
     End If
     
'   ALTERNATOR POWER AVAILABLE LAMPS
'     If elec_lamps(i + 8) Then
      single_char = Mid$(PC_INT_STRING, 425 + i, 1)
     If single_char = "1" Or single_char = "3" Then
       Power_available(i).FillColor = &HFFFFFF     ' WHITE
     Else
       Power_available(i).FillColor = &HC0C0C0     ' GREY
     End If
     
'   ALTERNATOR REVERSE POWER LAMPS
'     If elec_lamps(i + 41) Then
       single_char = Mid$(PC_INT_STRING, 429 + i, 1)
       
    ' === 21/8/98 ===
    ' If single_char = "2" Then
    '   REV_power(i).FillColor = &HFF&         ' RED
    ' Else
    '   REV_power(i).FillColor = &HC0C0C0      ' GREY
    ' End If
      If single_char = "2" Then
        Reverse_power(i).BackColor = &HFF&         ' RED
      Else
        Reverse_power(i).BackColor = &HC0C0C0      ' GREY
      End If
     
     
'   ALTERNATOR AUTO AVR Swithes
'    If elec_lamps(i + 45) Then
     If LSS_LOGICALS(i + 4) Then
       AVR_auto(i).BackColor = &HFF80FF      ' LIGHT PURPLE
        AVR_MAN(i).BackColor = &HC0C0C0      ' GREY
     Else
       AVR_auto(i).BackColor = &HC0C0C0      ' GREY
        AVR_MAN(i).BackColor = &HFFFF&       ' YELLOW
     End If
     
  ' End If
  
Next i

' =========================================

End Sub

Private Sub VScroll1_Change(index As Integer)
Dim AIvalue As Integer

' 20/3/98  set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 11/3/98
'  AIvalue = VScroll1(Index).Value
   AIvalue = 255 - VScroll1(index).Value
  Call setAI(AVR_pot(index), AIvalue)
  
End Sub
