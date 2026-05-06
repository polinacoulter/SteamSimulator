VERSION 5.00
Begin VB.Form Main_SB_new 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SY200 - MAIN SWITCHBOARD"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   10620
   Begin VB.Frame HavenStretch1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7635
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10275
      Begin VB.TextBox hembs 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   122
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox powfact 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   121
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox vembs 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Left            =   3600
         MousePointer    =   1  'Arrow
         TabIndex        =   120
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox abusb 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Left            =   3600
         MousePointer    =   1  'Arrow
         TabIndex        =   119
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox wbusb 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Left            =   4740
         MousePointer    =   1  'Arrow
         TabIndex        =   118
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   4
         Left            =   5340
         MousePointer    =   1  'Arrow
         TabIndex        =   117
         Text            =   " kW"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   3
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   116
         Text            =   " A"
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   2
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   115
         Text            =   " V"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   0
         Left            =   3180
         MousePointer    =   1  'Arrow
         TabIndex        =   114
         Text            =   " Hz"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   1
         Left            =   3180
         MousePointer    =   1  'Arrow
         TabIndex        =   113
         Text            =   " PF"
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox vebes 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Left            =   6780
         MousePointer    =   1  'Arrow
         TabIndex        =   111
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox aebes 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6780
         MousePointer    =   1  'Arrow
         TabIndex        =   110
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox webes 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Left            =   7680
         MousePointer    =   1  'Arrow
         TabIndex        =   109
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   7
         Left            =   8220
         MousePointer    =   1  'Arrow
         TabIndex        =   108
         Text            =   " kW"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   6
         Left            =   7320
         MousePointer    =   1  'Arrow
         TabIndex        =   107
         Text            =   " A"
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   287
         Index           =   5
         Left            =   7320
         MousePointer    =   1  'Arrow
         TabIndex        =   106
         Text            =   " V"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Height          =   255
         Left            =   9600
         TabIndex        =   105
         Top             =   360
         Width           =   195
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text40 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   5580
         MousePointer    =   1  'Arrow
         TabIndex        =   92
         Text            =   "RPM"
         Top             =   3420
         Width           =   495
      End
      Begin VB.TextBox Text36 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   91
         Text            =   " Hz"
         Top             =   3660
         Width           =   495
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   90
         Text            =   " Hz"
         Top             =   3660
         Width           =   435
      End
      Begin VB.TextBox Text37 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000007&
         Height          =   250
         Left            =   2820
         MousePointer    =   1  'Arrow
         TabIndex        =   89
         Text            =   "RPM"
         Top             =   3420
         Width           =   435
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5580
         MousePointer    =   1  'Arrow
         TabIndex        =   88
         Text            =   "VEx"
         Top             =   5040
         Width           =   555
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   87
         Text            =   " kW"
         Top             =   4380
         Width           =   495
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   86
         Text            =   " A"
         Top             =   4140
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   85
         Text            =   " V"
         Top             =   3900
         Width           =   495
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   84
         Text            =   " PF"
         Top             =   4620
         Width           =   375
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   83
         Text            =   " kW"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   7320
         MousePointer    =   1  'Arrow
         TabIndex        =   82
         Text            =   " kW"
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   81
         Text            =   " A"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   7320
         MousePointer    =   1  'Arrow
         TabIndex        =   80
         Text            =   " A"
         Top             =   3720
         Width           =   375
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   79
         Text            =   " kW"
         Top             =   4380
         Width           =   375
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   78
         Text            =   " A"
         Top             =   4140
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   8520
         MousePointer    =   1  'Arrow
         TabIndex        =   77
         Text            =   " V"
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   7320
         MousePointer    =   1  'Arrow
         TabIndex        =   76
         Text            =   " V"
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   2820
         MousePointer    =   1  'Arrow
         TabIndex        =   75
         Text            =   "V"
         Top             =   3900
         Width           =   315
      End
      Begin VB.TextBox wshore 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         TabIndex        =   74
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox ashore 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         TabIndex        =   73
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox vesrs 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         TabIndex        =   72
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox walt4 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
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
         Left            =   6720
         MousePointer    =   1  'Arrow
         TabIndex        =   71
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox aalt4 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
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
         Left            =   6720
         MousePointer    =   1  'Arrow
         TabIndex        =   70
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox valt4 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
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
         Left            =   6720
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Top             =   3480
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1200
         Top             =   6420
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   68
         Text            =   " PF"
         Top             =   4620
         Width           =   495
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   5580
         Visible         =   0   'False
         Width           =   615
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5820
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Top             =   3660
         Width           =   675
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Top             =   3660
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Top             =   3900
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   62
         Top             =   3900
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Top             =   4140
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   60
         Top             =   4140
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   59
         Top             =   4380
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   58
         Top             =   4380
         Width           =   675
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   57
         Top             =   4620
         Width           =   675
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   56
         Top             =   4620
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         Top             =   3540
         Visible         =   0   'False
         Width           =   555
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   54
         Top             =   3780
         Visible         =   0   'False
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   53
         Top             =   4020
         Visible         =   0   'False
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   52
         Top             =   4260
         Visible         =   0   'False
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
         Index           =   4
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   51
         Top             =   4500
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   50
         Top             =   3420
         Width           =   675
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   49
         Top             =   3420
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
         Left            =   840
         MousePointer    =   1  'Arrow
         TabIndex        =   48
         Top             =   3300
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   287
         Index           =   1
         Left            =   2100
         MousePointer    =   1  'Arrow
         TabIndex        =   47
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Index           =   2
         Left            =   4860
         MousePointer    =   1  'Arrow
         TabIndex        =   46
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
         Index           =   4
         Left            =   900
         MousePointer    =   1  'Arrow
         TabIndex        =   45
         Top             =   5220
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S/B 1"
         Height          =   675
         Left            =   1500
         Picture         =   "Main_SB_new.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S/B 2"
         Height          =   675
         Left            =   1500
         Picture         =   "Main_SB_new.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1620
         Width           =   675
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESB"
         Height          =   675
         Left            =   8700
         Picture         =   "Main_SB_new.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1500
         Width           =   675
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Control"
         Height          =   675
         Left            =   8160
         Picture         =   "Main_SB_new.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   6540
         Width           =   675
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   2
         Left            =   5820
         Max             =   255
         TabIndex        =   40
         Top             =   5340
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   1
         Left            =   2880
         Max             =   255
         TabIndex        =   39
         Top             =   5340
         Width           =   255
      End
      Begin VB.CommandButton AVR_MAN 
         Caption         =   "MAN"
         Height          =   257
         Index           =   1
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton AVR_MAN 
         Caption         =   "MAN"
         Height          =   257
         Index           =   2
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5400
         Width           =   675
      End
      Begin VB.CommandButton AVR_auto 
         BackColor       =   &H00FF80FF&
         Caption         =   "AUTO"
         Height          =   257
         Index           =   1
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5640
         Width           =   615
      End
      Begin VB.CommandButton AVR_auto 
         BackColor       =   &H00FF80FF&
         Caption         =   "AUTO"
         Height          =   257
         Index           =   2
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5640
         Width           =   675
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   4
         Left            =   1620
         Max             =   255
         TabIndex        =   34
         Top             =   5040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Raise 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   2640
         Picture         =   "Main_SB_new.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6540
         Width           =   555
      End
      Begin VB.CommandButton Raise 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   5400
         Picture         =   "Main_SB_new.frx":18F2
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6540
         Width           =   495
      End
      Begin VB.CommandButton Lower 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   2040
         Picture         =   "Main_SB_new.frx":234C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6540
         Width           =   555
      End
      Begin VB.CommandButton Lower 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   4860
         Picture         =   "Main_SB_new.frx":2DA6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   6540
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SScope"
         Height          =   675
         Left            =   8160
         Picture         =   "Main_SB_new.frx":3800
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5640
         Width           =   675
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C000&
         Caption         =   "ON"
         Height          =   435
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H000000FF&
         Caption         =   "OFF"
         Height          =   435
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1035
         Index           =   3
         Left            =   4380
         Max             =   255
         TabIndex        =   26
         Top             =   5340
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   4200
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         Text            =   "RPM"
         Top             =   3420
         Width           =   495
      End
      Begin VB.TextBox ALT_rpm 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   3420
         Width           =   675
      End
      Begin VB.TextBox ALT_exc 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   287
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Top             =   5040
         Width           =   675
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2820
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Text            =   "VEx"
         Top             =   5040
         Width           =   435
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4200
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Text            =   "VEx"
         Top             =   5040
         Width           =   495
      End
      Begin VB.CommandButton AVR_MAN 
         Caption         =   "MAN"
         Height          =   257
         Index           =   3
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton AVR_auto 
         BackColor       =   &H00FF80FF&
         Caption         =   "AUTO"
         Height          =   257
         Index           =   3
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Text            =   " PF"
         Top             =   4620
         Width           =   495
      End
      Begin VB.TextBox oalt 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Top             =   4620
         Width           =   675
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Text            =   " kW"
         Top             =   4380
         Width           =   495
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Text            =   " A"
         Top             =   4140
         Width           =   495
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Text            =   " V"
         Top             =   3900
         Width           =   495
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   4140
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Text            =   " Hz"
         Top             =   3660
         Width           =   495
      End
      Begin VB.TextBox HEALT 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   3660
         Width           =   675
      End
      Begin VB.TextBox vsense 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   3900
         Width           =   675
      End
      Begin VB.TextBox aealt 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   4140
         Width           =   675
      End
      Begin VB.TextBox wealt 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   250
         Index           =   3
         Left            =   3480
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   4380
         Width           =   675
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
         Left            =   900
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Text            =   "VEx"
         Top             =   5040
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Raise 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   4020
         Picture         =   "Main_SB_new.frx":3B0A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6540
         Width           =   495
      End
      Begin VB.CommandButton Lower 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   3420
         Picture         =   "Main_SB_new.frx":4564
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6540
         Width           =   555
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1860
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
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1860
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
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1860
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
         Index           =   4
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   6
         Left            =   9300
         Shape           =   3  'Circle
         Top             =   2760
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OLE Sin_wave 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   4
         Left            =   1040
         OleObjectBlob   =   "Main_SB_new.frx":4FBE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Sin_wave.vsd"
         TabIndex        =   131
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE Sin_wave 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   3800
         OleObjectBlob   =   "Main_SB_new.frx":C7D6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Sin_wave.vsd"
         TabIndex        =   130
         Top             =   2520
         Width           =   495
      End
      Begin VB.OLE Sin_wave 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   5120
         OleObjectBlob   =   "Main_SB_new.frx":13FEE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Sin_wave.vsd"
         TabIndex        =   129
         Top             =   2520
         Width           =   495
      End
      Begin VB.OLE Sin_wave 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   2420
         OleObjectBlob   =   "Main_SB_new.frx":1B806
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Sin_wave.vsd"
         TabIndex        =   128
         Top             =   2520
         Width           =   495
      End
      Begin VB.Line Top_Line 
         BorderWidth     =   4
         Index           =   4
         Visible         =   0   'False
         X1              =   960
         X2              =   1500
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   4
         Left            =   960
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   3
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   555
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   2
         Left            =   5040
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   555
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   1
         Left            =   2340
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   555
      End
      Begin VB.Line Top_Line 
         BorderWidth     =   4
         Index           =   3
         X1              =   3720
         X2              =   4260
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Top_Line 
         BorderWidth     =   4
         Index           =   2
         X1              =   5040
         X2              =   5580
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Top_Line 
         BorderWidth     =   4
         Index           =   0
         X1              =   6960
         X2              =   7500
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Top_Line 
         BorderWidth     =   4
         Index           =   1
         X1              =   2340
         X2              =   2880
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.OLE Sin_wave 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   7040
         OleObjectBlob   =   "Main_SB_new.frx":2301E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Sin_wave.vsd"
         TabIndex        =   127
         Top             =   2520
         Width           =   495
      End
      Begin VB.OLE SSG_ON 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   915
         Index           =   5
         Left            =   9120
         OleObjectBlob   =   "Main_SB_new.frx":2A836
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\GEN_on_NEW.vsd"
         TabIndex        =   126
         Top             =   4320
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OLE SSG_OFF 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   795
         Index           =   5
         Left            =   9180
         OleObjectBlob   =   "Main_SB_new.frx":2D84E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Gen_OFF_new.vsd"
         TabIndex        =   125
         Top             =   3480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Line Top_Line 
         BorderWidth     =   4
         Index           =   5
         X1              =   8040
         X2              =   8580
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN BUS"
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
         Left            =   4800
         MousePointer    =   1  'Arrow
         TabIndex        =   123
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EMERG BUS"
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
         Left            =   7680
         MousePointer    =   1  'Arrow
         TabIndex        =   112
         Top             =   660
         Width           =   990
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
         Left            =   6360
         TabIndex        =   104
         Top             =   6300
         Visible         =   0   'False
         Width           =   855
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
         Left            =   6360
         TabIndex        =   103
         Top             =   5520
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   7680
         Shape           =   3  'Circle
         Top             =   2340
         Width           =   255
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
         Left            =   5940
         MousePointer    =   1  'Arrow
         TabIndex        =   102
         Top             =   1620
         Width           =   690
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
         Left            =   6120
         MousePointer    =   1  'Arrow
         TabIndex        =   101
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SHORE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8100
         MousePointer    =   1  'Arrow
         TabIndex        =   100
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#2 SSTG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4920
         MousePointer    =   1  'Arrow
         TabIndex        =   99
         Top             =   3120
         Width           =   810
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EDA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7080
         MousePointer    =   1  'Arrow
         TabIndex        =   98
         Top             =   3120
         Width           =   420
      End
      Begin VB.Image ALT_image 
         Appearance      =   0  'Flat
         Height          =   855
         Index           =   0
         Left            =   9180
         Picture         =   "Main_SB_new.frx":30666
         Stretch         =   -1  'True
         Top             =   5400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1110
         X2              =   1110
         Y1              =   2427
         Y2              =   1095
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1380
         X2              =   1380
         Y1              =   2424
         Y2              =   1365
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1260
         X2              =   1260
         Y1              =   2425
         Y2              =   1230
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
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   97
         Top             =   3075
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SSDG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2340
         MousePointer    =   1  'Arrow
         TabIndex        =   96
         Top             =   3120
         Width           =   570
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   499
         Index           =   5
         Left            =   8040
         Picture         =   "Main_SB_new.frx":30970
         Stretch         =   -1  'True
         Top             =   1574
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   2
         Left            =   5040
         Picture         =   "Main_SB_new.frx":30C7A
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   3
         Left            =   3720
         Picture         =   "Main_SB_new.frx":30F84
         Stretch         =   -1  'True
         Top             =   1575
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   6
         Left            =   6000
         Picture         =   "Main_SB_new.frx":3128E
         Stretch         =   -1  'True
         Top             =   915
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   4
         Left            =   960
         Picture         =   "Main_SB_new.frx":31598
         Stretch         =   -1  'True
         Top             =   1575
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image VBKR_closed 
         Height          =   480
         Left            =   480
         Picture         =   "Main_SB_new.frx":318A2
         Top             =   5880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image VBKR_open 
         Height          =   480
         Left            =   0
         Picture         =   "Main_SB_new.frx":31BAC
         Top             =   4980
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image GEN_off 
         Height          =   900
         Left            =   9120
         Picture         =   "Main_SB_new.frx":31EB6
         Stretch         =   -1  'True
         Top             =   6420
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image GEN_on 
         Height          =   480
         Left            =   0
         Picture         =   "Main_SB_new.frx":321C0
         Top             =   6300
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image HBKR_open 
         Height          =   480
         Left            =   0
         Picture         =   "Main_SB_new.frx":324CA
         Top             =   4080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image HBKR_closed 
         Height          =   480
         Left            =   0
         Picture         =   "Main_SB_new.frx":327D4
         Top             =   4500
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   6960
         Picture         =   "Main_SB_new.frx":32ADE
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   495
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   3060
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   5820
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   2460
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Power_available 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   8880
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   255
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   3420
         X2              =   3420
         Y1              =   1560
         Y2              =   900
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   3300
         X2              =   3300
         Y1              =   1560
         Y2              =   900
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   3180
         X2              =   3180
         Y1              =   1560
         Y2              =   900
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   3975
         Left            =   1920
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   5820
         Shape           =   3  'Circle
         Top             =   2340
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   3060
         Shape           =   3  'Circle
         Top             =   2340
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   7680
         Shape           =   3  'Circle
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1620
         Shape           =   3  'Circle
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "#1 SSTG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3600
         MousePointer    =   1  'Arrow
         TabIndex        =   95
         Top             =   3120
         Width           =   810
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   2340
         Picture         =   "Main_SB_new.frx":32DE8
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   495
      End
      Begin VB.Shape REV_power 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   2340
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   7
         Left            =   420
         Picture         =   "Main_SB_new.frx":330F2
         Stretch         =   -1  'True
         Top             =   660
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image breaker 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   8
         Left            =   120
         Picture         =   "Main_SB_new.frx":333FC
         Stretch         =   -1  'True
         Top             =   1620
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SWBD #4"
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
         Left            =   3000
         MousePointer    =   1  'Arrow
         TabIndex        =   94
         Top             =   1980
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   3975
         Left            =   3300
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   3975
         Left            =   4740
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   6600
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         Height          =   975
         Left            =   7800
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Line Line22 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8340
         X2              =   8340
         Y1              =   2421
         Y2              =   1226
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   8475
         X2              =   8475
         Y1              =   2421
         Y2              =   1362
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   8190
         X2              =   8190
         Y1              =   2421
         Y2              =   1089
      End
      Begin VB.Line Line19 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   7260
         X2              =   7260
         Y1              =   2421
         Y2              =   1226
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   7395
         X2              =   7395
         Y1              =   2421
         Y2              =   1362
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   7110
         X2              =   7110
         Y1              =   2421
         Y2              =   1089
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   1440
         X2              =   8760
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   1440
         X2              =   8820
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line23 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   1440
         X2              =   8760
         Y1              =   1200
         Y2              =   1215
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   5340
         X2              =   5340
         Y1              =   2425
         Y2              =   1230
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   5460
         X2              =   5460
         Y1              =   2439
         Y2              =   1380
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5190
         X2              =   5190
         Y1              =   2412
         Y2              =   1080
      End
      Begin VB.Line Line13 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   4020
         X2              =   4020
         Y1              =   2395
         Y2              =   1200
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   4140
         X2              =   4140
         Y1              =   2424
         Y2              =   1365
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   3880
         X2              =   3880
         Y1              =   2427
         Y2              =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   2780
         X2              =   2780
         Y1              =   2439
         Y2              =   1380
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2640
         X2              =   2640
         Y1              =   2395
         Y2              =   1200
      End
      Begin VB.Line Line24 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   2490
         X2              =   2490
         Y1              =   2412
         Y2              =   1080
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   795
         Left            =   6600
         Top             =   120
         Width           =   2235
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   795
         Left            =   2520
         Top             =   120
         Width           =   3375
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   0
         Left            =   6960
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   555
      End
      Begin VB.OLE Sin_wave 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   555
         Index           =   5
         Left            =   8120
         OleObjectBlob   =   "Main_SB_new.frx":33706
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Sin_wave.vsd"
         TabIndex        =   124
         Top             =   2520
         Width           =   495
      End
      Begin VB.Shape Gen_Shape 
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   5
         Left            =   8040
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   555
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1860
      TabIndex        =   0
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "Main_SB_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim set_dummy_focus As Integer

Private Sub ALT_image_Click(Index As Integer)


' 21/8/98 MOVED to Reverse_power() command

'Control_PC_check
'If This_Control_status = False Then Exit Sub

'If Index > 0 And Index < 5 Then
'  Call setPCDI(REVRESET(Index))
'End If


End Sub

Private Sub AVR_auto_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 9/3/98
Call setDI(AVRMANSW(Index), False)
Call setDI(AVRAUTOSW(Index), True)

End Sub

Private Sub AVR_MAN_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 9/3/98
Call setDI(AVRMANSW(Index), True)
Call setDI(AVRAUTOSW(Index), False)

End Sub

Private Sub breaker_Click(Index As Integer)
Dim IAI_PHASE As Integer
Dim single_char As String * 1

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
If Index = 0 Then

    ' 7/6/98
    '   Call setPCDI(1098)   '     Emerg ALT BREAKER   (Open/Close)
    
    
    ' === 21/01/02 ===
    ' PCDI (748), PC_MERGBKSW

   ' Call setPCDI(777)    ' LSS Emerg ALT BREAKER   (Open/Close)
     Call setPCDI(748)    ' LSS Emerg ALT BREAKER   (Open/Close)
    
    
    
ElseIf Index > 0 And Index < 5 Then

' 7/6/98
'    If elec_lamps(Index) Then
     single_char = Mid$(PC_INT_STRING, 429 + Index, 1)
     If single_char = "1" Then    ' ( Breaker r is closed so Open it )
         Call setPCDI(OPENMBKSW(Index))    ' OPEN the Main Breaker
     Else
       If AIPHASE > 0 And AIPHASE < 256 Then
           IAI_PHASE = AIPHASE
'         Call setAI(59, IAI_PHASE)
          Call setAI(35, IAI_PHASE)  ' (for LSS)
       End If
         Call setPCDI(MBKSW(Index))        ' CLOSE the Main Breaker
     End If
     
     
ElseIf Index = 5 Then
     single_char = Mid$(PC_INT_STRING, 421, 1)
     If single_char = "1" Then
     
'       20/9/99  STEAM_SIM
'        Call setPCDI(782)    ' OPEN the LSS SHORE Breaker
         Call setPCDI(752)    ' OPEN the LSS SHORE Breaker
         
     Else
     
'        Call setPCDI(790)    ' CLOSE the SHORE Breaker
         Call setPCDI(751)    ' CLOSE the SHORE Breaker
         
     End If
     
ElseIf Index = 6 Then
     single_char = Mid$(PC_INT_STRING, 422, 1)
     If single_char = "1" Then
     Else
     End If
     
' 3/9/98
ElseIf Index = 7 Then   ' Bow Thruster/Shaft Alt Tie Breaker
     single_char = Mid$(PC_INT_STRING, 424, 1)
     If single_char = "1" Then
         Call setPCDI(477)    ' OPEN the Bow Thruster/Shaft Alt Tie Breaker
     Else
         Call setPCDI(479)    ' CLOSE the Bow Thruster/Shaft Alt Tie Breaker
     End If
     
' 3/9/98
ElseIf Index = 8 Then   ' Bow Thruster Breaker
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

'sy205_LSS.Show
main_SB_1.Show

End Sub

Private Sub Command3D22_Click()

End Sub

Private Sub Command4_Click()

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

Private Sub Command8_Click()

White_background = Not White_background
If White_background Then
  HavenStretch1.BackColor = &HFFFFFF     ' WHITE
Else
  HavenStretch1.BackColor = &HC0C0C0     'GREY
End If

End Sub

Private Sub Form_Load()
Dim i As Integer

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


' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = False
    
  ' 6/9/00
  For i = 1 To 3
      Raise(i).Visible = False
      Lower(i).Visible = False
  Next i

  
Else
    Command3.Visible = True
    Command4.Visible = True
    Command5.Visible = True

  ' 6/9/00
  For i = 1 To 3
      Raise(i).Visible = True
      Lower(i).Visible = True
  Next i
  
End If

   
    
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

Private Sub OFFpump_Click(Index As Integer)

End Sub

Private Sub ONpump_Click(Index As Integer)


End Sub

Private Sub sbyPUMP_Click(Index As Integer)


End Sub

Private Sub Lower_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


'  30/9/99   see EIO.COM in FORTRAN
'       EQUIVALENCE (DI( 835),INCGOV3),(DI( 836),DECGOV3),
'     +             (DI( 944),INCGOV1),(DI( 943),DECGOV1),
'     +             (DI( 845),INCGOV2),(DI( 846),DECGOV2)



If Index = 1 Then
  Call setPCDI(943)    ' DECGOV1
ElseIf Index = 2 Then
  Call setPCDI(846)    ' DECGOV2
ElseIf Index = 3 Then
  Call setPCDI(836)    ' DECGOV3
End If

End Sub

Private Sub Raise_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

'  30/9/99   see EIO.COM in FORTRAN
'       EQUIVALENCE (DI( 835),INCGOV3),(DI( 836),DECGOV3),
'     +             (DI( 944),INCGOV1),(DI( 943),DECGOV1),
'     +             (DI( 845),INCGOV2),(DI( 846),DECGOV2)



If Index = 1 Then
 ' Call setPCDI(595)    ' INCGOV1
   Call setPCDI(944)    ' INCGOV1
ElseIf Index = 2 Then
 ' Call setPCDI(588)    ' INCGOV2
   Call setPCDI(845)    ' INCGOV2
ElseIf Index = 3 Then
 ' Call setPCDI(460)     ' INCGOV3
   Call setPCDI(835)     ' INCGOV3
End If

End Sub

Private Sub Reverse_power_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If Index > 0 And Index < 5 Then
  Call setPCDI(REVRESET(Index))
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim single_char As String * 1


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If



    ' 22/10/99
If Not PTT Then

' We are a full mission sim ( Hardware exists )

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

  ' === 28/01/02 ===
   'ALT_image(0).Picture = GEN_on.Picture
    Gen_Shape(0).FillColor = &HFF00&    ' GREEN
Else

  ' === 28/01/02 ===
  ' ALT_image(0).Picture = GEN_off.Picture
    Gen_Shape(0).FillColor = &HFFFFFF      'WHITE
   
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

  ' === 28/01/02 ===
   'ALT_image(5).Picture = GEN_on.Picture
    Gen_Shape(5).FillColor = &HFF00&    ' GREEN

Else
  ' === 28/01/02 ===
  ' ALT_image(5).Picture = GEN_off.Picture
    Gen_Shape(5).FillColor = &HFFFFFF      'WHITE
    
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


' === 28/01/02 ===
'For i = 1 To 4
For i = 1 To 3


 ' 20/8/98 ADDED T/A
 ' Skip the T/A
 ' If i <> 3 Then
  
'   ALTERNATOR Pictures
'    If elec_lamps(i + 4) Then
     single_char = Mid$(PC_INT_STRING, 425 + i, 1)
     
     If single_char = "2" Or single_char = "3" Then
     
        ' === 28/01/02 ===
        '   ALT_image(i).Picture = GEN_on.Picture
            Gen_Shape(i).FillColor = &HFF00&    ' GREEN
    Else
     
        ' === 28/01/02 ===
        '   ALT_image(i).Picture = GEN_off.Picture
            Gen_Shape(i).FillColor = &HFFFFFF      'WHITE
           
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


' === 28/01/02 ===
'SSG(0).Visible = True
'SSG(2).Visible = True
'SSG(3).Visible = True


End Sub

Private Sub VScroll1_Change(Index As Integer)
Dim AIvalue As Integer

' 20/3/98  set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 11/3/98
'  AIvalue = VScroll1(Index).Value
   AIvalue = 255 - VScroll1(Index).Value
  Call setAI(AVR_pot(Index), AIvalue)
  
End Sub
