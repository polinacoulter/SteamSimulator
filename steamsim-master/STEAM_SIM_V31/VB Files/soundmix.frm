VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form SoundMix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Mixer"
   ClientHeight    =   11745
   ClientLeft      =   3105
   ClientTop       =   1605
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11745
   ScaleWidth      =   7860
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   52
      Top             =   10980
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   51
      Top             =   10980
      Width           =   975
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4740
      TabIndex        =   47
      Top             =   10980
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   46
      Top             =   10980
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   " Sound Mixer 1"
      Height          =   2115
      Left            =   120
      TabIndex        =   5
      Top             =   8700
      Width           =   7635
      Begin VB.CheckBox Check1 
         Caption         =   "Spare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   96
         Top             =   1680
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed1 
         Height          =   255
         Index           =   4
         Left            =   4920
         Max             =   1000
         TabIndex        =   95
         Top             =   1680
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume1 
         Height          =   255
         Index           =   4
         Left            =   3120
         Max             =   5000
         TabIndex        =   94
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox FreqVal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6600
         TabIndex        =   93
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Spare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   92
         Top             =   1320
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed1 
         Height          =   255
         Index           =   3
         Left            =   4920
         Max             =   1000
         TabIndex        =   91
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume1 
         Height          =   255
         Index           =   3
         Left            =   3120
         Max             =   5000
         TabIndex        =   90
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox FreqVal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6600
         TabIndex        =   89
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox FreqVal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox FreqVal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   44
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HASP 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.HScrollBar scrVolume1 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   5000
         TabIndex        =   33
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed1 
         Height          =   255
         Index           =   2
         Left            =   4920
         Max             =   1000
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "# 1 Turbo Feed Pump (R)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2835
      End
      Begin VB.HScrollBar scrVolume1 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   5000
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed1 
         Height          =   255
         Index           =   1
         Left            =   4920
         Max             =   1000
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "# 1 Turbo Feed Pump (L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2835
      End
      Begin VB.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Sound Mixer 0"
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   20
         Left            =   4920
         Max             =   1000
         TabIndex        =   112
         Top             =   7440
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   20
         Left            =   3120
         Max             =   5000
         TabIndex        =   111
         Top             =   7440
         Width           =   1575
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   6600
         TabIndex        =   110
         Top             =   7440
         Width           =   855
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Spare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   109
         Top             =   8220
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Spare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   108
         Top             =   7440
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   6600
         TabIndex        =   107
         Top             =   7080
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   19
         Left            =   3120
         Max             =   5000
         TabIndex        =   106
         Top             =   7080
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   19
         Left            =   4920
         Max             =   1000
         TabIndex        =   105
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Spare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   104
         Top             =   7080
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   6600
         TabIndex        =   103
         Top             =   6720
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   18
         Left            =   3120
         Max             =   5000
         TabIndex        =   102
         Top             =   6720
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   18
         Left            =   4920
         Max             =   1000
         TabIndex        =   101
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Lub Oil Bell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   100
         Top             =   6720
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6780
         TabIndex        =   99
         Top             =   8280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   0
         Left            =   3240
         Max             =   5000
         TabIndex        =   98
         Top             =   8220
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   0
         Left            =   4980
         Max             =   1000
         TabIndex        =   97
         Top             =   8220
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   17
         Left            =   4920
         Max             =   1000
         TabIndex        =   88
         Top             =   6360
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   17
         Left            =   3120
         Max             =   5000
         TabIndex        =   87
         Top             =   6360
         Width           =   1575
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   6600
         TabIndex        =   86
         Top             =   6360
         Width           =   855
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Transfer Request Bell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   85
         Top             =   6360
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Main Turbine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   84
         Top             =   5640
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   6600
         TabIndex        =   83
         Top             =   5640
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   15
         Left            =   3120
         Max             =   5000
         TabIndex        =   82
         Top             =   5640
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   15
         Left            =   4920
         Max             =   1000
         TabIndex        =   81
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Boom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   80
         Top             =   5280
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   6600
         TabIndex        =   79
         Top             =   5280
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   14
         Left            =   3120
         Max             =   5000
         TabIndex        =   78
         Top             =   5280
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   14
         Left            =   4920
         Max             =   1000
         TabIndex        =   77
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Superheater Safety Valve Blowing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   76
         Top             =   4920
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   6600
         TabIndex        =   75
         Top             =   4920
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   13
         Left            =   3120
         Max             =   5000
         TabIndex        =   74
         Top             =   4920
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   13
         Left            =   4920
         Max             =   1000
         TabIndex        =   73
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "EOT Bell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   72
         Top             =   6000
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   6600
         TabIndex        =   71
         Top             =   6000
         Width           =   855
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   16
         Left            =   3120
         Max             =   5000
         TabIndex        =   70
         Top             =   6000
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   16
         Left            =   4920
         Max             =   1000
         TabIndex        =   69
         Top             =   6000
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   11
         Left            =   4920
         Max             =   1000
         TabIndex        =   68
         Top             =   4200
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   11
         Left            =   3120
         Max             =   5000
         TabIndex        =   67
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   6600
         TabIndex        =   66
         Top             =   4200
         Width           =   855
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Small Circuit Breaker (Right)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   65
         Top             =   4200
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   10
         Left            =   4920
         Max             =   1000
         TabIndex        =   64
         Top             =   3840
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   10
         Left            =   3120
         Max             =   5000
         TabIndex        =   63
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   6600
         TabIndex        =   62
         Top             =   3840
         Width           =   855
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Main Circuit Breaker  (Right)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   61
         Top             =   3840
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   9
         Left            =   4920
         Max             =   1000
         TabIndex        =   60
         Top             =   3480
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   9
         Left            =   3120
         Max             =   5000
         TabIndex        =   59
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   6600
         TabIndex        =   58
         Top             =   3480
         Width           =   855
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Compressor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Top             =   3480
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   12
         Left            =   4920
         Max             =   1000
         TabIndex        =   56
         Top             =   4560
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   12
         Left            =   3120
         Max             =   5000
         TabIndex        =   55
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   6600
         TabIndex        =   54
         Top             =   4560
         Width           =   855
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Superheater Vent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   53
         Top             =   4560
         Width           =   2835
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   6600
         TabIndex        =   49
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6600
         TabIndex        =   48
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   6600
         TabIndex        =   43
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   6600
         TabIndex        =   42
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   6600
         TabIndex        =   41
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6600
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   39
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HASP 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   8
         Left            =   3120
         Max             =   5000
         TabIndex        =   31
         Top             =   3120
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   7
         Left            =   3120
         Max             =   5000
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   6
         Left            =   3120
         Max             =   5000
         TabIndex        =   29
         Top             =   2400
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   5
         Left            =   3120
         Max             =   5000
         TabIndex        =   28
         Top             =   2040
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   4
         Left            =   3120
         Max             =   5000
         TabIndex        =   27
         Top             =   1680
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   3
         Left            =   3120
         Max             =   5000
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   5000
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   7
         Left            =   4920
         Max             =   1000
         TabIndex        =   24
         Top             =   2760
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   6
         Left            =   4920
         Max             =   1000
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   5
         Left            =   4920
         Max             =   1000
         TabIndex        =   22
         Top             =   2040
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   4
         Left            =   4920
         Max             =   1000
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   3
         Left            =   4920
         Max             =   1000
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   2
         Left            =   4920
         Max             =   1000
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   8
         Left            =   4920
         Max             =   1000
         TabIndex        =   18
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Propshaft"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "FD Fan #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "FD Fan #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Drum Safety Valve Blowing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Pump"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "SSTG #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "SSTG #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "SSDG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   1
         Left            =   4920
         Max             =   1000
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   5000
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Frequency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5940
      Top             =   11340
   End
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   3015
      Left            =   7140
      Top             =   3120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5318
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   50
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "'Sound Test' Is Only Enabled When The Model Has Been HALTED And Sounds Are ENABLED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   10980
      Width           =   3615
   End
End
Attribute VB_Name = "SoundMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' 6/4/98
Dim ice_strings(100) As String
Sub InitializeSliders()
Dim i As Integer


' === 13/11/02 ===
'For i = 1 To 16
For i = 1 To 18
    SoundMix.scrSpeed0(i).Value = sndFreq(0, i) / 100
    SoundMix.scrVolume0(i).Value = sndVolume(0, i)
Next

For i = 1 To 2
    SoundMix.scrSpeed1(i).Value = sndFreq(1, i) / 100
    SoundMix.scrVolume1(i).Value = sndVolume(1, i)
Next


End Sub



Private Sub cmdExit_Click()
Dim i As Integer

'copy new state on enable flags then exit

    '9/12/00 For i = 1 To 8
'    For i = 1 To 16
'        If Check0(i).Value = 1 Then
'            SndEnabled(0, i) = True
'        Else
 '           SndEnabled(0, i) = False
'       End If
'    Next


'    For i = 1 To 2
'        If Check1(i).Value = 1 Then
'            SndEnabled(1, i) = True
'        Else
'            SndEnabled(1, i) = False
'       End If
'    Next

Unload Me


End Sub

Private Sub CmdLoad_Click()
Dim FileNo As Integer
Dim i As Integer
Dim MyStr As String

Call LoadDefaultSoundMixData


'FileNo = FreeFile(0)

'Open "c:\STEAM_SIM\Sounds\SoundConfig.dat" For Input As #FileNo
        
'10/12/00 For i = 1 To 8
'For i = 1 To 31
'    Input #FileNo, MyStr
'    sndFreq(0, i) = Val(MyStr)
 '   Input #FileNo, MyStr
'    sndFreq(1, i) = Val(MyStr)
''    Input #FileNo, MyStr
 '   sndVolume(0, i) = Val(MyStr)
 '   Input #FileNo, MyStr
 '   sndVolume(1, i) = Val(MyStr)
'Next

'Close #FileNo
Call InitializeSliders

Exit Sub

ErrH:
resp = MsgBox("There was an error reading C:\Steam_Sim\SoundConfig.dat" & Chr(10) & _
       "Ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
'Close #FileNo
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
Dim FileNo As Integer
Dim i As Integer


On Error GoTo ErrH
Call SaveDefaultSoundMixdata

'FileNo = FreeFile(0)

'Open "c:\STEAM_SIM\Sounds\SoundConfig.dat" For Output As #FileNo
        
'10/12/00 For i = 1 To 8
'For i = 1 To 31
'    Print #FileNo, sndFreq(0, i)
'    Print #FileNo, sndFreq(1, i)
'    Print #FileNo, sndVolume(0, i)
'    Print #FileNo, sndVolume(1, i)
'Next

''Input #FileNo, PC_name

''This_PC_number = Val(Mid$(PC_name, Len(PC_name) - 1, 2))


'Close #FileNo

Exit Sub

ErrH:
resp = MsgBox("There was an error writing to C:\Steam_Sim\SoundConfig.dat" & Chr(10) & _
       "Ensure this file pathname exists", vbCritical + vbOKOnly, "Initialization Error")
'Close #FileNo
      

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim file_num As Integer
' ===== 19/11/97 =========================
Dim ii As Integer

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

Sounds_Enabled = True



' === 20/12/01 ===

' === 31/05/02 ===
'If TARGET_this_sim = MARYLAND Then
If TARGET_this_sim = MARYLAND Or TARGET_this_sim = TEXAS Then
    Frame2.Visible = False
    Option1(0).Caption = ""
    Option1(1).Caption = ""
End If

If PTT Then
    ' Leave as is
Else
    Check0(16).Caption = "Spare"
    Check0(17).Caption = "Spare"
    Check0(18).Caption = "Spare"
    For i = 16 To 20
        Check0(i).Enabled = False
    Next i
End If


'29/11/99 load current settings for Sound Enable Flags
'10/12/00 For i = 1 To 8

' === 28/11/01 ===
'For i = 1 To 16

' === 13/11/02 ===
'For i = 1 To 17
For i = 1 To 18
    If SndEnabled(0, i) Then
        Check0(i) = 1
    Else
        Check0(i) = 0
    End If
Next

' === 28/11/01 ===
'For i = 1 To 2
For i = 1 To 4
    If SndEnabled(1, i) Then
        Check1(i) = 1
    Else
        Check1(i) = 0
    End If
Next
'11/12/00
Call LoadDefaultSoundMixData
Call InitializeSliders

   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    If DummyPCNumber = 0 Then
       ' === 13/11/02 ===
       'For i = 1 To 16
        For i = 1 To 18
            sndVolume(0, i) = scrVolume0(i).Value
        Next
    ElseIf DummyPCNumber = 1 Then
        For i = 1 To 2
            sndVolume(1, i) = scrVolume1(i).Value
        Next
    End If
    ' If not running, then as we close we also turn off the test sounds
    If Not d_run Then

       ' === 13/11/02 ===
       'For i = 1 To 16
        For i = 1 To 18
            frmHiddenSounds.snd.DSChannel = i
            frmHiddenSounds.snd.DSCommand = 4
        Next i
    End If
    
   '14/12/00
   Sound_test_on = False
    
End Sub



Private Sub Option1_Click(Index As Integer)
Dim i As Integer

If Index = 0 Then
    Option1(1).Value = False
Else
    Option1(0).Value = False
End If

Debug.Print "option1(0).value = " & Option1(0).Value
Debug.Print "option1(1).value = " & Option1(1).Value


'turn off sounds & reset
If DummyPCNumber = 0 Then

   ' === 13/11/02 ===
   'For i = 1 To 16
    For i = 1 To 18
        frmHiddenSounds.snd.DSChannel = i
        frmHiddenSounds.snd.DSCommand = 4  'Stop Ch
        frmHiddenSounds.snd.DSCommand = 5  'Close Ch & Remove from Memory ready for different wav
    Next i
End If

If DummyPCNumber = 1 Then
    For i = 1 To 2
        frmHiddenSounds.snd.DSChannel = i
        frmHiddenSounds.snd.DSCommand = 4  'Stop Ch
        frmHiddenSounds.snd.DSCommand = 5  'Close Ch & Remove from Memory ready for different wav
    Next i
End If
frmHiddenSounds.snd.DSCommand = 6 'close direct sounds
DummyPCNumber = Index
Call frmHiddenSounds.InitialiseSounds

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim eng_set_point As Single
Dim accn As Single
Dim eng_speed As Single

On Error GoTo ErrH


If Sounds_Enabled Then
         Sound_test_on = True
        
        'eng_set_point = hscrEng(0).Value
        'accn = (eng_set_point - eng_speed) * 0.5
        'eng_speed = eng_speed + accn * (Timer1.Interval * 0.001)
        
        
        Label5.Caption = DummyPCNumber
        
        
        If DummyPCNumber = 0 Then 'do for PC 0
        
        
            '10/12/00 For i = 1 To 8
            
            ' === 28/11/01 ===
            'For i = 1 To 16
            
            ' === 13/11/02 ===
            ' For i = 1 To 17
             For i = 1 To 18
               
                sndFreq(0, i) = scrSpeed0(i).Value * 100
                FreqVal0(i) = sndFreq(0, i)
                '11/12/00
                sndVolume(0, i) = scrVolume0(i).Value
                
                
                frmHiddenSounds.snd.DSChannel = i
                If i = 1 Then
                i = i
                End If '
                frmHiddenSounds.snd.DSFrequency = sndFreq(0, i)
                '11/12/00
                frmHiddenSounds.snd.DSVolume = -sndVolume(0, i)

                If Check0(i) Then
                    If frmHiddenSounds.snd.DSStatus <> 1 Then
                        frmHiddenSounds.snd.DSCommand = 3       ' Play the sound
                    End If
                Else
                    frmHiddenSounds.snd.DSCommand = 4           ' Stop the sound
                End If
            Next
            
        ElseIf DummyPCNumber = 1 Then
        
           ' === 28/11/01 ===
           'For i = 1 To 2
            For i = 1 To 4
            
                'If SndEnabled(1, i) Then
                    'Check1(i).Value = 1
                'Else
                    'Check1(i).Value = 0
               ' End If
                sndFreq(1, i) = scrSpeed1(i).Value * 100
                FreqVal1(i) = sndFreq(1, i)
                '11/12/00
                sndVolume(1, i) = scrVolume1(i).Value
                'frmHiddenSounds.snd.DSVolume = -sndVolume(1, i)
                frmHiddenSounds.snd.DSChannel = i        '
                frmHiddenSounds.snd.DSFrequency = sndFreq(1, i)
                '11/12/00
                frmHiddenSounds.snd.DSVolume = -sndVolume(1, i)

                If Check1(i) Then
                    If frmHiddenSounds.snd.DSStatus <> 1 Then
                        frmHiddenSounds.snd.DSCommand = 3       ' Play the sound
                    End If
                Else
                    frmHiddenSounds.snd.DSCommand = 4           ' Stop the sound
                End If
            Next
        
        
        End If
        

        
End If

Exit Sub


ErrH:



End Sub


