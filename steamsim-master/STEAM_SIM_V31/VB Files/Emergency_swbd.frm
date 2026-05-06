VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form Emergency_swbd 
   BackColor       =   &H00C0C0C0&
   Caption         =   "EMERGENCY SWITCHBOARD"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   8550
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   10821
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   94
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   93
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   92
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   91
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   10
         Left            =   5760
         TabIndex        =   90
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   11
         Left            =   6960
         TabIndex        =   89
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   12
         Left            =   960
         TabIndex        =   88
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   87
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   86
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   85
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   16
         Left            =   5760
         TabIndex        =   84
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   17
         Left            =   6960
         TabIndex        =   83
         Top             =   4920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   82
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   81
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   80
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   79
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   78
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Trip_IT 
         Caption         =   "T"
         Height          =   255
         Index           =   5
         Left            =   6960
         TabIndex        =   77
         Top             =   1200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   4920
         TabIndex        =   75
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   540
         TabIndex        =   73
         Top             =   120
         Width           =   315
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   5940
         Top             =   3720
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FO TRNFR PUMP # 1 "
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   3
         Left            =   3960
         TabIndex        =   11
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label DEMO_label 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "WARNING: DEMO VERSION  (ONLY SW Pump 1  Bilge  and D/A 1 pre-lub breakers operational)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   1980
         Width           =   9975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Switchboard # 2"
         Height          =   195
         Left            =   5400
         TabIndex        =   74
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Switchboard # 1"
         Height          =   195
         Left            =   1080
         TabIndex        =   72
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   17
         Left            =   6960
         TabIndex        =   71
         Top             =   5220
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   5760
         TabIndex        =   70
         Top             =   5220
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   4560
         TabIndex        =   69
         Top             =   5220
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   3360
         TabIndex        =   68
         Top             =   5220
         Width           =   270
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   17
         Left            =   6420
         TabIndex        =   67
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   5220
         TabIndex        =   66
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   4020
         TabIndex        =   65
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   2820
         TabIndex        =   64
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   17
         Left            =   6420
         TabIndex        =   63
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   5220
         TabIndex        =   62
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   4020
         TabIndex        =   61
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   2820
         TabIndex        =   60
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAV. AIDS"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   17
         Left            =   6300
         TabIndex        =   59
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EMERG LIGHTING"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   16
         Left            =   5100
         TabIndex        =   58
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DE HUMIDIFIER"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   15
         Left            =   3900
         TabIndex        =   57
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STEERING GEAR #1"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   14
         Left            =   2700
         TabIndex        =   56
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   5535
         Left            =   240
         Top             =   480
         Width           =   7485
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   2160
         TabIndex        =   55
         Top             =   5220
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   960
         TabIndex        =   54
         Top             =   5220
         Width           =   270
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   1620
         TabIndex        =   53
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   420
         TabIndex        =   52
         Top             =   5220
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   1620
         TabIndex        =   51
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   420
         TabIndex        =   50
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   6960
         TabIndex        =   49
         Top             =   3360
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   5760
         TabIndex        =   48
         Top             =   3360
         Width           =   270
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   6420
         TabIndex        =   47
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   5220
         TabIndex        =   46
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   6420
         TabIndex        =   45
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   5220
         TabIndex        =   44
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BILGE PUMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   13
         Left            =   1500
         TabIndex        =   43
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FIRE PUMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   12
         Left            =   300
         TabIndex        =   42
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DO SERVICE PUMP"
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   11
         Left            =   6360
         TabIndex        =   41
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EMERGENCY FEED PUMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   10
         Left            =   5100
         TabIndex        =   40
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   4560
         TabIndex        =   39
         Top             =   3360
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   3360
         TabIndex        =   38
         Top             =   3360
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   2160
         TabIndex        =   37
         Top             =   3360
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   960
         TabIndex        =   36
         Top             =   3360
         Width           =   270
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   4020
         TabIndex        =   35
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   2820
         TabIndex        =   34
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1620
         TabIndex        =   33
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   4020
         TabIndex        =   32
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   2820
         TabIndex        =   31
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1620
         TabIndex        =   30
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   29
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   28
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "#1 COND'ST PUMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   9
         Left            =   3900
         TabIndex        =   27
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SSDG WDG. HTR"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   8
         Left            =   2700
         TabIndex        =   26
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRE LO PUMP SSDG"
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   7
         Left            =   1500
         TabIndex        =   25
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   LO PUMP    # 1"
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   6
         Left            =   300
         TabIndex        =   24
         Top             =   2520
         Width           =   1020
         WordWrap        =   -1  'True
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SW SERV. PUMP # 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   300
         TabIndex        =   23
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   22
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   21
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   975
         TabIndex        =   20
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1620
         TabIndex        =   19
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1620
         TabIndex        =   18
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   17
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONTROL AIR COMP"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   1
         Left            =   1500
         TabIndex        =   16
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   2820
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FO SERV. PUMP # 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   2760
         TabIndex        =   14
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   2820
         TabIndex        =   13
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   3360
         TabIndex        =   12
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FD FAN # 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   4
         Left            =   5100
         TabIndex        =   10
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "#1 FEED PUMP LO"
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   5
         Left            =   6300
         TabIndex        =   9
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   4020
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   5220
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label LIVE 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   6420
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   4020
         TabIndex        =   5
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   5220
         TabIndex        =   4
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label OFF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   6420
         TabIndex        =   3
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   4560
         TabIndex        =   2
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   5760
         TabIndex        =   1
         Top             =   1500
         Width           =   270
      End
      Begin VB.Label TRIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   6960
         TabIndex        =   0
         Top             =   1500
         Width           =   270
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3660
      TabIndex        =   95
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Emergency_swbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer




Private Sub Command1_Click()

' 25/1/99
' 12/12/99
'sy205_LSS.Show
main_SB_1.Show

End Sub

Private Sub Command2_Click()

' 25/1/99
' 12/11/99
'sy206_LSS.Show
main_SB_2.Show

End Sub

Private Sub Form_Load()
Dim i As Integer

On Error Resume Next
Call JYAScreenInit(Me, idsy207)


' 15/10/98
DEMO_label.Visible = DEMO_flag


' === 13/11/01 ===
If I_AM_HOST Then
    For i = 0 To 17
      Trip_IT(i).Visible = True
    Next i
End If
    
 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)
   
End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy207)
    HavenStretch1.Visible = True
End If
End Sub


Private Sub LIVE_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub

' === 08/01/03 ADDED: ===
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
' =======================

' 15/10/98
' 5/3/99...
'If DEMO_flag Then
'   Exit Sub
'End If

' 15/10/98
If DEMO_flag Then
  If Index = 4 Or Index = 10 Or Index = 12 Then
     Call setPCDI(BKR(2, Index + 49))
  End If
  Exit Sub
End If
   
'   Note:  BKR(2=IN,nBKR)
'  Call setPCDI(BKR(2, Index + 1))
   Call setPCDI(BKR(2, Index + 49))

End Sub

Private Sub OFF_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' === 08/01/03 ADDED: ===
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
' =======================
   
   
'   Note:  BKR(1=OUT,nBKR)
'   Call setPCDI(BKR(1, Index + 1))
    Call setPCDI(BKR(1, Index + 49))

End Sub

Private Sub Timer1_Timer()
Dim i As Integer

Dim single_char As String * 1



' === 13/11/01 ===
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' BREAKER STATUS
For i = 0 To 17

single_char = Mid$(PC_INT_STRING, 482 + i, 1)

' 5/6/98
'  If bkr_status(i + 1) = 2 Then    ' TRIPPED
   If single_char = "2" Then
       LIVE(i).BackColor = &H4000&       ' DARK GREEN
       OFF(i).BackColor = &HFFFFFF       ' WHITE
       TRIP(i).BackColor = &HFF&         ' RED
'  ElseIf bkr_status(i + 1) = 1 Then  ' BREAKER IN
   ElseIf single_char = "1" Then
       LIVE(i).BackColor = &HFF00&       ' GREEN
       OFF(i).BackColor = &H404040       ' DARK GREY
       TRIP(i).BackColor = &H40&         ' DARK RED
   Else                              ' BREAKER OUT
       LIVE(i).BackColor = &H4000&       ' DARK GREEN
       OFF(i).BackColor = &HFFFFFF       ' WHITE
       TRIP(i).BackColor = &H40&         ' DARK RED
   End If
Next i

End Sub

Private Sub TRIP_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' === 08/01/03 ADDED: ===
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
' =======================
  
   
   
'   Note:  BKR(3=RESET,nBKR)
'  Call setPCDI(BKR(3, Index + 1))
   Call setPCDI(BKR(3, Index + 49))

End Sub

Private Sub Trip_IT_Click(Index As Integer)

' === 13/11/01 ===
status_flag = True
Call set_TRIP_BREAKER(Index + 49, status_flag)
set_dummy_focus = True

End Sub
