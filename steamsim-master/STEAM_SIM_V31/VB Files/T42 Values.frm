VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_t42_values 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T42 Status"
   ClientHeight    =   9210
   ClientLeft      =   345
   ClientTop       =   1725
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9210
   ScaleWidth      =   14445
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   11520
      TabIndex        =   117
      Text            =   "Text1(5)"
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   11520
      TabIndex        =   116
      Text            =   "Text1(4)"
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Clear DI changes"
      Height          =   375
      Left            =   11400
      TabIndex        =   115
      Top             =   6480
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   5130
      Left            =   9960
      TabIndex        =   114
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command36 
      Caption         =   "DI TEST  "
      Height          =   375
      Left            =   9960
      TabIndex        =   113
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command35 
      Caption         =   "AI TEST  "
      Height          =   375
      Left            =   9960
      TabIndex        =   112
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Set I/O Delay"
      Height          =   375
      Left            =   9960
      TabIndex        =   111
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command33 
      Caption         =   "DO TEST "
      Height          =   375
      Left            =   9960
      TabIndex        =   110
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command30 
      Caption         =   "AO TEST "
      Height          =   375
      Left            =   9960
      TabIndex        =   109
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Re-load Snap 2"
      Height          =   375
      Left            =   5760
      TabIndex        =   108
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Take Snap 2"
      Height          =   375
      Left            =   5760
      TabIndex        =   107
      Top             =   7800
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   7560
      TabIndex        =   106
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Start Model      (SPLIT Mode)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   105
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox IoTime 
      Height          =   285
      Left            =   5280
      TabIndex        =   104
      Text            =   "IOTime"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   7800
      TabIndex        =   103
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   6960
      TabIndex        =   102
      Text            =   "Text1(3)"
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton EndTimer 
      Caption         =   "EndTimer"
      Height          =   375
      Left            =   7200
      TabIndex        =   101
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7200
      TabIndex        =   100
      Text            =   "IO Count"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton CallDelay 
      Caption         =   "Start IO"
      Height          =   375
      Left            =   7200
      TabIndex        =   99
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command26 
      Caption         =   "STOP M2"
      Height          =   255
      Left            =   8400
      TabIndex        =   98
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      Caption         =   "START M2"
      Height          =   255
      Left            =   8400
      TabIndex        =   97
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "EDC loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   92
      Top             =   3720
      Width           =   2415
      Begin VB.Label Label1 
         Caption         =   "SCP Loading %"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   96
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "PCP Loading %"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   94
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   93
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Clear ALL Faults (CA)"
      Height          =   735
      Left            =   6960
      TabIndex        =   91
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Enter Fault (Exxx)"
      Height          =   735
      Left            =   6960
      TabIndex        =   90
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Synch Pegs M1/M2"
      Height          =   735
      Left            =   8040
      TabIndex        =   87
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Open M2 Bkr"
      Height          =   375
      Left            =   6480
      TabIndex        =   67
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Close M2 Bkr"
      Height          =   375
      Left            =   6480
      TabIndex        =   66
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Open X2 Bkr"
      Height          =   375
      Left            =   4920
      TabIndex        =   65
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Close X2 Bkr"
      Height          =   375
      Left            =   4920
      TabIndex        =   64
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Resume Exercise"
      Height          =   375
      Left            =   4320
      TabIndex        =   63
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Re-load Snap 1"
      Height          =   375
      Left            =   4320
      TabIndex        =   62
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Take Snap 1"
      Height          =   375
      Left            =   4320
      TabIndex        =   61
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   60
      Text            =   "Text1(2)"
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Open X1 Bkr"
      Height          =   375
      Left            =   1800
      TabIndex        =   59
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Close X1 Bkr"
      Height          =   375
      Left            =   1800
      TabIndex        =   58
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "START M1"
      Height          =   255
      Left            =   7440
      TabIndex        =   57
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "STOP M1"
      Height          =   255
      Left            =   7440
      TabIndex        =   56
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "END DLL TASKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      TabIndex        =   55
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Open L1 Bkr"
      Height          =   375
      Left            =   3360
      TabIndex        =   54
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Open M1 SACE Bkr"
      Height          =   375
      Left            =   0
      TabIndex        =   53
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Reset M1 OVPU"
      Height          =   375
      Left            =   4320
      TabIndex        =   52
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Close L1 Bkr"
      Height          =   375
      Left            =   3360
      TabIndex        =   51
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start Model      (Normal Mode)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   50
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close M1 SACE Bkr"
      Height          =   375
      Left            =   0
      TabIndex        =   49
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   48
      Text            =   "Text1(1)"
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   44
      Text            =   "Text1(0)"
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "STOP"
      Height          =   255
      Left            =   960
      TabIndex        =   42
      Top             =   5880
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   6000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lower"
      Height          =   255
      Left            =   1680
      TabIndex        =   41
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Raise"
      Height          =   255
      Left            =   1680
      TabIndex        =   40
      Top             =   5760
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shore Supplies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   20
      Top             =   240
      Width           =   2415
      Begin VB.Label Label1 
         Caption         =   "PCP S/S aft   volts"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "PCP S/S ford volts"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SCP S/S aft   volts"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SCP S/S ford volts"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " PCP S/S aft    Hz"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " PCP S/S ford  Hz"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " SCP S/S aft    Hz"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " SCP S/S ford  Hz"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   29
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   27
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   26
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   25
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   23
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   22
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "   Alternator Parameters "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   2520
      TabIndex        =   10
      Top             =   2040
      Width           =   4335
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   86
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   85
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   84
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   83
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   82
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   81
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   80
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   79
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   78
         Top             =   960
         Width           =   735
      End
      Begin VB.Label M2label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M2label"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   77
         Top             =   600
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   76
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   75
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   74
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   73
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   72
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   71
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   70
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   69
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   68
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Governor %"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   39
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Frequency Hz"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "   M1         M2    "
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
         Left            =   0
         TabIndex        =   37
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "VG        volts"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "A          amps"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "P           kW"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Q           kVar"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "VA         kVA"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cosphi"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tan Delta"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cos Delta"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label M1label 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M1label"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bus-Bar Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   2535
      Begin VB.Label Label1 
         Caption         =   "P Nom Bus"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Q Nom Bus"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "V volts"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "F Hz"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   8760
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DTREnable       =   -1  'True
      DataBits        =   7
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   495
      Left            =   8280
      TabIndex        =   88
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      _Version        =   327680
      MouseIcon       =   "T42 Values.frx":0000
      SelStart        =   5
      Value           =   5
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   495
      Left            =   1200
      TabIndex        =   46
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      _Version        =   327680
      MouseIcon       =   "T42 Values.frx":001C
      SelStart        =   5
      Value           =   5
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   1200
      TabIndex        =   43
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      _Version        =   327680
      MouseIcon       =   "T42 Values.frx":0038
      SelStart        =   5
      Value           =   5
   End
   Begin VB.Label Label6 
      Caption         =   "PCP Phase error"
      Height          =   375
      Left            =   8280
      TabIndex        =   89
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "   M1 SCP AVR"
      Height          =   375
      Left            =   0
      TabIndex        =   47
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "   M1 PCP AVR"
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   7680
      Width           =   1215
   End
End
Attribute VB_Name = "frm_t42_values"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'C           Bus-Bar Parameters :
'       EQUIVALENCE (T42BUSR(1),PNOMBUS(1)),(T42BUSR(13),QNOMBUS(1))
'     +          ,(T42BUSR(25),V(1)),(T42BUSR(41),F(1))
'C           Alternator Parameters which are a function of bus-loading :
'     +          ,(T42BUSR(57),VG(1)),(T42BUSR(69),A(1))
'     +          ,(T42BUSR(81),P(1)),(T42BUSR(93),Q(1)),(T42BUSR(105),VA(1))
'     +          ,(T42BUSR(117),COSPHI(1))
'     +          ,(T42BUSR(129),TANDELTA(1)),(T42BUSR(141),COSDELTA(1))

Sub show_model_vals()
Dim i As Integer

'For i = 0 To 9
For i = 0 To 7
    Label2(i).Caption = Format$(EXECMODELR(i + 1), "###0.00")
Next i
For i = 8 To 9
    Label2(i).Caption = Format$(EXECMODELR(i + 1) * 100#, "###0")     ' PCP and SCP loading in %
Next i

Label2(10).Caption = Format$(T42BUSR(1) * 0.001, "###0.000")       ' Pnombus
Label2(11).Caption = Format$(T42BUSR(13) * 0.001, "###0.000")      ' Qnombus

Label2(12).Caption = Format$(T42BUSR(25), "###0")       ' Dusbar volts
Label2(13).Caption = Format$(T42BUSR(41), "##0.00")    'Busbar Freq

M1label(0).Caption = Format$(T42BUSR(57), "###0")       ' VG(1)
M1label(1).Caption = Format$(T42BUSR(69), "####0")      ' Amps
M1label(2).Caption = Format$(T42BUSR(81) * 0.001, "###0.000") ' Watts
M1label(3).Caption = Format$(T42BUSR(93) * 0.001, "###0.000") ' Var
M1label(4).Caption = Format$(T42BUSR(105) * 0.001, "###0.000") ' VA
M1label(5).Caption = Format$(T42BUSR(117), "#0.00")      ' cos phi
M1label(6).Caption = Format$(T42BUSR(129), "#0.00")     ' tan delta
M1label(7).Caption = Format$(T42BUSR(141), "#0.00")    ' cos delta
M1label(8).Caption = Format$(T42DIESELR(121), "##0.00")  ' FG(1)
M1label(9).Caption = Format$(T42DIESELR(109) * 100#, "##0.0")  ' GL(1)

M2label(0).Caption = Format$(T42BUSR(59), "###0")       ' VG(3)
M2label(1).Caption = Format$(T42BUSR(71), "####0")      ' Amps
M2label(2).Caption = Format$(T42BUSR(83) * 0.001, "###0.000") ' Watts
M2label(3).Caption = Format$(T42BUSR(95) * 0.001, "###0.000") ' Var
M2label(4).Caption = Format$(T42BUSR(107) * 0.001, "###0.000") ' VA
M2label(5).Caption = Format$(T42BUSR(119), "#0.00")      ' cos phi
M2label(6).Caption = Format$(T42BUSR(131), "#0.00")     ' tan delta
M2label(7).Caption = Format$(T42BUSR(143), "#0.00")    ' cos delta
M2label(8).Caption = Format$(T42DIESELR(123), "##0.00")  ' FG(3)
M2label(9).Caption = Format$(T42DIESELR(111) * 100#, "##0.0")  ' GL(3)

End Sub



Private Sub CallDelay_Click()
Dim xx As Integer


'20/3/97 kick off io
'enable timer to allow normal io running
IOCounter = 0
frm_host.IO_timer.Enabled = True



End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub


Private Sub Command1_Click()

' 23/1/97  Raise M1 frequency
    DI(338) = True
    DI(344) = False
'  Call VB_EXEC(DI(0))
   AI_chan_no = -1  ' Do not set an AI
' 20/3/97
'  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
           
End Sub

Private Sub Command10_Click()
       
' End the FORTRAN tasks
         task_array(1) = False
           task_array(2) = True
         Call Init_Tasks(task_array(1))
           task_array(2) = False
       
End Sub

Private Sub Command11_Click()

' 31/1/97 Stop M1
             Igenstart = 1
        start_gen_flag = False
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
 
End Sub

Private Sub Command12_Click()

' 31/1/97 Start M1
             Igenstart = 1
        start_gen_flag = True
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)

End Sub


Private Sub Command13_Click()

' 31/1/97  Close X1 breaker
' 5/2/97
'   DI(143) = True
    DI(271) = True
    DI(270) = False
   AI_chan_no = -1  ' Do not set an AI
' 20/3/97
'  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
    DI(271) = False
    DI(270) = False
   
End Sub

Private Sub Command14_Click()

' Open X1 Breaker
    DI(271) = False
    DI(270) = True
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
    DI(270) = False
   
End Sub

Private Sub Command15_Click()
Dim Isnap_number As Integer

' 7/2/97
Isnap_number = 1
Call TPSNAP(Isnap_number)

End Sub

Private Sub Command16_Click()
Dim Isnap_number As Integer

' 7/2/97
Isnap_number = 1
Call TPRELSNAP(Isnap_number)


End Sub

Private Sub Command17_Click()

               ICODE = 1
        poke_exec(5) = True   ' To Resume running the Type 42MODEL
        Call TPRUN(ICODE, poke_exec(1))
' 7/2/97
        poke_exec(5) = False

End Sub

Private Sub Command18_Click()

' Close X(2)
    DI(243) = True
    DI(235) = False
   AI_chan_no = -1  ' Do not set an AI
   
'  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
    DI(243) = False
   
End Sub

Private Sub Command19_Click()

' Open X(2)
    DI(235) = True
    DI(243) = False
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
    DI(235) = False

End Sub


Private Sub Command2_Click()

' 23/12/96 Lower M1 frequency
'  VB_DI(344) = True

' 23/1/97  Lower M1 frequency
    DI(338) = False
    DI(344) = True
'  Call VB_EXEC(DI(0))
    AI_chan_no = -1  ' Do not set an AI
    
'    Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
    
  
End Sub

Private Sub Command20_Click()

' Close M2 breaker
    DI(216) = True
    DI(203) = False
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
    DI(216) = False
   
End Sub

Private Sub Command21_Click()

' Open M2 breaker
    DI(216) = False
    DI(203) = True
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))

    DI(203) = False
   
End Sub


Private Sub Command22_Click()

' Insert Synchronising pegs:
    DI(329) = True      ' PPEG1C
    DI(300) = True      ' PPEG3A
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   

End Sub

Private Sub Command23_Click()

' 14/2/97   Enter a Fault

    Ifreq = 0
   Ivolts = 0
   
try_again:
   Ifault = InputBox("Enter fault number", "Fault Entry")
   If Ifault < 1 Or Ifault > 999 Then GoTo try_again:
   
' SPECIAL Faults  ( see the variable SPECIALFAULT in TPFC1.F90 )
    If Ifault = 11 Or Ifault = 17 Then          ' 11 for PCP , 17 for SCP(SPLIT mode)
again11:       Ifreq = InputBox("Enter S/S aft frequency [ 55-65 Hz ]", "Shore Supply")
            If Ifreq < 55 Or Ifreq > 65 Then GoTo again11:
    ElseIf Ifault = 14 Or Ifault = 20 Then      ' 14 for PCP , 20 for SCP(SPLIT mode)
again14:       Ifreq = InputBox("Enter S/S ford frequency [ 55-65 Hz ]", "Shore Supply")
            If Ifreq < 55 Or Ifreq > 65 Then GoTo again14:
    ElseIf Ifault = 71 Or Ifault = 77 Then      ' 71 for PCP , 77 for SCP(SPLIT mode)
again71:       Ivolts = InputBox("Enter S/S aft voltage [ 410-475 volts ]", "Shore Supply")
            If Ivolts < 410 Or Ivolts > 475 Then GoTo again71:
     ElseIf Ifault = 74 Or Ifault = 80 Then      ' 74 for PCP , 80 for SCP(SPLIT mode)
again77:       Ivolts = InputBox("Enter S/S ford voltage [ 410-475 volts ]", "Shore Supply")
            If Ivolts < 410 Or Ivolts > 475 Then GoTo again77:
    End If
    
   
   VDU_TYPE = 0
   
   poke_exec(10) = True        ' To ENTER a fault i.e. Exxx in TPFAULT.F90
   Call TPFAULT(Ifault, VDU_TYPE, Ifreq, Ivolts, poke_exec(1))
   poke_exec(10) = False

End Sub

Private Sub Command24_Click()

' 14/2/97 Clear ALL Faults
   VDU_TYPE = 0
   
   poke_exec(9) = True        ' To ENTER a fault i.e. Exxx in TPFAULT.F90
   Call TPFAULT(Ifault, VDU_TYPE, Ifreq, Ivolts, poke_exec(1))
   poke_exec(9) = False

End Sub

Private Sub Command25_Click()

' 28/2/97 Start M2
             Igenstart = 3
        start_gen_flag = True
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)

End Sub

Private Sub Command26_Click()

' 28/2/97 Stop M2
             Igenstart = 3
        start_gen_flag = False
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
       
End Sub

Private Sub Command27_Click()

' =======  23/1/97 =============================
       poke_exec(1) = True   ' To CLEAR the LOGO  (YOU MUST DO THIS or you will HANG in the DLL)
           poke_exec(3) = True   ' To select T42 SPLIT mode
'          poke_exec(2) = True   ' To select T42 NORMAL mode
       poke_vb(1) = False    ' Goes true at the exit of INITLZ.F90
       Call INITLZ(poke_exec(1), poke_vb(1))
          poke_exec(2) = False
          poke_exec(3) = False

' 28/2/97
'       Call Sleep(2000)
       
         task_array(1) = True
         task_array(2) = False
         Call Init_Tasks(task_array(1))
         
' 28/2/97      Call Sleep(2000)
         
' 24/1/97 Don't need         Call TPMMU
  
      Instructor_name = "RWJ"
          Course_name = "Course 1"
           Class_name = "Class 2"
           Call TPEXEC(Instructor_name, Class_name, Course_name)
'          Call TPEXEC(Rx_strings(2), Rx_strings(3), Rx_strings(4))

' 28/2/97       Call Sleep(2000)
  
               ICODE = 1
        poke_exec(4) = True   ' To RUN the Type 42MODEL
        Call TPRUN(ICODE, poke_exec(1))
' 7/2/97
        poke_exec(4) = False
        
' 28/2/97      Call Sleep(2000)
  
          Iovpureset = 0
              Igenstart = 1
         start_gen_flag = True
           Ipcp_loading = 80
'          Iovpureset = Val(Rx_strings(2))
'          Igenstart = Val(Rx_strings(3))
'          If Rx_strings(4) = "START" Then
'              start_gen_flag = True
'          Else
'              start_gen_flag = False
'          End If
'          Ipcp_loading = Val(Rx_strings(5))
          
'        Debug.Print "Iovpureset = " & Iovpureset
'        Debug.Print "Igenstart = " & Igenstart
'        If start_gen_flag Then
'           Debug.Print "start_gen_flag =   TRUE"
'        Else
'           Debug.Print "start_gen_flag =   FALSE"
'        End If
'
'        Debug.Print "Ipcp_loading = " & Ipcp_loading
        
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
       
              Igenstart = 3
' 31/1/97 Start M2
' 12/2/97       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
       
End Sub

Private Sub Command28_Click()
Dim Isnap_number As Integer

' 9/4/97
Isnap_number = 2
Call TPSNAP(Isnap_number)

End Sub

Private Sub Command29_Click()
Dim Isnap_number As Integer

'9/4/97
Isnap_number = 2
Call TPRELSNAP(Isnap_number)

End Sub

Private Sub Command3_Click()

' 23/1/97  Stop Raise/Lower M1 frequency
    DI(338) = False
    DI(344) = False
'  Call VB_EXEC(DI(0))
    AI_chan_no = -1  ' Do not set an AI
    
 '   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
    Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
  
End Sub

Private Sub Command30_Click()

' 17/4/97
itest_type = 2  ' For AO test

itest_no = 0    ' Meters Zero
itest_no = 1    ' Meters 1/2 scale
itest_no = 2    ' Meters FSD

ichannel = 0
ivalue = 0

tryagain:
   itest_no = Val(InputBox("Enter:" & Chr$(13) & "0 for Meter Zero TEST" & Chr$(13) & "1 for Meter 1/2 FSD Test" & Chr$(13) & "2 for Meter FSD Test", " ** AO TEST **"))
   If itest_no < 0 Or itest_no > 2 Then GoTo tryagain:

Call TPTEST(itest_type, itest_no, ichannel, ivalue)

End Sub

Private Sub Command31_Click()



End Sub


Private Sub Command32_Click()

' 21/4/97
   nch = Val(InputBox("Enter:" & Chr$(13) & "DI changed ", " ** Clear DIchanges **"))
   Call RsetDIch(nch)

End Sub

Private Sub Command33_Click()

' 18/4/97    *** DO TEST ***

itest_type = 4  ' For DO test

itest_no = 0    ' Breaker test (HMS Sultan)
itest_no = 1    ' Lamp test    (HMS Sultan)
itest_no = 2    ' Restore Breakers and Lamps (after Breaker or Lamp Test)
itest_no = 3    ' Individual DO test (HMS Sultan)
itest_no = 4    ' Individual DO test (General)

ichannel = 0
ivalue = 0


again1:
   itest_no = Val(InputBox("Enter:" & Chr$(13) & "0 for Breaker Test (SULTAN)" & Chr$(13) _
   & "1 for a Lamp Test (SULTAN) " & Chr$(13) & "2 to RESTORE breakers & Lamps (SULTAN)" & Chr$(13) _
   & "3 Individual DO (SULTAN)" & Chr$(13) & "4 Individual DO (GENERAL)", " ** DO TEST **"))
   
   If itest_no < 0 Or itest_no > 4 Then GoTo again1:
   
again2:
   ichannel = Val(InputBox("Enter:" & Chr$(13) & "Channel no: ", " ** DO TEST **"))
   If ichannel < 0 Or ichannel > 5119 Then GoTo again2:

Call TPTEST(itest_type, itest_no, ichannel, ivalue)

End Sub

Private Sub Command34_Click()

DD:
   Dcount = Val(InputBox("Enter:" & Chr$(13) & "Delay loop count ", " ** SET DelayLoopCount **"))
   If Dcount < 0 Or Dcount > 30000 Then GoTo DD:

Call SETDELAY(Dcount)

End Sub


Private Sub Command35_Click()

' 18/4/97    *** AI TEST ***

itest_type = 1  ' For AI test

itest_no = 0   ' AI Test SULTAN
itest_no = 1   ' AI Test GENERAL

ichannel = 0
ivalue = 0


again1:
   itest_no = Val(InputBox("Enter:" & Chr$(13) & "0 for SULTAN AI TEST" & Chr$(13) _
   & "1 for GENERAL AI TEST " & Chr$(13), " ** AI TEST **"))
   
   If itest_no < 0 Or itest_no > 1 Then GoTo again1:

Call TPTEST(itest_type, itest_no, ichannel, ivalue)

End Sub

Private Sub Command36_Click()

itest_type = 3  ' For DI test

itest_no = 0   ' DI Test SULTAN
itest_no = 1   ' DI Test GENERAL

ichannel = 0
ivalue = 0

DI1:
   itest_no = Val(InputBox("Enter:" & Chr$(13) & "0 for SULTAN DI TEST" & Chr$(13) _
   & "1 for GENERAL DI TEST (0:48) " & Chr$(13), " ** DI TEST **"))
   
   If itest_no < 0 Or itest_no > 1 Then GoTo DI1:

Call TPTEST(itest_type, itest_no, ichannel, ivalue)

End Sub


Private Sub Command4_Click()

' 23/1/97  Close M1 breaker
    DI(712) = True
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
   
End Sub

Private Sub Command5_Click()

' =======  23/1/97 =============================
       poke_exec(1) = True   ' To CLEAR the LOGO  (YOU MUST DO THIS or you will HANG in the DLL)
'         poke_exec(3) = True   ' To select T42 SPLIT mode
          poke_exec(2) = True   ' To select T42 NORMAL mode
       poke_vb(1) = False    ' Goes true at the exit of INITLZ.F90
       Call INITLZ(poke_exec(1), poke_vb(1))
          poke_exec(2) = False
          poke_exec(3) = False
'2/4/97
' 2/4/97
'       GoTo skip

' 28/2/97
'       Call Sleep(2000)
       
'  6/5/97
'         task_array(1) = True
'         task_array(2) = False
'         Call Init_Tasks(task_array(1))
         
     GoTo skip
         
         
' 28/2/97      Call Sleep(2000)
         
' 24/1/97 Don't need         Call TPMMU
  
      Instructor_name = "RWJ"
          Course_name = "Course 1"
           Class_name = "Class 2"
           Call TPEXEC(Instructor_name, Class_name, Course_name)
'          Call TPEXEC(Rx_strings(2), Rx_strings(3), Rx_strings(4))

' 28/2/97       Call Sleep(2000)
  
               ICODE = 1
        poke_exec(4) = True   ' To RUN the Type 42MODEL
        Call TPRUN(ICODE, poke_exec(1))
' 7/2/97
        poke_exec(4) = False
        
' 28/2/97      Call Sleep(2000)
  
          Iovpureset = 0
              Igenstart = 1
         start_gen_flag = True
           Ipcp_loading = 80
'          Iovpureset = Val(Rx_strings(2))
'          Igenstart = Val(Rx_strings(3))
'          If Rx_strings(4) = "START" Then
'              start_gen_flag = True
'          Else
'              start_gen_flag = False
'          End If
'          Ipcp_loading = Val(Rx_strings(5))
          
'        Debug.Print "Iovpureset = " & Iovpureset
'        Debug.Print "Igenstart = " & Igenstart
'        If start_gen_flag Then
'           Debug.Print "start_gen_flag =   TRUE"
'        Else
'           Debug.Print "start_gen_flag =   FALSE"
'        End If
'
'        Debug.Print "Ipcp_loading = " & Ipcp_loading
        
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
       
              Igenstart = 3
' 31/1/97 Start M2
' 12/2/97       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
       
skip:
End Sub

Private Sub Command6_Click()

' 23/1/97  Close L1 breaker
    DI(714) = True
    DI(716) = False
   AI_chan_no = -1  ' Do not set an AI
   
 '  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
    Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
  
   Call Sleep(2000)
   DI(714) = False
   
End Sub

Private Sub Command7_Click()

         Iovpureset = 1
              Igenstart = 0      ' Don't Start or stop a genny
         start_gen_flag = True   ' Don't care because Igenstart = 0
'           Ipcp_loading = 80    ' Leave as previous value
                  
       Call TPINSTRP(Iovpureset, Igenstart, start_gen_flag, Ipcp_loading)
       

End Sub

Private Sub Command8_Click()

' 23/1/97  Open M1 breaker
    DI(712) = False
   AI_chan_no = -1  ' Do not set an AI
   
 '  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
   
End Sub

Private Sub Command9_Click()

' 23/1/97  Open L1 breaker
    DI(714) = False
    DI(716) = True
   AI_chan_no = -1  ' Do not set an AI
   
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   
   Call Sleep(2000)
   DI(716) = False
   
End Sub

Private Sub EndTimer_Click()
    Text2.Text = "zero"
End Sub

Private Sub Form_Load()
'frm_exec.mnu_sim_run.Checked = True

Static i As Integer

' 10/3/97
' 12/3/97 Conflicts with com1: in the FORTRAN DLL
'If Not MSComm1.PortOpen Then
'   MSComm1.PortOpen = True
'End If

' Set the status to "Running" and show this
' status in the status bar on the MDI form
Sim_Running_Status = "Running"
frm_host.StatusBar1.Panels.item(2).Text = "Simulator " & Sim_Running_Status

show_model_vals

End Sub


Private Sub Form_Unload(Cancel As Integer)
frm_host.mnu_sim_run.Checked = False
End Sub







Private Sub MSComm1_OnComm()

' 10/3/97
'Text1(3).Text = MSComm1.Input

End Sub

Private Sub Slider1_Change()

    AI_chan_no = 9        ' set  AI(M1PCPPOT)
    AI_Value = Slider1.Value * 25.5
    
    Text1(0).Text = Slider1.Value & "  " & AI_Value
    
 ' 7/4/97
       AI(AI_chan_no) = AI_Value
     
 '  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
 ' 29/4/97
 ' Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
 
End Sub

Private Sub Slider2_Change()

    AI_chan_no = 24        ' set  AI(M1SCPPOT)
    AI_Value = Slider2.Value * 25.5
    
    Text1(1).Text = Slider2.Value & "  " & AI_Value
    
 ' 7/4/97
       AI(AI_chan_no) = AI_Value
 
 '  Call VB_EXEC(DI(0), AI_chan_no, AI_Value)
 ' 29/4/97
 ' Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
  
End Sub

Private Sub Slider3_Click()

    AI_chan_no = 13        ' set  AI(PCPPHASE)
    AI_Value = Slider3.Value * 25.5
    
    Text1(1).Text = Slider3.Value & "  " & AI_Value
 ' 7/4/97
       AI(AI_chan_no) = AI_Value
     
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))
   

End Sub


Private Sub Timer1_Timer()
Dim i As Integer

' 5/2/97
If VB_DO(503) Then
   Text1(2).Text = "X1 is closed"
Else
   Text1(2).Text = "X1 is open"
End If

' 11/4/97
Text1(3).Text = "AIACTUAL(9)=" & AIACTUAL(9)

show_model_vals

' 12/3/97
List1.Clear
 For i = 1 To 20
      If fault_in(i) > 0 Then
         List1.AddItem "Fault " & fault_in(i) & " = TRUE"
    End If
 Next i
 
' 8/4/97
List2.Clear
 For i = 1 To 4
     List2.AddItem "par_alts(" & i & ") = " & par_alts(i)
 Next i
 
' 21/4/97
List3.Clear
 For i = 0 To 20
   If (VB_DIchanges(i) <> 0) Then
     List3.AddItem "VB_DIchanges(" & i & ") = " & VB_DIchanges(i)
   End If
 Next i

IoTime.Text = VB_IO_TIME

' 20/3/97
   AI_chan_no = -1  ' Do not set an AI

' 29/4/97 Should ONLY do this if VB calls IO
'   Call VB_EXEC(DI(0), AI_chan_no, AI_Value, AI(0))


' 6/5/97
If DI363 Then
   Text1(4).Text = "DI363 = true"
Else
   Text1(4).Text = "DI363 = false"
End If

' 14/5/97
Text1(5).Text = "EDCR(25),(31) = " & EDCR(25) & EDCR(31)
End Sub


