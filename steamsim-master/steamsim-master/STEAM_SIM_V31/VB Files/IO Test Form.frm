VERSION 5.00
Begin VB.Form IOtest 
   Caption         =   "IO Test"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   8565
   Begin VB.Frame Frame3 
      Caption         =   "Analog TEST"
      Height          =   1095
      Left            =   4800
      TabIndex        =   116
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command5 
         Caption         =   "Analog TEST"
         Height          =   495
         Left            =   120
         TabIndex        =   117
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Digital TEST"
      Height          =   1215
      Left            =   4800
      TabIndex        =   113
      Top             =   1200
      Width           =   3615
      Begin VB.CommandButton Command6 
         Caption         =   "Digital TEST"
         Height          =   495
         Left            =   120
         TabIndex        =   114
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "WARNING: ONLY SCAN DI Boards which actually exist"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1800
         TabIndex        =   115
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Set DI boards to DEFAULT"
      Height          =   495
      Left            =   5880
      TabIndex        =   112
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Disable ALL DI boards"
      Height          =   495
      Left            =   600
      TabIndex        =   111
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "IO"
      Height          =   2415
      Left            =   120
      TabIndex        =   102
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command11 
         Caption         =   "HALT IO"
         Height          =   375
         Left            =   240
         TabIndex        =   121
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "SynchroScope TEST"
         Height          =   495
         Left            =   2760
         TabIndex        =   120
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "STOP Bus TEST"
         Height          =   495
         Left            =   240
         TabIndex        =   118
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox bus_link_check 
         Caption         =   "Bus_link_check"
         Height          =   255
         Left            =   1920
         TabIndex        =   110
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "START Bus TEST"
         Height          =   495
         Left            =   240
         TabIndex        =   108
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox IO_check 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   107
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Run IO"
         Height          =   375
         Left            =   240
         TabIndex        =   105
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Free IO Board"
         Height          =   255
         Left            =   1560
         TabIndex        =   104
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Initialize IO"
         Height          =   255
         Left            =   360
         TabIndex        =   103
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Only Initialise ONCE"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   2520
         TabIndex        =   119
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "STATUS"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   109
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "STATUS"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   106
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   47
      Left            =   7560
      TabIndex        =   81
      Top             =   6480
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   46
      Left            =   7560
      TabIndex        =   80
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   45
      Left            =   7560
      TabIndex        =   79
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   44
      Left            =   7560
      TabIndex        =   78
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   43
      Left            =   7560
      TabIndex        =   77
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   42
      Left            =   7560
      TabIndex        =   76
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   41
      Left            =   7560
      TabIndex        =   75
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   40
      Left            =   7560
      TabIndex        =   74
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   39
      Left            =   7560
      TabIndex        =   73
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   38
      Left            =   7560
      TabIndex        =   72
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   37
      Left            =   7560
      TabIndex        =   71
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   36
      Left            =   7560
      TabIndex        =   70
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   35
      Left            =   7560
      TabIndex        =   69
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   34
      Left            =   7560
      TabIndex        =   68
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   33
      Left            =   7560
      TabIndex        =   67
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   32
      Left            =   7560
      TabIndex        =   66
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   31
      Left            =   4800
      TabIndex        =   49
      Top             =   6480
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   30
      Left            =   4800
      TabIndex        =   48
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   29
      Left            =   4800
      TabIndex        =   47
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   28
      Left            =   4800
      TabIndex        =   46
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   27
      Left            =   4800
      TabIndex        =   45
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   26
      Left            =   4800
      TabIndex        =   44
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   25
      Left            =   4800
      TabIndex        =   43
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   24
      Left            =   4800
      TabIndex        =   42
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   23
      Left            =   4800
      TabIndex        =   41
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   22
      Left            =   4800
      TabIndex        =   40
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   21
      Left            =   4800
      TabIndex        =   39
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   20
      Left            =   4800
      TabIndex        =   38
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   19
      Left            =   4800
      TabIndex        =   37
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   18
      Left            =   4800
      TabIndex        =   36
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   17
      Left            =   4800
      TabIndex        =   35
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   16
      Left            =   4800
      TabIndex        =   34
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   15
      Left            =   2280
      TabIndex        =   31
      Top             =   6480
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   14
      Left            =   2280
      TabIndex        =   30
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   13
      Left            =   2280
      TabIndex        =   29
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   12
      Left            =   2280
      TabIndex        =   28
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   11
      Left            =   2280
      TabIndex        =   27
      Top             =   5520
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   10
      Left            =   2280
      TabIndex        =   26
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   9
      Left            =   2280
      TabIndex        =   25
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   8
      Left            =   2280
      TabIndex        =   24
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   2280
      TabIndex        =   23
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   2280
      TabIndex        =   22
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   2280
      TabIndex        =   21
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   20
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   19
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   18
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   17
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox DI_board_check 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   16
      Top             =   2880
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   3120
   End
   Begin VB.Label Label1 
      Caption         =   "STATUS"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   101
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DI Board"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   100
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "STATUS"
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   99
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DI Board"
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   98
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   7F HEX"
      Height          =   255
      Index           =   47
      Left            =   5880
      TabIndex        =   97
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   7E HEX"
      Height          =   255
      Index           =   46
      Left            =   5880
      TabIndex        =   96
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   7D HEX"
      Height          =   255
      Index           =   45
      Left            =   5880
      TabIndex        =   95
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   7C HEX"
      Height          =   255
      Index           =   44
      Left            =   5880
      TabIndex        =   94
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   7B HEX"
      Height          =   255
      Index           =   43
      Left            =   5880
      TabIndex        =   93
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   7A HEX"
      Height          =   255
      Index           =   42
      Left            =   5880
      TabIndex        =   92
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   79 HEX"
      Height          =   255
      Index           =   41
      Left            =   5880
      TabIndex        =   91
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   78 HEX"
      Height          =   255
      Index           =   40
      Left            =   5880
      TabIndex        =   90
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   77 HEX"
      Height          =   255
      Index           =   39
      Left            =   5880
      TabIndex        =   89
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   76 HEX"
      Height          =   255
      Index           =   38
      Left            =   5880
      TabIndex        =   88
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   75 HEX"
      Height          =   255
      Index           =   37
      Left            =   5880
      TabIndex        =   87
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   74 HEX"
      Height          =   255
      Index           =   36
      Left            =   5880
      TabIndex        =   86
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   73 HEX"
      Height          =   255
      Index           =   35
      Left            =   5880
      TabIndex        =   85
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   72 HEX"
      Height          =   255
      Index           =   34
      Left            =   5880
      TabIndex        =   84
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   71 HEX"
      Height          =   255
      Index           =   33
      Left            =   5880
      TabIndex        =   83
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   70 HEX"
      Height          =   255
      Index           =   32
      Left            =   5880
      TabIndex        =   82
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   6F HEX"
      Height          =   255
      Index           =   31
      Left            =   3120
      TabIndex        =   65
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   6E HEX"
      Height          =   255
      Index           =   30
      Left            =   3120
      TabIndex        =   64
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   6D HEX"
      Height          =   255
      Index           =   29
      Left            =   3120
      TabIndex        =   63
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   6C HEX"
      Height          =   255
      Index           =   28
      Left            =   3120
      TabIndex        =   62
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   6B HEX"
      Height          =   255
      Index           =   27
      Left            =   3120
      TabIndex        =   61
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   6A HEX"
      Height          =   255
      Index           =   26
      Left            =   3120
      TabIndex        =   60
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   69 HEX"
      Height          =   255
      Index           =   25
      Left            =   3120
      TabIndex        =   59
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   68 HEX"
      Height          =   255
      Index           =   24
      Left            =   3120
      TabIndex        =   58
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   67 HEX"
      Height          =   255
      Index           =   23
      Left            =   3120
      TabIndex        =   57
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   66 HEX"
      Height          =   255
      Index           =   22
      Left            =   3120
      TabIndex        =   56
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   65 HEX"
      Height          =   255
      Index           =   21
      Left            =   3120
      TabIndex        =   55
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   64 HEX"
      Height          =   255
      Index           =   20
      Left            =   3120
      TabIndex        =   54
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   63 HEX"
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   53
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   62 HEX"
      Height          =   255
      Index           =   18
      Left            =   3120
      TabIndex        =   52
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   61 HEX"
      Height          =   255
      Index           =   17
      Left            =   3120
      TabIndex        =   51
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   60 HEX"
      Height          =   255
      Index           =   16
      Left            =   3120
      TabIndex        =   50
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "DI Board"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   33
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "STATUS"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   32
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   5F HEX"
      Height          =   255
      Index           =   15
      Left            =   600
      TabIndex        =   15
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   5E HEX"
      Height          =   255
      Index           =   14
      Left            =   600
      TabIndex        =   14
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   5D HEX"
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   13
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   5C HEX"
      Height          =   255
      Index           =   12
      Left            =   600
      TabIndex        =   12
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   5B HEX"
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   5A HEX"
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   59 HEX"
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   9
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   58 HEX"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   57 HEX"
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   56 HEX"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   55 HEX"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   54 HEX"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   53 HEX"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   52 HEX"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   51 HEX"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label DI_board_label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DI board   50 HEX"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "IOtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim response As Integer

bus_link_check.value = 1   ' Checked
' HALT the IO
old_RunIO = RunIO
RunIO = False
IO_check.value = 0     ' UN-Checked

response = MsgBox("BUS walking bit test will RUN when you click OK" & Chr$(13) _
& "Go to a back-plane and watch the LEDs" & Chr$(13) _
& "Wait for the STATUS check box to clear", vbOKOnly, " BUS TEST ")
   
  
' 20/2/98
VB_bus_test = True
Call set_VB_bus_test(VB_bus_test)

' Call the FORTRAN BUSLINK subroutine
Call BUSLINK
    
bus_link_check.value = 0   ' UN-Checked

' Restore the status of RunIO
RunIO = old_RunIO
      
End Sub

Private Sub Command10_Click()

' 20/2/98   DOES NOT WORK
' This is because this click event
' does NOT occur all the while you are
' in the START BUS LINK click event  ( Command1_click )
VB_bus_test = False
Call set_VB_bus_test(VB_bus_test)

End Sub

Private Sub Command11_Click()

   RunIO = False


' 28/4/98  HALT the IO task....
'Call setPCDI(2568)
VB_RunIO = False
Call set_RunIO(VB_RunIO)

End Sub

Private Sub Command2_Click()

Call set_DI_board_default
Call INITIO

End Sub

Private Sub Command3_Click()

   RunIO = True
    
' 28/4/98  RUN the IO ....
'Call setPCDI(2567)
    
 VB_RunIO = True
Call set_RunIO(VB_RunIO)
    
End Sub

Private Sub Command4_Click()
    Call free_io_board
End Sub

Private Sub Command5_Click()
    HWTestAnalog.Show
End Sub

Private Sub Command6_Click()
    HWTestDigital.Show
End Sub

Private Sub Command7_Click()


' 24/4/98
TEST_SCOPE.Show

End Sub

Private Sub Command8_Click()
Dim i As Integer

For i = &H50 To &H7F
 DI_board_present(i) = False
Next i

Call set_DI_boards(DI_board_present(1))


End Sub

Private Sub Command9_Click()

Call set_DI_board_default

End Sub

Private Sub DI_board_check_Click(Index As Integer)

' 19/2/98
' CAN NOT DO THIS HERE BECAUSE THIS EVENT ID TRIGGERED
' CONTINUOUSLY BY TIMER1_TIMER
'  SEE DI_BOARD_LABEL_CLICK instead

'If DI_board_check(Index).Value = 1 Then    ' CHECKED
'    DI_board_present(Index + &H50) = False
'    Call set_DI_boards(DI_board_present(1))
'Else
'    DI_board_present(Index + &H50) = True
'    Call set_DI_boards(DI_board_present(1))
'End If


End Sub

Private Sub DI_board_label_Click(Index As Integer)

If DI_board_check(Index).value = 1 Then    ' CHECKED
    DI_board_present(Index + &H50) = False
    Call set_DI_boards(DI_board_present(1))
Else
    DI_board_present(Index + &H50) = True
    Call set_DI_boards(DI_board_present(1))
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer



'22/4/98
Me.Height = 8000
Me.Width = 8685


' 29/4/98
'If runIO_counter = 5 Then
'      RunIO = True
'' 28/4/98  RUN the IO....
'     Call setPCDI(2567)
'End If


'If RunIO Then
If VB_RunIO Then

    '21/4/98 now done from host form Timer_IO:         Call io
    IO_check.value = 1     ' Checked
Else
    IO_check.value = 0     ' UN-Checked
End If


' 19/2/98
Call get_DI_boards(DI_board_present(1))

For i = &H50 To &H7F
   If DI_board_present(i) Then
         DI_board_check(i - &H50).value = 1   ' CHECKED
   Else
         DI_board_check(i - &H50).value = 0   ' NOT CHECKED
   End If
Next i

End Sub
