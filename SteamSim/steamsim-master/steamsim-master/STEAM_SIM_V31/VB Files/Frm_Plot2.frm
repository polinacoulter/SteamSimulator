VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "Comct232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "Graphs32.ocx"
Begin VB.Form frm_Plot2 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engine Dynamics"
   ClientHeight    =   9930
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   16365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   16365
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   6000
      Max             =   25
      TabIndex        =   102
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   2940
      Max             =   2000
      TabIndex        =   101
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11160
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cylinder Overview"
      Height          =   3315
      Left            =   12360
      TabIndex        =   71
      Top             =   3240
      Width           =   3855
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   19
         Left            =   960
         TabIndex        =   97
         Top             =   2880
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   18
         Left            =   960
         TabIndex        =   94
         Top             =   2460
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   17
         Left            =   960
         TabIndex        =   91
         Top             =   2040
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   16
         Left            =   960
         TabIndex        =   88
         Top             =   1620
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   15
         Left            =   960
         TabIndex        =   85
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   7
         Left            =   960
         TabIndex        =   82
         Top             =   780
         Width           =   624
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   74
         Text            =   "1"
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   73
         Text            =   "Cylinder 1-6"
         Top             =   360
         Width           =   1335
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   600
         TabIndex        =   72
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Text2(1)"
         BuddyDispid     =   196614
         BuddyIndex      =   1
         OrigLeft        =   840
         OrigTop         =   360
         OrigRight       =   1080
         OrigBottom      =   615
         Max             =   6
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   7
         Left            =   600
         TabIndex        =   83
         Top             =   780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(7)"
         BuddyDispid     =   196613
         BuddyIndex      =   7
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   15
         Left            =   600
         TabIndex        =   86
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(15)"
         BuddyDispid     =   196613
         BuddyIndex      =   15
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   16
         Left            =   600
         TabIndex        =   89
         Top             =   1620
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(16)"
         BuddyDispid     =   196613
         BuddyIndex      =   16
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   17
         Left            =   600
         TabIndex        =   92
         Top             =   2040
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(17)"
         BuddyDispid     =   196613
         BuddyIndex      =   17
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   18
         Left            =   600
         TabIndex        =   95
         Top             =   2460
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(18)"
         BuddyDispid     =   196613
         BuddyIndex      =   18
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   19
         Left            =   600
         TabIndex        =   98
         Top             =   2880
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(19)"
         BuddyDispid     =   196613
         BuddyIndex      =   19
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   19
         Left            =   1680
         TabIndex        =   96
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   18
         Left            =   1680
         TabIndex        =   93
         Top             =   2460
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   17
         Left            =   1680
         TabIndex        =   90
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   16
         Left            =   1680
         TabIndex        =   87
         Top             =   1620
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   15
         Left            =   1680
         TabIndex        =   84
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   81
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   13200
      TabIndex        =   70
      Top             =   9360
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Light Spring"
      Height          =   478
      Left            =   8880
      TabIndex        =   69
      Top             =   4320
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      Caption         =   "Engine Type"
      Height          =   3135
      Left            =   12360
      TabIndex        =   29
      Top             =   0
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "4-Stroke -Predictor"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   2700
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2-Stroke -Predictor"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   2280
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Medium 4-Stroke Sulzer ZA40"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Medium 4-Stroke Wartsila Vasa 32"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   3315
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Slow  2-Stroke MAN B&&W 6LM60"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   3315
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Slow  2-Stroke MAN B&&W 6L90MC"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   3315
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Caption         =   "Help Files"
      Height          =   1335
      Left            =   8880
      TabIndex        =   27
      Top             =   4920
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "Spray"
         Height          =   375
         Left            =   1740
         TabIndex        =   100
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "NOx"
         Height          =   375
         Left            =   480
         TabIndex        =   99
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Plot Type"
      Height          =   3920
      Left            =   8880
      TabIndex        =   19
      Top             =   60
      Width           =   3405
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   2628
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Visible         =   0   'False
         Width           =   2628
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   2628
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   2628
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   2628
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   2628
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   485
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2628
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame 1"
      Height          =   2715
      Left            =   240
      TabIndex        =   4
      Top             =   6540
      Width           =   15945
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   23
         Left            =   13020
         TabIndex        =   78
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   22
         Left            =   13020
         TabIndex        =   76
         Top             =   1200
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   21
         Left            =   13020
         TabIndex        =   66
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   20
         Left            =   13020
         TabIndex        =   65
         Top             =   240
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   14
         Left            =   6840
         TabIndex        =   57
         Top             =   2160
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   13
         Left            =   6840
         TabIndex        =   56
         Top             =   1680
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   12
         Left            =   6840
         TabIndex        =   55
         Top             =   1200
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   11
         Left            =   6840
         TabIndex        =   38
         Top             =   720
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   9
         Left            =   3840
         TabIndex        =   37
         Top             =   2160
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   8
         Left            =   3840
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   10
         Left            =   6840
         TabIndex        =   28
         Top             =   240
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   1
         Left            =   600
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   1200
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   3
         Left            =   600
         TabIndex        =   9
         Top             =   1680
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   4
         Left            =   600
         TabIndex        =   8
         Top             =   2160
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   5
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   624
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   332
         Index           =   6
         Left            =   3840
         TabIndex        =   5
         Top             =   720
         Width           =   624
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(0)"
         BuddyDispid     =   196613
         BuddyIndex      =   0
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   332
         Index           =   1
         Left            =   359
         TabIndex        =   44
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(1)"
         BuddyDispid     =   196613
         BuddyIndex      =   1
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   332
         Index           =   2
         Left            =   359
         TabIndex        =   45
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(2)"
         BuddyDispid     =   196613
         BuddyIndex      =   2
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   332
         Index           =   3
         Left            =   359
         TabIndex        =   46
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(3)"
         BuddyDispid     =   196613
         BuddyIndex      =   3
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   4
         Left            =   360
         TabIndex        =   47
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(4)"
         BuddyDispid     =   196613
         BuddyIndex      =   4
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   5
         Left            =   3600
         TabIndex        =   48
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(5)"
         BuddyDispid     =   196613
         BuddyIndex      =   5
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   6
         Left            =   3600
         TabIndex        =   49
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(6)"
         BuddyDispid     =   196613
         BuddyIndex      =   6
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   8
         Left            =   3600
         TabIndex        =   50
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(8)"
         BuddyDispid     =   196613
         BuddyIndex      =   8
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   9
         Left            =   3600
         TabIndex        =   51
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(9)"
         BuddyDispid     =   196613
         BuddyIndex      =   9
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   10
         Left            =   6600
         TabIndex        =   52
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(10)"
         BuddyDispid     =   196613
         BuddyIndex      =   10
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   11
         Left            =   6600
         TabIndex        =   53
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(11)"
         BuddyDispid     =   196613
         BuddyIndex      =   11
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   12
         Left            =   6600
         TabIndex        =   54
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(12)"
         BuddyDispid     =   196613
         BuddyIndex      =   12
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   13
         Left            =   6600
         TabIndex        =   58
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(13)"
         BuddyDispid     =   196613
         BuddyIndex      =   13
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   14
         Left            =   6600
         TabIndex        =   59
         Top             =   2160
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(14)"
         BuddyDispid     =   196613
         BuddyIndex      =   14
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   20
         Left            =   12780
         TabIndex        =   63
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(20)"
         BuddyDispid     =   196613
         BuddyIndex      =   20
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   21
         Left            =   12780
         TabIndex        =   64
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(21)"
         BuddyDispid     =   196613
         BuddyIndex      =   21
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   127
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   22
         Left            =   12780
         TabIndex        =   75
         Top             =   1200
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(22)"
         BuddyDispid     =   196613
         BuddyIndex      =   22
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   100
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   330
         Index           =   23
         Left            =   12780
         TabIndex        =   77
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   327681
         Value           =   78
         Alignment       =   0
         BuddyControl    =   "Text1(23)"
         BuddyDispid     =   196613
         BuddyIndex      =   23
         OrigLeft        =   10920
         OrigTop         =   6960
         OrigRight       =   11160
         OrigBottom      =   7455
         Max             =   100
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   23
         Left            =   13740
         TabIndex        =   80
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   22
         Left            =   13740
         TabIndex        =   79
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   21
         Left            =   13740
         TabIndex        =   68
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   20
         Left            =   13740
         TabIndex        =   67
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   14
         Left            =   7560
         TabIndex        =   62
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   13
         Left            =   7560
         TabIndex        =   61
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   12
         Left            =   7560
         TabIndex        =   60
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   11
         Left            =   7560
         TabIndex        =   42
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   10
         Left            =   7560
         TabIndex        =   41
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   9
         Left            =   4560
         TabIndex        =   40
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   8
         Left            =   4560
         TabIndex        =   39
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reverse X Axis"
      Height          =   478
      Left            =   8880
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit LU Table"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   9360
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   11760
      Top             =   4080
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   6105
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   8430
      _Version        =   327680
      _ExtentX        =   14870
      _ExtentY        =   10769
      _StockProps     =   96
      BorderStyle     =   1
      DrawMode        =   1
      FontName        =   "Times New Roman~Times New Roman~Times New Roman~Times New Roman"
      GraphType       =   6
      ThickLines      =   0
      XAxisStyle      =   2
      XAxisTicks      =   10
      LabelYFormat    =   ""
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   727
      TabIndex        =   3
      Top             =   1561
      Width           =   7935
   End
End
Attribute VB_Name = "frm_Plot2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Sub CylVar()

Select Case CylNumber

Case 1
    CylVariation = 1.01
    If Fault(101) Then
        CylVariation = CylVariation * 0.65
    End If
    If Fault(107) Then
        CylVariation = CylVariation * 0.9
    End If

Case 2
  '  19/1/99    CylVariation = 9.9
                CylVariation = 0.99
    If Fault(102) Then
        CylVariation = CylVariation * 0.8
    End If
    If Fault(108) Then
        CylVariation = CylVariation * 0.9
    End If


Case 3
    CylVariation = 0.95
    If Fault(103) Then
        CylVariation = CylVariation * 0.65
    End If
    If Fault(109) Then
        CylVariation = CylVariation * 0.9
    End If


Case 4
    CylVariation = 1.06
    If Fault(104) Then
        CylVariation = CylVariation * 0.65
    End If
    If Fault(110) Then
        CylVariation = CylVariation * 0.9
    End If
    
Case 5
    CylVariation = 1#
    If Fault(105) Then
        CylVariation = CylVariation * 0.65
    End If
    If Fault(111) Then
        CylVariation = CylVariation * 0.9
    End If
    
Case 6
    CylVariation = 1.04
    If Fault(106) Then
        CylVariation = CylVariation * 0.65
    End If
    If Fault(112) Then
        CylVariation = CylVariation * 0.9
    End If

    
End Select




End Sub
Sub Display_NOx()

Dim i As Integer

'Set up Graph
    Graph1.GraphTitle = "NOx Emmission"
    Graph1.LeftTitle = "NOx ppm"
    Graph1.LeftTitleStyle = 1  'up
    Graph1.BottomTitle = "Engine Speed"
    Graph1.GridStyle = 3
    Graph1.YAxisMax = 25
    Graph1.XAxisMin = 0
    'If TwoStroke Then
        'Graph1.XAxisMax = 150
    'Else
        'Graph1.XAxisMax = 1000
    'End If
    Graph1.XAxisMax = 2000
    
    Graph1.XAxisTicks = 20
    Graph1.NumSets = 1
    Graph1.NumPoints = 200
    Graph1.ThisSet = 1
    Graph1.ColorData = 0

'draw reference
    For i = 1 To 13
        Graph1.Data(i) = 17
        Graph1.xpos(i) = i * 10
    Next
    
    For i = 14 To 200
        Graph1.Data(i) = Graph1.Data(i - 1) - (((Graph1.Data(i - 1) - 9.8) * 0.16)) / 9.8
        Debug.Print "i = " & i & " data = " & Graph1.Data(i)
        'Pc = Pc - (((Pc - AmbientAirPressure) * ScavPDecayRate) / AmbientAirPressure)

        Graph1.xpos(i) = i * 10
    Next
    Graph1.DrawMode = 3
    
End Sub
Sub SetUp6LM60_P()
Dim i As Integer


bore = 600
'stroke = 3.2 * bore
stroke = 1944

throw = stroke / 2
CylLength = 2131

LengthConRod = 3196
PistonOffset = LengthConRod - throw
'PistonOffset = 3000
MaxEngRevs = 117

Gamma = 1.4
IgnitionVar = 0
ScavAirPressure = 3#
IgnitionVar = 0
ScavAirPressure = 3.1
AirValveOpenAngle = 60
ExhVVOpenAngle = 110
InjectionAngle = 12
IgnitionDelay = 3
FuelInjPeriod = 5
CetaneVal = 50
SprayAngle = 0

NomLinerTemp = 385
NomCylWallTemp = 486
NomPistonCrownTemp = 497
LinerTemp = NomLinerTemp
CylWallTemp = NomCylWallTemp
PistonCrownTemp = NomPistonCrownTemp

DiamInjParticle = 0.01
ScavEfficiency = 100
NOx = 9.8

ScavT = 41.2
ScavH = 68

CylNumber = 1

ExhVVDecayRate = 0.1
STEPs = 180
MinPressProportion = 0.2    'determines lowest pressure  as proportion of (P - Ambient P)
                            'when p falls after Exh VV opens at end of power stroke

ScavInletRiseRate = 0.15     'After lowest P at end of Power Stroke, Air V/V open and
                            'Scav Air Enters. This variable determines rate of rise back to Scav Press

ConstScavPAngle = 5         'At bdc, start of compression, ExhVV and AIr VV are open, but
                            'Scav Air maintains const pressure for a few deg of crank
                            'NOTE: MISTAKE - THIS IS SAME POINT AS AIR V/V OPEN !

ScavPDecayRate = 0.01        'At start Compression after Air VV closes, P falls a little before
                            'Ecxh VV closes

'Graph Values
Graph1.YAxisStyle = 2
Graph1.YAxisMin = 0
Graph1.YAxisMax = 150
YaxisOffset = 0.004



TwoStroke = True   'for now

'set UpDowns with default values, thopse commented out can't be set
UpDown1(0).Value = CylLength
UpDown1(1).Value = bore
UpDown1(2).Value = stroke
UpDown1(3).Value = LengthConRod
'UpDown1(4).value = CompRatio
'UpDown1(5).Value = 75  'for test
UpDown1(5).Max = MaxEngRevs
UpDown1(5).Value = MaxEngRevs * 0.9  'for test

'UpDown1(6).value = MaxSpeed
'UpDown1(7).value = MaxPCompression
UpDown1(8).Value = ScavAirPressure * 100 'UPDOWNs cant increment decimals
UpDown1(9).Value = AirValveOpenAngle
UpDown1(10).Value = ExhVVOpenAngle
UpDown1(11).Value = InjectionAngle
UpDown1(12).Value = IgnitionDelay
UpDown1(13).Value = FuelInjPeriod + 1
UpDown1(14).Value = CetaneVal + 1
UpDown1(15).Value = SprayAngle + 1
'UpDown1(16).value = LinerTemp
'UpDown1(17).value = CylWallTemp
'UpDown1(18).value = PistonCrownTemp
'UpDown1(19).value = DiamInjParticle
'UpDown1(20).value = ScavEfficiency
'UpDown1(21).value = NOx

'18/1/99 enable all relevant UpDown controls in Predictor Mode
'For i = 0 To UpDown1.Count - 1
    UpDown1(0).Visible = True 'Cylinder Length (mm)
    UpDown1(1).Visible = True 'bore (mm)
    UpDown1(2).Visible = True 'stroke (mm)
    UpDown1(3).Visible = True 'Con Rod Length (mm)
    UpDown1(5).Visible = True 'Engine speed(rpm)
    UpDown1(13).Visible = True 'Fuel Inj Period (mS)
    UpDown1(14).Visible = True 'CETANE Value
    UpDown1(19).Visible = True 'Diam of Inj Fuel (mm)


'Next

Call SetValues

End Sub
Sub SetUpSulzerZA40()
Dim i As Integer

'Med Speed SulzerZA40

bore = 400
'stroke = 3.2 * bore
stroke = 560

throw = stroke / 2
CylLength = 602 '(560 * 14.1)/(14.1-1)

LengthConRod = 903 'stroke * approx 1.5
PistonOffset = LengthConRod - throw
'PistonOffset = 3000
MaxEngRevs = 510

Gamma = 1.4
IgnitionVar = 0
ScavAirPressure = 3.1
AirValveOpenAngle = 60
ExhVVOpenAngle = 110
InjectionAngle = 12
IgnitionDelay = 3
FuelInjPeriod = 5
CetaneVal = 50
SprayAngle = 0

NomLinerTemp = 416
NomCylWallTemp = 501
NomPistonCrownTemp = 523
LinerTemp = NomLinerTemp
CylWallTemp = NomCylWallTemp
PistonCrownTemp = NomPistonCrownTemp

DiamInjParticle = 0.01
ScavEfficiency = 96.4

NOx = 9.83

ScavT = 44.7
ScavH = 68

CylNumber = 1

STEPs = 180
MinPressProportion = 0.2    'determines lowest pressure  as proportion of (P - Ambient P)
                            'when p falls after Exh VV opens at end of power stroke

ScavInletRiseRate = 0.15     'After lowest P at end of Power Stroke, Air V/V open and
                            'Scav Air Enters. This variable determines rate of rise back to Scav Press

ConstScavPAngle = 5         'At bdc, start of compression, ExhVV and AIr VV are open, but
                            'Scav Air maintains const pressure for a few deg of crank
                            'NOTE: MISTAKE - THIS IS SAME POINT AS AIR V/V OPEN !

ScavPDecayRate = 0.01        'At start Compression after Air VV closes, P falls a little before
                            'Ecxh VV closes

'Graph Values
Graph1.YAxisStyle = 2
Graph1.YAxisMin = 0
Graph1.YAxisMax = 150
YaxisOffset = 0.004



TwoStroke = False   'for now

'set UpDowns with default values, thopse commented out can't be set
UpDown1(0).Value = CylLength
UpDown1(1).Value = bore
UpDown1(2).Value = stroke
''UpDown1(3).Value = LengthConRod
'UpDown1(4).value = CompRatio
UpDown1(5).Max = MaxEngRevs
UpDown1(5).Value = MaxEngRevs * 0.9  'for test

'UpDown1(6).value = MaxSpeed
'UpDown1(7).value = MaxPCompression
UpDown1(8).Value = ScavAirPressure * 100 'UPDOWNs cant increment decimals
UpDown1(9).Value = AirValveOpenAngle
UpDown1(10).Value = ExhVVOpenAngle
UpDown1(11).Value = InjectionAngle
UpDown1(12).Value = IgnitionDelay
UpDown1(13).Value = FuelInjPeriod + 1
UpDown1(14).Value = CetaneVal + 1
UpDown1(15).Value = SprayAngle + 1
'UpDown1(16).value = LinerTemp
'UpDown1(17).value = CylWallTemp
'UpDown1(18).value = PistonCrownTemp
'UpDown1(19).value = DiamInjParticle
'UpDown1(20).value = ScavEfficiency
'UpDown1(21).value = NOx


'12/6/98 disable all UpDown controls - except in Predictor Mode
For i = 0 To UpDown1.Count - 1
    UpDown1(i).Visible = False
Next
'exception is speed for testing
UpDown1(5).Visible = True

Call SetValues

End Sub

Sub SetUpVasa32()
'Med Speed Vasa 32
Dim i As Integer

bore = 320
'stroke = 3.2 * bore
stroke = 350

throw = stroke / 2
CylLength = 377 '(350 * 14.1)/(14.1-1)

LengthConRod = 570 'stroke * approx 1.6
PistonOffset = LengthConRod - throw
'PistonOffset = 3000
MaxEngRevs = 750

Gamma = 1.4
IgnitionVar = 0
ScavAirPressure = 2.1
AirValveOpenAngle = 60
ExhVVOpenAngle = 110
InjectionAngle = 12
IgnitionDelay = 3
FuelInjPeriod = 5
CetaneVal = 50
SprayAngle = 0

NomLinerTemp = 455
NomCylWallTemp = 496
NomPistonCrownTemp = 513
LinerTemp = NomLinerTemp
CylWallTemp = NomCylWallTemp
PistonCrownTemp = NomPistonCrownTemp

DiamInjParticle = 0.01
ScavEfficiency = 100
NOx = 9.8

ScavT = 40.7
ScavH = 68

CylNumber = 1

ExhVVDecayRate = 0.1
STEPs = 180
MinPressProportion = 0.2    'determines lowest pressure  as proportion of (P - Ambient P)
                            'when p falls after Exh VV opens at end of power stroke

ScavInletRiseRate = 0.15     'After lowest P at end of Power Stroke, Air V/V open and
                            'Scav Air Enters. This variable determines rate of rise back to Scav Press

ConstScavPAngle = 5         'At bdc, start of compression, ExhVV and AIr VV are open, but
                            'Scav Air maintains const pressure for a few deg of crank
                            'NOTE: MISTAKE - THIS IS SAME POINT AS AIR V/V OPEN !

ScavPDecayRate = 0.01        'At start Compression after Air VV closes, P falls a little before
                            'Ecxh VV closes

'Graph Values
Graph1.YAxisStyle = 2
Graph1.YAxisMin = 0
Graph1.YAxisMax = 150
YaxisOffset = 0.004



TwoStroke = False   'for now

'set UpDowns with default values, thopse commented out can't be set
UpDown1(0).Value = CylLength
UpDown1(1).Value = bore
UpDown1(2).Value = stroke
UpDown1(3).Value = LengthConRod
'UpDown1(4).value = CompRatio
'UpDown1(5).Value = 75  'for test
UpDown1(5).Max = MaxEngRevs
UpDown1(5).Value = MaxEngRevs * 0.9  'for test

'UpDown1(6).value = MaxSpeed
'UpDown1(7).value = MaxPCompression
UpDown1(8).Value = ScavAirPressure * 100 'UPDOWNs cant increment decimals
UpDown1(9).Value = AirValveOpenAngle
UpDown1(10).Value = ExhVVOpenAngle
UpDown1(11).Value = InjectionAngle
UpDown1(12).Value = IgnitionDelay
UpDown1(13).Value = FuelInjPeriod + 1
UpDown1(14).Value = CetaneVal + 1
UpDown1(15).Value = SprayAngle + 1
'UpDown1(16).value = LinerTemp
'UpDown1(17).value = CylWallTemp
'UpDown1(18).value = PistonCrownTemp
'UpDown1(19).value = DiamInjParticle
'UpDown1(20).value = ScavEfficiency
'UpDown1(21).value = NOx


'12/6/98 disable all UpDown controls - except in Predictor Mode
For i = 0 To UpDown1.Count - 1
    UpDown1(i).Visible = False
Next
'exception is speed for testing
UpDown1(5).Visible = True

Call SetValues


End Sub
Sub Display_PV_LightSpring()
'Similar to PV at region near bdc bt with P plotted gainst Crank Angle


Dim i As Integer
Dim Vcx As Single
Dim ExhVVOpenVol As Single

'Set up Graph
    Graph1.YAxisMax = 10 'light spring
    Graph1.XAxisMin = 0
    Graph1.XAxisMax = 360 'normalised

    Graph1.NumSets = 1
    Graph1.NumPoints = 360
    Graph1.ThisSet = 1
    Graph1.ColorData = 0
    
    
    

'Initial Calculations of: Min, Swept & Max Vols, Compression Ratio and Unit of Vol as proportion max Vol
    MinVol = pi * bore * bore * (CylLength - (throw * 2)) * 0.000000001 'unswept vol
    SweptVol = (pi * bore * bore * (throw * 2) * 0.000000001)
    MaxVol = SweptVol + MinVol
    CompRatio = MaxVol / MinVol
    UnitOfVol = SweptVol / STEPs
   

'get initial volume bdc, we know pressure is the Scav Pressure so calc K1
    angle = 0
    Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
    Call GetVol(bore, CylLength, throw, Position, Vc)
    K1 = ScavAirPressure * Exp(Log(Vc + MinVol) * Gamma) ' find gamma from P*(V**Gamma) = Constant (K1)
    

'Find The Volume when the Air V/V closes
'    angle = AirValveOpenAngle 'Angle at which air v/v closes
'    Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
'    Call GetVol(bore, CylLength, throw, Position, Vc)
    
    
'how many units of normalised volume has the piston covered at this point ?
'    Start = CInt((SweptVol - Vc) / UnitOfVol)
 '   AirValveOpenVol = Start 'remember it because we'll need it again later
    
'Set the Pressure to ScavAirPress while air v/v is open
    For i = 1 To AirValveOpenAngle
        Points(i) = ScavAirPressure 'assume ambient pressure while air inlet v/v open
    Next
  
'  Debug.Print "AirVVCloses at " & Start & "  units of vol"
  
'Find The Volume when the Exh V/V closes
'        angle = 180 - ExhVVOpenAngle 'Angle at which air v/v closes is same as angle at which it opens
'        Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
'        Call GetVol(bore, CylLength, throw, Position, Vc)
    
    
'how many units of normalised volume has the piston covered at this point ?
'        ExhVVOpenVol = CInt((SweptVol - Vc) / UnitOfVol)
'        Debug.Print "ExhVVCloses at " & ExhVVOpenVol & "  units of vol"

'Set the Pressure to fall slowly towards ambient after Air V/V shuts and before EXh VV shuts
    Pc = ScavAirPressure
    For i = AirValveOpenAngle + 1 To 180 - ExhVVOpenAngle
            Pc = Pc - (((Pc - AmbientAirPressure) * ScavPDecayRate) / AmbientAirPressure)
            'Pc = Pc - (((Pc - ScavAirPressure) * ExhVVDecayRate) / ScavAirPressure)
            Points(i) = Pc 'assume ambient pressure while air inlet v/v open
    Next
  
'get a new K
    K1 = Pc * Exp(Log(Vc + MinVol) * Gamma) ' find gamma from P*(V**Gamma) = Constant (K1)

  

'Calc Pressure for each unit of Volume for main compression stroke, up to tdc
    MaxPCompression = 0
    'For i = Start + 1 To STEPs
        For i = 180 - ExhVVOpenAngle + 1 To 180
                'Vc = ((STEPs - i) * UnitOfVol) + MinVol 'vol left decreases with number of STEPs
                'Pc = K1 / Exp(Log(Vc) * Gamma) 'calc the pressure at this volume
                Call GetPosition(i, throw, LengthConRod, PistonOffset, Position)
                Call GetVol(bore, CylLength, throw, Position, Vc)
                Vc = Vc + MinVol
                Pc = K1 / Exp(Log(Vc) * Gamma) 'calc the pressure at this volume
                Points(i) = Pc
                If Pc > MaxPCompression Then
                    MaxPCompression = Pc
                End If
            
        Next
        
        Graph1.ColorData = 1 'set the plot colur (blue)
        Points(STEPs) = LastPressure 'remember the last pressure - we need it for the combustion part

        For i = 1 To 180 ' now store plot values for the compression stroke
            'If Check1.value = 1 Then
                'Graph1.Data(i) = Points(STEPs + 1 - i) 'graphdata starts at 1 not 0
            'Else
            Graph1.Data(i) = Points(i) 'graphdata starts at 1 not 0
           'End If
        
            'Graph1.xpos(i) = ((STEPs - i) / STEPs) 'normalised volume starting with pos(1) at 1.0
            Graph1.xpos(i) = (i) 'normalised volume starting with pos(1) at 1.0
            'Debug.Print "Point # " & i & "  Position = " & Graph1.xpos(i) & " Data = " & Graph1.Data(i)

        Next

'============================ END OF COMPRESSION =====================================
    
    
'now do same again for power (called exhaust) stroke
'first few points are from Look-up table - decide here which one to use
    
    
    NumberOfPoints = 0
    
'find how many points there are to be plotted from the look-up table (+ve values only)
    For i = 1 To UBound(PVLookUp)
        If PVLookUp(i) > 0 Then
            NumberOfPoints = NumberOfPoints + 1
        End If
    Next
    
'calculate a multiplier for x-axis based on Engine speed and some constant
    SpeedMult = ((1 - (engrevs / MaxEngRevs)) * 0.8) ' slower means fatter
    Xm(1) = YaxisOffset 'make first point close to but not on Y-axis so line appears continuous

    For i = 1 To NumberOfPoints 'fill Xm array with modified x-axis values
        If i <= UBound(Xm) Then
            Xm(i + 1) = i + (i * SpeedMult) 'so for a Mult Factor of 0.8, X-axis values is at  1.8, 3.6,5.4 etc
        End If
    Next
    Xm(0) = 0
    '------------------
                
    

    
    '------------------
    
    
    NumberOfNewPoints = Xm(NumberOfPoints)
        counter = 180
        For i = 1 To NumberOfPoints 'each i is a unit volume

            If PVLookUp(i) > 0 Then
                'LastPressure = MaxPCompression + (engrevs / MaxEngRevs) * MaxPCompression * (PVLookUp(i) / 100)
                'LastAngle = Xm(i)
                'Points(STEPs + i) = LastPressure
                'xpos(STEPs + i) = (Xm(i) / STEPs)
                LastPressure = MaxPCompression + (engrevs / MaxEngRevs) * MaxPCompression * (PVLookUp(i) / 100)
                LastAngle = CInt(Xm(i))
                counter = counter + 1
                Points(counter) = LastPressure
                xpos(counter) = 180 - (Xm(i))

            Else
                Exit For
            End If
        Next
    
        'calc a new k2 based on last vol and press
        Call GetPosition(LastAngle, throw, LengthConRod, PistonOffset, Position)
        Call GetVol(bore, CylLength, throw, Position, Vc)

        Vc = Vc + MinVol
        K2 = LastPressure * Exp(Log(Vc) * Gamma) '  P*(V**Gamma) = Constant (K1)
        

'Calculate P for each degree left during power stroke
        Gamma = 1.38
        'counter = NumberOfPoints
    
        incr = (STEPs - Start) / (STEPs - NumberOfPoints)
        Debug.Print "LastAngle = " & LastAngle & " Exhvvopenangle = " & ExhVVOpenAngle
'calculate P for each degree up to Exh VV open
        For i = LastAngle To ExhVVOpenAngle

            counter = counter + 1
            Call GetPosition(i, throw, LengthConRod, PistonOffset, Position)
            Call GetVol(bore, CylLength, throw, Position, Vc)
            Vc = Vc + MinVol
            Pc = K2 / Exp(Log(Vc) * Gamma) 'calc the pressure at this volume
            Points(counter) = Pc
            xpos(counter) = 180 - i
            Debug.Print "After LastAngle, i = " & i & " counter = " & counter & "xpos(counter) = " & xpos(counter)
            Debug.Print "Vc+MinVol = " & Vc & " Press = " & Pc
        Next

'now do the decay to Scav Air Press after the Exh V/V opens until Air V/V open

        'For ix = ExhVVOpenVol + 1 To STEPs STEP incr
        For ix = ExhVVOpenAngle + 1 To STEPs - AirValveOpenAngle

            counter = counter + 1
           
            Pc = Pc - (((Pc - (ScavAirPressure * MinPressProportion)) * ExhVVDecayRate) / (ScavAirPressure * MinPressProportion))
            
            Points(counter) = Pc
            xpos(counter) = 180 - (ix)
       Next

'now do the rise to Scav Air Press after the Air V/V open to bdc

        'For ix = ExhVVOpenVol + 1 To STEPs STEP incr
        For ix = STEPs - AirValveOpenAngle + 1 To STEPs

            counter = counter + 1
            Pc = Pc + (((ScavAirPressure - Pc) * ScavInletRiseRate) / ScavAirPressure)
            
            Points(counter) = Pc
            xpos(counter) = 180 - (ix)
       Next
        
        
        
        Graph1.ThisSet = 1
        Graph1.ColorData = 1

        For i = STEPs + 1 To (STEPs * 2) '181 to 360

            'If Check1.value = 0 Then
                Graph1.Data(i) = Points(i)
           'Else
                'Graph1.Data(i) = Points((STEPs * 2) + 1 - i)
            'End If
            Graph1.xpos(i) = xpos(i)
            'Debug.Print "Point # " & i & "  Position = " & Graph1.xpos(i) & " Data = " & Graph1.Data(i)

        Next
        
'draw the graph
        Graph1.DrawMode = 3
        
       
End Sub
Sub SetValues()
'Copies values of Cyl Length, etc into text boxes and UpDown controls

Text1(0).Text = CylLength
Text1(1).Text = bore
Text1(2).Text = stroke
Text1(3).Text = LengthConRod
Text1(4).Text = CompRatio

' 18/1/99 RWJ , Get some values from the MODEL
If Option2(0).Value Or Option2(2).Value Then   ' (Get the engine rpm from the MODEL)

   If Option2(0).Value Then ' SLOW speed
   
            engrevs = LSS_reals(50)
            
            DiamInjParticle = LSS_reals(745 + CylNumber) / 240# * 0.045 + 0.005
            
            T_exhaust = LSS_reals(68 + CylNumber)
            
   Else ' Medium Speed
   
            engrevs = VB_ENGINE(92)
            
            DiamInjParticle = VB_ENGINE(128 + CylNumber) / 240# * 0.045 + 0.005
            
            T_exhaust = VB_ENGINE(97 + CylNumber)
            
   End If
        UpDown1(5).Visible = False
Else
        UpDown1(5).Visible = True
        
End If


Text1(5).Text = Format(engrevs, "##0")

Text1(19).Text = Format(DiamInjParticle, "#0.000")


Text1(6).Text = MaxSpeed

'Text1(7).Text = MaxPCompression
Text1(7).Text = Format(MaxPCompression, "##0")

Text1(8).Text = ScavAirPressure
Text1(9).Text = AirValveOpenAngle
Text1(10).Text = ExhVVOpenAngle
Text1(11).Text = InjectionAngle
Text1(12).Text = IgnitionDelay
Text1(13).Text = FuelInjPeriod
Text1(14).Text = CetaneVal

' 18/1/99 RWJ Curve fit using 2 straight lines...
'  SprayAngle = m*DiamInjParticle + c
' Data:
'  SprayAngle    DiamInjParticle
'     40deg          0.05mm
'     50deg          0.01mm
'     60deg          0.005mm
If DiamInjParticle >= 0.05 Then
    SprayAngle = 40#
ElseIf DiamInjParticle >= 0.01 Then
    SprayAngle = -250# * DiamInjParticle + 52.5
ElseIf DiamInjParticle >= 0.005 Then
    SprayAngle = -2000# * DiamInjParticle + 70#
Else
    SprayAngle = 60#
End If

Text1(15).Text = Format(SprayAngle, "##0")


' 19/1/99
If Option2(0).Value Or Option2(2).Value Then   ' (Get the engine rpm from the MODEL)
    Diam_constant = DiamInjParticle + 0.99       ' ( range 0.005+0.99=0.995  to  0.05+0.99=1.04 )
    LinerTemp = T_exhaust * Diam_constant * 1.05

    Text1(16).Text = Format(LinerTemp, "##0")
Else
    Text1(16).Text = Format(LinerTemp + (LinerTemp * 0.1 * engrevs / MaxEngRevs), "##0")
End If


If Option2(0).Value Or Option2(2).Value Then   ' (Get the engine rpm from the MODEL)
    CylWallTemp = T_exhaust * Diam_constant * 0.95
    Text1(17).Text = Format(CylWallTemp, "##0")
Else
    Text1(17).Text = Format(CylWallTemp + (CylWallTemp * 0.05 * engrevs / MaxEngRevs), "##0")
End If


If Option2(0).Value Or Option2(2).Value Then   ' (Get the engine rpm from the MODEL)
    PistonCrownTemp = T_exhaust * Diam_constant * 1.15
    Text1(18).Text = Format(PistonCrownTemp, "##0")
Else
    Text1(18).Text = Format(PistonCrownTemp + (PistonCrownTemp * 0.15 * engrevs / MaxEngRevs), "##0")
End If


Text1(20).Text = ScavEfficiency
Text1(21).Text = NOx
'new 17/6/98
Text1(22).Text = ScavT + (ScavT * 0.15 * engrevs / MaxEngRevs)
Text1(23).Text = ScavH

Text2(1).Text = CylNumber





End Sub
Sub SetUp6LM60()
'Slow Speed B&W 6L90MC
Dim i As Integer

bore = 600
'stroke = 3.2 * bore
stroke = 1944

throw = stroke / 2
CylLength = 2131

LengthConRod = 3196
PistonOffset = LengthConRod - throw
'PistonOffset = 3000
MaxEngRevs = 117

Gamma = 1.4
IgnitionVar = 0
ScavAirPressure = 3.1
AirValveOpenAngle = 60
ExhVVOpenAngle = 110
InjectionAngle = 12
IgnitionDelay = 3
FuelInjPeriod = 5
CetaneVal = 50
SprayAngle = 0

NomLinerTemp = 385
NomCylWallTemp = 486
NomPistonCrownTemp = 497
LinerTemp = NomLinerTemp
CylWallTemp = NomCylWallTemp
PistonCrownTemp = NomPistonCrownTemp

DiamInjParticle = 0.01
ScavEfficiency = 100
NOx = 9.8

ScavT = 41.2
ScavH = 68

CylNumber = 1

ExhVVDecayRate = 0.1


STEPs = 180
MinPressProportion = 0.2    'determines lowest pressure  as proportion of (P - Ambient P)
                            'when p falls after Exh VV opens at end of power stroke

ScavInletRiseRate = 0.15     'After lowest P at end of Power Stroke, Air V/V open and
                            'Scav Air Enters. This variable determines rate of rise back to Scav Press

ConstScavPAngle = 5         'At bdc, start of compression, ExhVV and AIr VV are open, but
                            'Scav Air maintains const pressure for a few deg of crank
                            'NOTE: MISTAKE - THIS IS SAME POINT AS AIR V/V OPEN !

ScavPDecayRate = 0.01        'At start Compression after Air VV closes, P falls a little before
                            'Ecxh VV closes

'Graph Values
Graph1.YAxisStyle = 2
Graph1.YAxisMin = 0
Graph1.YAxisMax = 150
YaxisOffset = 0.004



TwoStroke = True   'for now

'set UpDowns with default values, thopse commented out can't be set
UpDown1(0).Value = CylLength
UpDown1(1).Value = bore
UpDown1(2).Value = stroke
UpDown1(3).Value = LengthConRod
'UpDown1(4).value = CompRatio
'UpDown1(5).Value = 75  'for test
UpDown1(5).Max = MaxEngRevs
UpDown1(5).Value = MaxEngRevs * 0.9  'for test

'UpDown1(6).value = MaxSpeed
'UpDown1(7).value = MaxPCompression
UpDown1(8).Value = ScavAirPressure * 100 'UPDOWNs cant increment decimals
UpDown1(9).Value = AirValveOpenAngle
UpDown1(10).Value = ExhVVOpenAngle
UpDown1(11).Value = InjectionAngle
UpDown1(12).Value = IgnitionDelay
UpDown1(13).Value = FuelInjPeriod + 1
UpDown1(14).Value = CetaneVal + 1
UpDown1(15).Value = SprayAngle + 1
'UpDown1(16).value = LinerTemp
'UpDown1(17).value = CylWallTemp
'UpDown1(18).value = PistonCrownTemp
'UpDown1(19).value = DiamInjParticle
'UpDown1(20).value = ScavEfficiency
'UpDown1(21).value = NOx

'12/6/98 disable all UpDown controls - except in Predictor Mode
For i = 0 To UpDown1.Count - 1
    UpDown1(i).Visible = False
Next
'exception is speed for testing
UpDown1(5).Visible = True


Call SetValues


End Sub

Sub SetUp6L90MC()
Dim i As Integer

'Slow Speed B&W 6L90MC


bore = 900
'stroke = 3.2 * bore
stroke = 2916

throw = stroke / 2
'CylLength = 3230.8
CylLength = 3202

LengthConRod = 4500
PistonOffset = LengthConRod - throw
'PistonOffset = 3000
MaxEngRevs = 78

Gamma = 1.4
IgnitionVar = 0
ScavAirPressure = 3.1
AirValveOpenAngle = 60
ExhVVOpenAngle = 110
InjectionAngle = 12
IgnitionDelay = 3
FuelInjPeriod = 5
CetaneVal = 50
SprayAngle = 0

NomLinerTemp = 385
NomCylWallTemp = 486
NomPistonCrownTemp = 497
LinerTemp = NomLinerTemp
CylWallTemp = NomCylWallTemp
PistonCrownTemp = NomPistonCrownTemp

DiamInjParticle = 0.01
ScavEfficiency = 100
NOx = 9.8

ScavT = 43.6
ScavH = 68

CylNumber = 1




ExhVVDecayRate = 0.1




STEPs = 180
MinPressProportion = 0.2    'determines lowest pressure  as proportion of (P - Ambient P)
                            'when p falls after Exh VV opens at end of power stroke

ScavInletRiseRate = 0.15     'After lowest P at end of Power Stroke, Air V/V open and
                            'Scav Air Enters. This variable determines rate of rise back to Scav Press

ConstScavPAngle = 5         'At bdc, start of compression, ExhVV and AIr VV are open, but
                            'Scav Air maintains const pressure for a few deg of crank
                            'NOTE: MISTAKE - THIS IS SAME POINT AS AIR V/V OPEN !

ScavPDecayRate = 0.01        'At start Compression after Air VV closes, P falls a little before
                            'Ecxh VV closes

'Graph Values
Graph1.YAxisStyle = 2
Graph1.YAxisMin = 0
Graph1.YAxisMax = 150
YaxisOffset = 0.004



TwoStroke = True   'for now

'set UpDowns with default values, thopse commented out can't be set
UpDown1(0).Value = CylLength
UpDown1(1).Value = bore
UpDown1(2).Value = stroke
UpDown1(3).Value = LengthConRod
'UpDown1(4).value = CompRatio
'UpDown1(5).Value = 75  'for test
UpDown1(5).Max = MaxEngRevs
UpDown1(5).Value = MaxEngRevs * 0.9  'for test

'UpDown1(6).value = MaxSpeed
'UpDown1(7).value = MaxPCompression
UpDown1(8).Value = ScavAirPressure * 100 'UPDOWNs cant increment decimals
UpDown1(9).Value = AirValveOpenAngle
UpDown1(10).Value = ExhVVOpenAngle
UpDown1(11).Value = InjectionAngle
UpDown1(12).Value = IgnitionDelay
UpDown1(13).Value = FuelInjPeriod + 1
UpDown1(14).Value = CetaneVal + 1
UpDown1(15).Value = SprayAngle + 1
'UpDown1(16).value = LinerTemp
'UpDown1(17).value = CylWallTemp
'UpDown1(18).value = PistonCrownTemp
'UpDown1(19).value = DiamInjParticle
'UpDown1(20).value = ScavEfficiency
'UpDown1(21).value = NOx

'12/6/98 disable all UpDown controls - except in Predictor Mode
For i = 0 To UpDown1.Count - 1
    UpDown1(i).Visible = False
Next

'exception is speed for testing
UpDown1(5).Visible = True


Call SetValues


'set default background colour & show in text box



End Sub
Sub Display_PCrankAngle()

Dim i As Integer
Dim ix As Integer
Dim Vcx As Single
Dim ExhVVOpenVol As Single

'Set up Graph
    Graph1.GraphTitle = "Pressure / Crank Angle)"
    Graph1.LeftTitle = "Pressure (bar)"
    Graph1.LeftTitleStyle = 1  'up
    Graph1.BottomTitle = "Crank Angle (degrees)"
    Graph1.GridStyle = 3
    Graph1.YAxisMax = 160
    Graph1.XAxisMin = 0
    Graph1.XAxisMax = 360
    Graph1.XAxisTicks = 4
    Graph1.NumSets = 1
    Graph1.NumPoints = STEPs * 2 'for 2-stroke
    Graph1.ThisSet = 1
    Graph1.ColorData = 0
    
    For i = 1 To Graph1.NumPoints
        Points(i) = 0#
        Graph1.xpos(i) = 0
    Next
    
    counter = 0
    

'Initial Calculations of: Min, Swept & Max Vols, Compression Ratio and Unit of Vol as proportion max Vol
    MinVol = pi * bore * bore * (CylLength - (throw * 2)) * 0.000000001 'unswept vol
    SweptVol = (pi * bore * bore * (throw * 2) * 0.000000001)
    MaxVol = SweptVol + MinVol
    CompRatio = MaxVol / MinVol
    UnitOfVol = SweptVol / STEPs
   

'get initial volume bdc, we know pressure is the Scav Pressure so calc K1
    angle = 0
    Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
    Call GetVol(bore, CylLength, throw, Position, Vc)
    K1 = ScavAirPressure * Exp(Log(Vc + MinVol) * Gamma) ' find gamma from P*(V**Gamma) = Constant (K1)
    

'Find The Volume when the Air V/V closes
    angle = AirValveOpenAngle 'Angle at which air v/v closes
    Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
    Call GetVol(bore, CylLength, throw, Position, Vc)
    
'How many STEPs to the air valve opening ?
'    Start = STEPs * 180 / angle
    
        
'Set the Pressure to ScavAirPress while air v/v is open
    For i = 1 To AirValveOpenAngle
        Points(i) = ScavAirPressure 'assume ambient pressure while air inlet v/v open
        xpos(i) = i
    Next
    Debug.Print "just set all points from 1 to " & AirValveOpenAngle & "   = ScavPress"


'Calculate the Point of Ignition before tdc
'Ideally, this will be at (Inj Angle - Inj Delay), both of whic are variable but
'may also be affected by faults, so add another variable with default value of zero

IgnitionAngle = 180 - (InjectionAngle - IgnitionDelay + IgnitionVar)




'Calc Pressure for each STEP for main compression stroke, up to tdc
    MaxPCompression = 0
    'For i = AirValveOpenAngle + 1 To STEPs
    For i = AirValveOpenAngle + 1 To IgnitionAngle
          Call GetPosition(i, throw, LengthConRod, PistonOffset, Position)
          Call GetVol(bore, CylLength, throw, Position, Vc)
          Vc = Vc + MinVol
          Pc = K1 / Exp(Log(Vc) * Gamma) 'calc the pressure at this volume
          Points(i) = Pc
       
          If Pc > MaxPCompression Then
                MaxPCompression = Pc
          End If
          xpos(i) = i
    Next
    counter = IgnitionAngle 'keep track of last angle
    
        Graph1.ColorData = 1 'set the plot colur (blue)
        Points(STEPs) = LastPressure 'remember the last pressure - we need it for the combustion part

        'For i = 1 To STEPs ' now store plot values for the compression stroke
                'Graph1.Data(i) = Points(i) 'graphdata starts at 1 not 0
                'Graph1.xpos(i) = i  'use angle from -180 to +180 as X axis value
                
        'Next

'============================ END OF COMPRESSION =====================================
    'GoTo Skip
    
    
'now do same again for power (called exhaust) stroke
'first few points are from Look-up table - decide here which one to use
    
    
    NumberOfPoints = 0
    
'find how many points there are to be plotted from the look-up table (+ve values only)
    For i = 1 To UBound(PCLookUp)
        If PCLookUp(i) > 0 Then
            NumberOfPoints = NumberOfPoints + 1
        End If
    Next
    
'calculate a multiplier for x-axis based on Engine speed and some constant
    SpeedMult = ((1 - (engrevs / MaxEngRevs)) * 0.8) ' slower means fatter
    
    For i = 1 To NumberOfPoints 'fill Xm array with modified x-axis values
        If i <= UBound(Xm) Then
            Xm(i + 1) = i + (i * SpeedMult) 'so for a Mult Factor of 0.8, X-axis values is at  1.8, 3.6,5.4 etc
        End If
    Next
    Xm(0) = 0
    
    
    Call CylVar
    
    NumberOfNewPoints = Xm(NumberOfPoints)
        
        For i = 1 To NumberOfPoints 'each i is a unit volume

            If PVLookUp(i) > 0 Then
                LastPressure = MaxPCompression + (engrevs / MaxEngRevs) * MaxPCompression * (PCLookUp(i) / 100)
                LastPressure = LastPressure * CylVariation
                
                LastPressure = LastPressure * Fudge1
                
                'LastAngle = CInt(Xm(i)) + 180
                LastAngle = CInt(Xm(i)) + IgnitionAngle
                'Points(STEPs + i) = LastPressure
                'xpos(STEPs + i) = (Xm(i) + 180)
                'counter = STEPs + i 'keep a track of how many points are assigned
                counter = counter + 1
                Points(counter) = LastPressure
                xpos(counter) = (Xm(i)) + IgnitionAngle
                
            Else
                Exit For
            End If
        Next
        'xpos(180) = 180
        'Points(180) = LastPressure

        Debug.Print "After PVLookup, lastAngle = " & LastAngle & "  Last Press = " & LastPressure
        Debug.Print "last counter value  = " & counter
        
       
        
        
        'calc a new k2 based on last vol and press
        Call GetPosition(LastAngle, throw, LengthConRod, PistonOffset, Position)
        Call GetVol(bore, CylLength, throw, Position, Vc)

        Vc = Vc + MinVol
        K2 = LastPressure * Exp(Log(Vc) * Gamma) '  P*(V**Gamma) = Constant (K1)
        

'Calculate P for each degree left during power stroke
        Gamma = 1.38
    
       
        
'calculte P for each unit of V up to Exh VV open
        For ix = LastAngle To ExhVVOpenAngle + 180
            'angle = CInt(ix)
            counter = counter + 1
            Call GetPosition(ix, throw, LengthConRod, PistonOffset, Position)
            Call GetVol(bore, CylLength, throw, Position, Vc)
            Vc = Vc + MinVol
            Pc = K2 / Exp(Log(Vc) * Gamma)
            'Points(ix) = Pc
            'xpos(ix) = ix
            Points(counter) = Pc
            xpos(counter) = ix

            'Debug.Print "in first PV calc, ix = " & ix & "Vc = " & Vc & "Pc = " & Pc
            

        Next

'now do the decay to Scav Air Press after the Exh V/V opens
        For ix = ExhVVOpenAngle + 181 To 360
            counter = counter + 1

            Call GetPosition(ix, throw, LengthConRod, PistonOffset, Position)
            Call GetVol(bore, CylLength, throw, Position, Vc)
            Vc = Vc + MinVol
            Pc = Pc - (((Pc - ScavAirPressure) * ExhVVDecayRate) / ScavAirPressure)
            'Points(ix) = Pc
            'xpos(ix) = ix
            Points(counter) = Pc
            xpos(counter) = ix


       Next
Skip:
        Graph1.NumPoints = counter
        Graph1.ThisSet = 1
        Graph1.ColorData = 1
        
        'For i = STEPs + 1 To (STEPs * 2) '181 to 360
        'For i = STEPs + 1 To counter '181 to 360
        For i = 1 To counter '1 to 360

            Graph1.Data(i) = Points(i)
            Graph1.xpos(i) = xpos(i)

        Next
'draw the graph
        Graph1.DrawMode = 3
                
        For i = 180 To 200
        
        Debug.Print " i = " & i & "  data(" & i & ") = " & Graph1.Data(i) & "  Position = " & Graph1.xpos(i)
    
        Next
        
End Sub

Sub SetDefaults()

'default values


Fudge1 = 1
CylVariation = 1

'default values for all Engine Dynamic Stuff


Graph1.NumSets = 1
Graph1.ThisSet = 1
Graph1.ColorData = 1

Graph1.YAxisStyle = 2
Graph1.YAxisMin = 0
Graph1.YAxisMax = 150
Graph1.Background = 78 'pale green & yellow text


'deselect all options
DrawCardOn = False
PistonSpeedOn = False
Stroke_CrankOn = False
P_VOn = False
P_CrankOn = False
NOxOn = False

AmbientAirPressure = 1

'disable UpDown buttons which are not used (ie no chance to alter)
UpDown1(4).Visible = False  'Comp ratio
UpDown1(6).Visible = False  'Max Piston Speed
UpDown1(7).Visible = False  'max Comp Pressure
UpDown1(16).Visible = False 'Liner Temp
UpDown1(17).Visible = False 'Cyl wall Temp
UpDown1(18).Visible = False 'Pist Crown Temp
UpDown1(19).Visible = False 'Diam of Inj Fuel
UpDown1(20).Visible = False 'Scav Effcy
UpDown1(21).Visible = False 'NOx

'set max & min &values for UpDown butto
UpDown1(0).Max = 5000
UpDown1(0).Min = 100 'CylLength

UpDown1(1).Max = 1000
UpDown1(1).Min = 100 'bore

UpDown1(2).Max = 5000
UpDown1(2).Min = 100 'stroke

UpDown1(3).Max = 10000
UpDown1(3).Min = 100 'LengthConRod

UpDown1(5).Max = 1000
UpDown1(5).Min = 1 ' engrevs

UpDown1(8).Max = 600
UpDown1(8).Min = 1 'ScavAirPressure * 100

UpDown1(9).Max = 180
UpDown1(9).Min = 1 ' AirValveOpenAngle

UpDown1(10).Max = 360
UpDown1(10).Min = 1 'ExhVVOpenAngle

UpDown1(11).Max = 30
UpDown1(11).Min = 1 ' InjectionAngle

UpDown1(12).Max = 10
UpDown1(12).Min = 1 'IgnitionDelay

UpDown1(13).Max = 500
UpDown1(13).Min = 1 'FuelInjPeriod

UpDown1(14).Max = 1000
UpDown1(14).Min = 1 ' CetaneVal

UpDown1(15).Max = 1000
UpDown1(15).Min = 1 'SprayAngle












Call SetUp6L90MC
Call SetValues
Option2(0).Value = True 'default to big slow speed

End Sub
Sub display_PV()



Dim i As Integer
Dim Vcx As Single
Dim ExhVVOpenVol As Single

'Set up Graph
    If Check2.Value Then
        Graph1.YAxisMax = 10 'light spring
    Else
        Graph1.YAxisMax = 160
    End If
    Graph1.XAxisMin = 0
    Graph1.XAxisMax = 1 'normalised

    Graph1.NumSets = 1
    Graph1.NumPoints = STEPs * 2 'for 2-stroke
    Graph1.ThisSet = 1
    Graph1.ColorData = 0
    
    
    

'Initial Calculations of: Min, Swept & Max Vols, Compression Ratio and Unit of Vol as proportion max Vol
    MinVol = pi * bore * bore * (CylLength - (throw * 2)) * 0.000000001 'unswept vol
    SweptVol = (pi * bore * bore * (throw * 2) * 0.000000001)
    MaxVol = SweptVol + MinVol
    CompRatio = MaxVol / MinVol
    UnitOfVol = SweptVol / STEPs
   

'get initial volume bdc, we know pressure is the Scav Pressure so calc K1
    angle = 0
    Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
    Call GetVol(bore, CylLength, throw, Position, Vc)
    K1 = ScavAirPressure * Exp(Log(Vc + MinVol) * Gamma) ' find gamma from P*(V**Gamma) = Constant (K1)
    

'Find The Volume when the Air V/V closes
    angle = AirValveOpenAngle 'Angle at which air v/v closes
    Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
    Call GetVol(bore, CylLength, throw, Position, Vc)
    
    
'how many units of normalised volume has the piston covered at this point ?
    Start = CInt((SweptVol - Vc) / UnitOfVol)
    AirValveOpenVol = Start 'remember it because we'll need it again later
    
'Set the Pressure to ScavAirPress while air v/v is open
    For i = 1 To Start
        Points(i) = ScavAirPressure 'assume ambient pressure while air inlet v/v open
    Next
  
  Debug.Print "AirVVCloses at " & Start & "  units of vol"
  
'Find The Volume when the Exh V/V closes
        angle = 180 - ExhVVOpenAngle 'Angle at which air v/v closes is same as angle at which it opens
        Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
        Call GetVol(bore, CylLength, throw, Position, Vc)
    
    
'how many units of normalised volume has the piston covered at this point ?
        ExhVVOpenVol = CInt((SweptVol - Vc) / UnitOfVol)
        Debug.Print "ExhVVCloses at " & ExhVVOpenVol & "  units of vol"

'Set the Pressure to fall slowly towards ambient after Air V/V shuts and before EXh VV shuts
    Pc = ScavAirPressure
    For i = Start + 1 To ExhVVOpenVol
            Pc = Pc - (((Pc - AmbientAirPressure) * ScavPDecayRate) / AmbientAirPressure)
            'Pc = Pc - (((Pc - ScavAirPressure) * ExhVVDecayRate) / ScavAirPressure)
            Points(i) = Pc 'assume ambient pressure while air inlet v/v open
    Next
  
'get a new K
    K1 = Pc * Exp(Log(Vc + MinVol) * Gamma) ' find gamma from P*(V**Gamma) = Constant (K1)

  

'Calc Pressure for each unit of Volume for main compression stroke, up to tdc
    MaxPCompression = 0
    'For i = Start + 1 To STEPs
    For i = ExhVVOpenVol + 1 To STEPs
            'If Check2.value Then 'it is a light spring so cut off pressures at Exh VV close
            '    Points(i) = Pc
            'Else
                Vc = ((STEPs - i) * UnitOfVol) + MinVol 'vol left decreases with number of STEPs
                Pc = K1 / Exp(Log(Vc) * Gamma) 'calc the pressure at this volume
                Points(i) = Pc
       
                If Pc > MaxPCompression Then
                    MaxPCompression = Pc
                End If
            'End If
            
        Next
        
        Graph1.ColorData = 1 'set the plot colur (blue)
        Points(STEPs) = LastPressure 'remember the last pressure - we need it for the combustion part

        For i = 1 To STEPs ' now store plot values for the compression stroke
            If Check1.Value = 1 Then
                Graph1.Data(i) = Points(STEPs + 1 - i) 'graphdata starts at 1 not 0
            Else
                 Graph1.Data(i) = Points(i) 'graphdata starts at 1 not 0
           End If
        
            Graph1.xpos(i) = ((STEPs - i) / STEPs) + YaxisOffset 'normalised volume starting with pos(1) at 1.0

        Next

'============================ END OF COMPRESSION =====================================
    
    
'now do same again for power (called exhaust) stroke
'first few points are from Look-up table - decide here which one to use
    
    
    NumberOfPoints = 0
    
'find how many points there are to be plotted from the look-up table (+ve values only)
    For i = 1 To UBound(PVLookUp)
        If PVLookUp(i) > 0 Then
            NumberOfPoints = NumberOfPoints + 1
        End If
    Next
    
'calculate a multiplier for x-axis based on Engine speed and some constant
    SpeedMult = ((1 - (engrevs / MaxEngRevs)) * 0.8) ' slower means fatter
    Xm(1) = YaxisOffset 'make first point close to but not on Y-axis so line appears continuous

    For i = 1 To NumberOfPoints 'fill Xm array with modified x-axis values
        If i <= UBound(Xm) Then
            Xm(i + 1) = i + (i * SpeedMult) 'so for a Mult Factor of 0.8, X-axis values is at  1.8, 3.6,5.4 etc
        End If
    Next
    Xm(0) = 0
    
    
    
    NumberOfNewPoints = Xm(NumberOfPoints)
        
        For i = 1 To NumberOfPoints 'each i is a unit volume

            If PVLookUp(i) > 0 Then
                LastPressure = MaxPCompression + (engrevs / MaxEngRevs) * MaxPCompression * (PVLookUp(i) / 100)
                LastAngle = Xm(i)
                Points(STEPs + i) = LastPressure
                xpos(STEPs + i) = (Xm(i) / STEPs)
            Else
                Exit For
            End If
        Next
    
        Vc = (LastAngle * UnitOfVol) + MinVol
        K2 = LastPressure * Exp(Log(Vc) * Gamma) '  P*(V**Gamma) = Constant (K1)
        
'how many units of normalised volume is this ?
        Start = LastAngle
        If CInt(LastAngle) > Int(LastAngle) Then 'after decimal point is > 0.5 so stsrt with cint
            Start = CInt(LastAngle)
        Else
            Start = CInt(LastAngle) + 1
        End If
     
     
        Debug.Print "End of Ignition bit is at start = " & Start

'Calculate P for each STEP during power stroke
        Gamma = 1.38
        counter = NumberOfPoints
    
        'Debug.Print "start = " & Start & " NumberOfPoints = " & NumberOfPoints
        incr = (STEPs - Start) / (STEPs - NumberOfPoints)
        
'Find The Volume when the Exh V/V closes
        'angle = ExhVVOpenAngle 'Angle at which air v/v closes
        'Call GetPosition(angle, throw, LengthConRod, PistonOffset, Position)
        'Call GetVol(bore, CylLength, throw, Position, Vc)
    
    
'how many units of normalised volume has the piston covered at this point ?
        'ExhVVOpenVol = CInt((SweptVol - Vc) / UnitOfVol)

        'Debug.Print "ExhVVOpenVol = " & ExhVVOpenVol

        
'calculte P for each unit of V up to Exh VV open
        For ix = Start To STEPs - ExhVVOpenVol Step incr

            counter = counter + 1
            Vc = (ix * UnitOfVol) + MinVol
            Pc = K2 / Exp(Log(Vc) * Gamma)
            Points(STEPs + counter) = Pc
            'xpos(counter) = ix
            xpos(STEPs + counter) = (ix / STEPs)

            'Points(counter + 180) = Pc
            'xpos(counter + 180) = ix

        Next

'now do the decay to Scav Air Press after the Exh V/V opens until Air V/V open

        'For ix = ExhVVOpenVol + 1 To STEPs STEP incr
        For ix = STEPs - ExhVVOpenVol + 1 To STEPs - AirValveOpenVol Step incr

            counter = counter + 1
            'Pc = Pc - (((Pc - (ScavAirPressure)) * ExhVVDecayRate) / ScavAirPressure)
            
            Pc = Pc - (((Pc - (ScavAirPressure * MinPressProportion)) * ExhVVDecayRate) / (ScavAirPressure * MinPressProportion))
            
            Points(STEPs + counter) = Pc
            xpos(STEPs + counter) = (ix / STEPs)
       Next
        Debug.Print " At End of POwer, AirVV Opens at : " & AirValveOpenVol
        Debug.Print " At End of POwer, ExhVV Opens at : " & ExhVVOpenVol
   
'now do the rise to Scav Air Press after the Air V/V open to bdc

        'For ix = ExhVVOpenVol + 1 To STEPs STEP incr
        For ix = STEPs - AirValveOpenVol + 1 To STEPs Step incr

            counter = counter + 1
            Pc = Pc + (((ScavAirPressure - Pc) * ScavInletRiseRate) / ScavAirPressure)
            
            Points(STEPs + counter) = Pc
            xpos(STEPs + counter) = (ix / STEPs)
       Next
        
        
        
        Graph1.ThisSet = 1
        Graph1.ColorData = 1

        For i = STEPs + 1 To (STEPs * 2) '181 to 360

            If Check1.Value = 0 Then
                Graph1.Data(i) = Points(i)
           Else
                Graph1.Data(i) = Points((STEPs * 2) + 1 - i)
            End If
            Graph1.xpos(i) = xpos(i)
            
        Next
        
'draw the graph
        Graph1.DrawMode = 3
        
        
End Sub
Sub Display_StrokeVsCrank()
Dim i As Integer


        Graph1.NumPoints = 360
        Graph1.YAxisStyle = 0 'Auto Scale as this won't change much
        Graph1.GraphTitle = "Piston Stroke vs Crank Angle"
        Graph1.LeftTitle = "Piston Stroke (mm)"
        Graph1.LeftTitleStyle = 1  'up
        Graph1.BottomTitle = "Crank Angle (deg)"
        Graph1.XAxisMin = 0
        Graph1.XAxisMax = 360

        Graph1.YAxisTicks = 10
        Graph1.XAxisTicks = 8
        Label2(6).Visible = False ' Hide Max Piston Speed as this is not calculated here
        Text1(6).Visible = False
        

        For i = 1 To 360
            Call GetPosition(i, throw, LengthConRod, PistonOffset, Position)
            '13/5/98 Graph1.Data(i) = Position
            '3/6/98 Graph1.Data(i) = Position - throw
            Graph1.Data(i) = Position

            Graph1.xpos(i) = i
        Next
        Graph1.DrawMode = 3

End Sub
Sub Display_PistonSpeed()
Dim i As Integer

    Graph1.YAxisStyle = 2
    Graph1.YAxisMin = -20
    Graph1.YAxisMax = 20
    Graph1.YAxisTicks = 4
    Graph1.XAxisMin = 0
    Graph1.XAxisMax = 360

    Graph1.GraphTitle = "Piston Speed vs Crank Angle"
    Graph1.LeftTitle = "Piston Speed (m/s)"
    Graph1.LeftTitleStyle = 1  'up
    Graph1.BottomTitle = "Crank Angle (deg)"
    Graph1.XAxisTicks = 4
    Graph1.NumPoints = 360

    Label2(6).Visible = True ' Show Max Piston Speed
    Text1(6).Visible = True
    
        
        If engrevs >= 10 Then
            If Not Graph1.Visible Then 'don't make it visible every pass as this flickers
                Graph1.Visible = True
                Label1.Visible = False
            End If

            OneDegTime = 60 / (engrevs * 360) ' time in secs to turn 1 degree
            
            MaxSpeed = 0
            For i = 1 To 360
                Call GetPosition(i, throw, LengthConRod, PistonOffset, Position)
                Pos1 = Position
                Call GetPosition(i + 1, throw, LengthConRod, PistonOffset, Position)
                Pos2 = Position
               
                PistSpeed(i) = (Pos2 - Pos1) / (OneDegTime * 1000) 'convert mm/s to m/s

                Graph1.Data(i) = PistSpeed(i)
                Graph1.xpos(i) = i
                
                If PistSpeed(i) > MaxSpeed Then
                    MaxSpeed = PistSpeed(i)
                End If
            Next
            Graph1.DrawMode = 3 'Update using bitmap to remove flicker
            'Text1(6).Text = Format(MaxSpeed, "###.00")         'Display Max Speed
        Else
            Graph1.Visible = False
            Label1.Visible = True
        End If
        'Debug.Print "OneDegTime = " & OneDegTime
    Debug.Print "In Disp Piston Speed, Throw = " & throw & " LenConRod = " & LengthConRod & "Max Speed = " & MaxSpeed

End Sub

Sub Display_Values(CylLength As Single, bore As Single, throw As Single, _
                   LengthConRod As Single, CompRatio As Single, engrevs As Single, _
                   MaxSpeed As Single, MaxPCompression As Single)
            
            Text1(0).Text = CylLength                           'Display Cylinder Length
            Text1(1).Text = bore                                'Display Bore
            Text1(2).Text = Format(throw * 2, "#####.00")       'Display Stroke (Throw * 2)
            Text1(3).Text = LengthConRod                        'Display Length Con Rod
            Text1(4).Text = Format(CompRatio, "###.00")         'Display Compression Ratio
            Text1(5).Text = Format(engrevs, "###.00")           'Display Engine RPM
            Text1(6).Text = Format(MaxSpeed, "###.00")          'Display Max Speed
            Text1(7).Text = Format(MaxPCompression, "###.00")          'Display Max Speed

End Sub
Sub Dynamic_Displays()
Dim i As Integer
Dim counter As Integer







'stroke = 3001
Graph1.NumSets = 1
Graph1.ThisSet = 1
Graph1.ColorData = 1
'temp values for this piston

'engrevs = Slider1.value * 1.001  ' For test Only !!!
'engrevs = VB_ENGINE(92)

'Calculate Compression Ratio (not needed for Graph)
MinVol = pi * bore * bore * (CylLength - (throw * 2)) * 0.000000001 'unswept vol
SweptVol = (pi * bore * bore * (throw * 2) * 0.000000001)
MaxVol = SweptVol + MinVol
CompRatio = MaxVol / MinVol




If PistonSpeedOn Then 'display Piston Speed vs Crank Angle
    Debug.Print " calling Piston Speed"
    Call Display_PistonSpeed
       
        
ElseIf Stroke_CrankOn Then
    
    Call Display_StrokeVsCrank
    
ElseIf P_VOn Or P_CrankOn Then

    Label2(6).Visible = False ' Hide Max Piston Speed as this is not calculated here
    Text1(6).Visible = False


    'STEPs = 360
    'Graph1.NumSets = 1
    Graph1.ThisSet = 1
    Graph1.ColorData = 1
    Graph1.YAxisStyle = 2
    Graph1.YAxisMin = 0
    Graph1.YAxisMax = 80
    Graph1.YAxisTicks = 8

    Graph1.NumSets = 1



    If P_VOn Then   'Plot P against normalised V for compression stroke
    
        Graph1.GraphTitle = "Pressure / Volume)"
        Graph1.LeftTitle = "Pressure (bar)"
        Graph1.LeftTitleStyle = 1  'up
        Graph1.BottomTitle = "Volume (normalised)"

        Graph1.XAxisTicks = 10
        Graph1.GridStyle = 3

        If Check2.Value Then
            Call Display_PV_LightSpring
        Else
            Debug.Print "calling display_PV from dynamic displays"

            Call display_PV
        End If
        
    ElseIf P_CrankOn Then 'Plot P against angle for compression stroke

        Call Display_PCrankAngle
    End If
    
ElseIf NOxOn Then
    Call Display_NOx
End If

'Display Engine Data In Text Boxes:
Call Display_Values(CylLength, bore, throw, LengthConRod, CompRatio, engrevs, MaxSpeed, MaxPCompression)

End Sub



Sub GetPVConstant(Pressure As Single, volume As Single, Gamma As Single, K As Single)
'calculates the constant K given by P*V**gamma = Constant from a known Vol & Pressure

K = Pressure * Exp(Log(volume) * Gamma)        'P*(V**Gamma) = Constant (K2)


End Sub
Sub ArcSin(x As Single, ArcSin_X As Single)
'calculates the inverse sine (ArcSin) of an angle X (in Radians) using Atn function (ArcSIn not available in VB5)

    ArcSin_X = Atn(x / Sqr(-x * x + 1))
    
End Sub
Sub GetVol(bore As Single, CylLength As Single, throw As Single, Position As Single, volume As Single)
'Dim UnsweptLength As Single

'calculate volume under piston at any POSITION (mm) of the CYLLENGTH (mm).
'Assume Perfect Cylindrical Volume with BORE as radius (mm).
'return vol in cubic metres

'<---------------CYL LENGTH----------------------->
'TDC 180 deg                                      >BDC (0 deg)
' un-  | Limit of Travel
'swept <---------------- throw * 2 --------------->
'                    <--------Position------------>
'      <--Vol-------->
'<---------------Max Vol ------------------------->
'      <---------------Swept Vol ----------------->

If Position <= CylLength Then
    'Volume = pi * bore * bore * (stroke - Position) * 0.000001 'in m3
    'UnsweptLength = CylLength - (throw * 2)
    'Volume = pi * bore * bore * (UnsweptLength + Position) * 0.000000001 'in m3
    'volume = pi * bore * bore * (CylLength - Position) * 0.000000001 'in m3
     volume = pi * bore * bore * ((throw * 2) - Position) * 0.000000001 'in m3
   
End If

End Sub
Sub GetNormalisedVolume(bore As Single, CylLength As Single, throw As Single, volume As Single, NormalisedVolume As Single)
Dim MaxVol As Single

MaxVol = pi * bore * bore * (throw * 2) * 0.000000001

NormalisedVolume = volume / MaxVol
'Debug.Print "In GNV: vol= " & volume & "maxvol = " & MaxVol & " NV = " & NormalisedVolume
End Sub
Sub GetPistonSpeed(stroke As Single, engrevs As Single, PistonSpeed As Single)

End Sub


Sub GetPosition(angle As Integer, throw As Single, LengthConRod As Single, _
                PistonOffset As Single, Position As Single)
'calculates the position of the piston within the cylinder, starting at bottom when Angle = 0
'Throw (mm)is radius of cam and = Stroke / 2
'LengthConRod (mm) is length of ConRod
'PistonOffset (mm) is the height of cylinder above cam
'Angle in degrees

Dim AngInRad As Single
Dim r As Single
Dim Lc As Single
Dim ArcSin_X As Single
Dim L As Single
Dim ConAngle As Integer


Lc = LengthConRod
r = throw
    'ConAngle = 180 - angle  'the ConShaft angle used in the Position calculation
                            'is the compliment of the crank angle
    '15/5/98 make it +ve
    ConAngle = Abs(180 - angle)  'the ConShaft angle used in the Position calculation

    AngInRad = r * Sin(ConAngle * pi / 180) / Lc 'calculate r.SinA/Lc with angle in radians
    
    Call ArcSin(AngInRad, ArcSin_X)
    
    L = Sqr((Lc * Lc) + (r * r) - (2 * Lc * r * Cos((pi - (ConAngle * pi / 180) - ArcSin_X))))
    Position = L - PistonOffset

End Sub

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub Command2_Click()
    CommonDialog1.HelpFile = "C:\haven\help\NITROGEN OXIDES.HLP"
    CommonDialog1.HelpCommand = cdlHelpContents
    CommonDialog1.ShowHelp
End Sub

Private Sub Command3_Click()
    CommonDialog1.HelpFile = "C:\haven\help\SPRAY.HLP"
    CommonDialog1.HelpCommand = cdlHelpContents
    CommonDialog1.ShowHelp

End Sub

Private Sub Command4_Click()
LUTableEdit.Show

End Sub

Private Sub Form_Load()
'Dim resp As Integer
Dim ErrorFlag As Byte
Dim filenum As Integer
Dim CharStr As String
Dim i As Integer
Dim resp As Integer

    

   Call GetLanguageStrings("Plot2.txt", ErrorFlag)

    If Not ErrorFlag Then
        frm_Plot2.Caption = Language_Str(1) 'Engine Dynamics
        Option1(0).Caption = Language_Str(2) 'Piston Speed
        Option1(1).Caption = Language_Str(3) 'Stroke-Crank Angle
        Option1(2).Caption = Language_Str(4) 'P-V
        Option1(3).Caption = Language_Str(5) 'P-Crank Angle
        Option1(4).Caption = Language_Str(6) 'NOx
        Option1(5).Caption = Language_Str(7) 'DISABLED FOR NOW dP-Crank Angle

        Option1(6).Caption = Language_Str(8) 'DISABLED FOR NOW dHeat-Crank Angle
        Check1.Caption = Language_Str(9) 'DISABLED FOR NOW Reverse X-Axis
        Check2.Caption = Language_Str(10) 'DISABLED FOR NOW Light Spring
        Command1.Caption = Language_Str(11) 'Exit
        Command4.Caption = Language_Str(12) 'Edit LU Table - DIABLE FOR CUSTOMER VERSION
        Label1.Caption = Language_Str(13)   'Engine Speed Must Exceed 10 RPM To Plot Graph
                                            'This label is behind graph
        For i = 0 To Label2.Count - 1       'Label2 s are for Frame 1, thre are 24 of them from Label1(0) to (23)
            Label2(i).Caption = Language_Str(14 + i)
        Next
       
        Frame1.Caption = Language_Str(38) 'Engine Data
        Frame2.Caption = Language_Str(39) 'Plot Type
        Frame3.Caption = Language_Str(40) 'Graph Point Data
        Frame4.Caption = Language_Str(41) 'Engine Type
        
        For i = 0 To Option2.Count - 1      ' Option 2 is for Engine Model, thre are 6 from (0) to (5)
           Option2(i).Caption = Language_Str(42 + i)
        Next


    End If
    
Graph1.Hot = 1
'get Pressure-Vol Look-Up values

' 19/1/99
Graph1.SDKPaint = 1

filenum = FreeFile
'On Error GoTo ErrHandler1

'Read in LU table values for P-V diagram
Open "C:\HAVEN\LU_PV.dat" For Input Access Read As #filenum

For i = LBound(PVLookUp) To UBound(PVLookUp)
    Input #filenum, CharStr
    PVLookUp(i) = CSng(CharStr)
Next
    
    
Close #filenum

'On Error GoTo ErrHandler2

'Read in LU table values for P-Crank Angle diagram
Open "C:\HAVEN\LU_PC.dat" For Input Access Read As #filenum

For i = LBound(PCLookUp) To UBound(PCLookUp)
    Input #filenum, CharStr
    PCLookUp(i) = CSng(CharStr)
Next
    
    
Close #filenum

Call SetDefaults
Exit Sub

ErrHandler1:
resp = MsgBox("The file LU_PV.dat is missing or corrupted - please reload", vbOKOnly, "File Error")
Exit Sub
Errhandler2:
resp = MsgBox("The file LU_PC.dat is missing or corrupted - please reload", vbOKOnly, "File Error")


End Sub

Private Sub Form_Resize()
'11/5/98 skip until we have resizing sorted
Exit Sub


End Sub

Private Sub Graph1_HotHit(hitset As Integer, hitpoint As Integer)
MsgBox "Set" + Str$(hitset) + ", Point" + Str$(hitpoint) + ", value =" + Format$(GSDataGetAmp(hitpoint - 1, hitset - 1), "0.00"), 64, "Hot-Graph"
'Text1(8).Text = Str$(hitpoint) + ", value =" + Format$(GSDataGetAmp(hitpoint - 1, hitset - 1), "0.00")
'Text1(9).Text = Str$(hitpoint) + ", value =" + Format$(GSDataGetAmp(hitpoint - 1, hitset - 1), "0.00")
 
End Sub

Private Sub Graph1_SDKPaint()

' 19/1/99
Dim XLen As Single
Dim YLen As Single
Dim XOrg As Single
Dim YOrg As Single

Dim Yscale As Single
Dim Xscale As Single
Dim x As Single
Dim Y As Single

Dim Nox_ppm As Single

Dim symbol_colour As Integer

'symbol_colour = &HFFFFFF
symbol_colour = 4


XLen = Graph1.SDKInfo(5)
YLen = Graph1.SDKInfo(6)
XOrg = Graph1.SDKInfo(7)
YOrg = Graph1.SDKInfo(8)


Debug.Print "SDKPaint"

If Option1(4).Value Then    ' Nox plot is selected
    If (Graph1.SDKInfo(3) - Graph1.SDKInfo(4)) > 0 And (Graph1.SDKInfo(1) - Graph1.SDKInfo(2)) > 0 Then

        Yscale = Graph1.SDKInfo(6) / (Graph1.SDKInfo(3) - Graph1.SDKInfo(4))
        Xscale = Graph1.SDKInfo(5) / (Graph1.SDKInfo(1) - Graph1.SDKInfo(2))

      '     Y = YOrg + Yscale * HScroll2.Value
      '     X = XOrg + Xscale * HScroll1.Value
      
            
        If engrevs > 10# Then
            x = XOrg + Xscale * engrevs
            
            Nox_ppm = -0.005 * engrevs + 13#
            Y = YOrg + Yscale * Nox_ppm
        Else
            x = XOrg
            Y = YOrg
        End If

        resp = GSSymbol(x, Y, 13, symbol_colour)
    
        Debug.Print Graph1.SDKInfo(3), Graph1.SDKInfo(4)
   
    End If
End If



End Sub

Private Sub Option1_Click(index As Integer)


PistonSpeedOn = False
Stroke_CrankOn = False
P_VOn = False
P_CrankOn = False
DrawCardOn = False

Check1.Visible = False 'No need to offer Reverse X-Axis except for P-V
Check2.Visible = False 'No need to offer Light Spring except for P-V

Select Case index

Case 0 'Piston Speed
        PistonSpeedOn = True

Case 1 'Stroke - Crank Angle
        Stroke_CrankOn = True

Case 2 'P-V
        P_VOn = True
        'Check1.Visible = True 'Reverse X-Axis option
        'Check2.Visible = True 'Light Spring Option

        
Case 3 'P-Crank Angle
        P_CrankOn = True
        
Case 4 'NOx
        NOxOn = True

Case 5 'dP-Crank Angle
    
Case 6 'dHeat-Crank Angle

End Select


End Sub

Private Sub Option2_Click(index As Integer)
Select Case index

Case 0 'B&W6L90MC
    Call SetUp6L90MC
Case 1
    Call SetUp6LM60
Case 2
    Call SetUpVasa32
Case 3
    Call SetUpSulzerZA40
Case 4
    Call SetUp6LM60_P
Case 5
    'Call SetUpVasa32
End Select





End Sub

Private Sub Timer1_Timer()
'If DrawCardOn Then
    'Call DrawCard
'ElseIf PistonSpeedOn Then
    'Call PistonSpeed
'End If
Debug.Print "calling Dynamic display, CylLength = " & CylLength

Call Dynamic_Displays
Call SetValues 'update values in text boxes
End Sub

Private Sub UpDown1_Change(index As Integer)
Select Case index

Case 0
    CylLength = UpDown1(index).Value
    Debug.Print "setting cylLength = " & CylLength & " In Up_Down Change"
    
Case 1
    bore = UpDown1(index).Value
Case 2
    stroke = UpDown1(index).Value
Case 3
    LengthConRod = UpDown1(index).Value
Case 5
    engrevs = UpDown1(index).Value
Case 8
    ScavAirPressure = UpDown1(index).Value / 100
Case 9
    AirValveOpenAngle = UpDown1(index).Value
Case 10
    ExhVVOpenAngle = UpDown1(index).Value
Case 11
    InjectionAngle = UpDown1(index).Value
Case 12
    IgnitionDelay = UpDown1(index).Value
Case 13
    FuelInjPeriod = UpDown1(index).Value
Case 14
    CetaneVal = UpDown1(index).Value
Case 15
    SprayAngle = UpDown1(index).Value
End Select

End Sub

Private Sub UpDown2_Change()
    CylNumber = UpDown2.Value
    
End Sub
