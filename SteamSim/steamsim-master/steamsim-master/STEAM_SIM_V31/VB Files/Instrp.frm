VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "Comct232.OCX"
Begin VB.Form Instrp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructor Parameters"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7950
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   6180
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6060
      Top             =   6180
   End
   Begin VB.Frame frame1 
      Caption         =   "Bridge"
      Height          =   1335
      Index           =   5
      Left            =   3840
      TabIndex        =   75
      Top             =   5040
      Width           =   3975
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   18
         Left            =   2640
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "7"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   18
         Left            =   3000
         TabIndex        =   77
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Parameter(18)"
         BuddyDispid     =   196610
         BuddyIndex      =   18
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "mins"
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   79
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bridge Auto accelleration time"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   78
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Diesel Alternator Governors"
      Height          =   1335
      Index           =   4
      Left            =   0
      TabIndex        =   66
      Top             =   5040
      Width           =   3735
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   16
         Left            =   2280
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   17
         Left            =   2280
         TabIndex        =   67
         Text            =   "2"
         Top             =   840
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   16
         Left            =   2640
         TabIndex        =   69
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(16)"
         BuddyDispid     =   196610
         BuddyIndex      =   16
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   17
         Left            =   2640
         TabIndex        =   70
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(17)"
         BuddyDispid     =   196610
         BuddyIndex      =   17
         OrigLeft        =   2640
         OrigTop         =   840
         OrigRight       =   2880
         OrigBottom      =   1215
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Diesel Alternator 1 Droop"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Diesel Alternator 2 Droop"
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   73
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   72
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   71
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Main Engine Governor "
      Height          =   2895
      Index           =   3
      Left            =   3840
      TabIndex        =   45
      Top             =   2040
      Width           =   3975
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   15
         Left            =   2640
         TabIndex        =   61
         Text            =   "3"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   14
         Left            =   2640
         TabIndex        =   60
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   13
         Left            =   2640
         TabIndex        =   48
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   12
         Left            =   2640
         TabIndex        =   47
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   11
         Left            =   2640
         TabIndex        =   46
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   11
         Left            =   3000
         TabIndex        =   49
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(11)"
         BuddyDispid     =   196610
         BuddyIndex      =   11
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   12
         Left            =   3000
         TabIndex        =   50
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(12)"
         BuddyDispid     =   196610
         BuddyIndex      =   12
         OrigLeft        =   2640
         OrigTop         =   840
         OrigRight       =   2880
         OrigBottom      =   1215
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   13
         Left            =   3000
         TabIndex        =   51
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(13)"
         BuddyDispid     =   196610
         BuddyIndex      =   13
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   14
         Left            =   3000
         TabIndex        =   62
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(14)"
         BuddyDispid     =   196610
         BuddyIndex      =   14
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   15
         Left            =   3000
         TabIndex        =   63
         Top             =   2280
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(15)"
         BuddyDispid     =   196610
         BuddyIndex      =   15
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   13
         Left            =   3360
         TabIndex        =   65
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   12
         Left            =   3360
         TabIndex        =   64
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Droop"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   59
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Scavenge Air High Limit (2.15 bar)"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   58
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   11
         Left            =   3360
         TabIndex        =   57
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Scavenge Air Low Limit (0.5 bar)"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   55
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   9
         Left            =   3360
         TabIndex        =   54
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "High Limit(810 rpm)"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   53
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Low Limit (590 rpm)"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   52
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Fuel"
      Height          =   1935
      Index           =   2
      Left            =   3840
      TabIndex        =   32
      Top             =   0
      Width           =   3975
      Begin VB.TextBox CALVAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   81
         Top             =   480
         Width           =   1155
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   10
         Left            =   2640
         TabIndex        =   35
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   9
         Left            =   2640
         TabIndex        =   34
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   8
         Left            =   2640
         TabIndex        =   33
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   8
         Left            =   3000
         TabIndex        =   36
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(8)"
         BuddyDispid     =   196610
         BuddyIndex      =   8
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   9
         Left            =   3000
         TabIndex        =   37
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(9)"
         BuddyDispid     =   196610
         BuddyIndex      =   9
         OrigLeft        =   2640
         OrigTop         =   840
         OrigRight       =   2880
         OrigBottom      =   1215
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   10
         Left            =   3000
         TabIndex        =   38
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(10)"
         BuddyDispid     =   196610
         BuddyIndex      =   10
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Calorific Value"
         Height          =   195
         Left            =   240
         TabIndex        =   80
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ash"
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   44
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sulphur"
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   42
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   41
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Water"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   40
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   39
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Ship and Environment "
      Height          =   2895
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   2040
      Width           =   3735
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   7
         Left            =   2280
         TabIndex        =   28
         Text            =   "0"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   27
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   5
         Left            =   2280
         TabIndex        =   26
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   25
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   15
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(3)"
         BuddyDispid     =   196610
         BuddyIndex      =   3
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   735
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   24
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(4)"
         BuddyDispid     =   196610
         BuddyIndex      =   4
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   735
         Min             =   -10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   29
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(5)"
         BuddyDispid     =   196610
         BuddyIndex      =   5
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   735
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   6
         Left            =   2640
         TabIndex        =   30
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(6)"
         BuddyDispid     =   196610
         BuddyIndex      =   6
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   735
         Min             =   -10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   31
         Top             =   2280
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(7)"
         BuddyDispid     =   196610
         BuddyIndex      =   7
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   735
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fish Load"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sea State"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Hull Fouling"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Wind"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ballast / Draft"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Ambient Temperatures"
      Height          =   1935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   10
         Text            =   "3"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Parameter(0)"
         BuddyDispid     =   196610
         BuddyIndex      =   0
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   1
         Left            =   3420
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Parameter(1)"
         BuddyDispid     =   196610
         BuddyIndex      =   1
         OrigLeft        =   2640
         OrigTop         =   840
         OrigRight       =   2880
         OrigBottom      =   1215
         Max             =   40
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   11
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(2)"
         BuddyDispid     =   196610
         BuddyIndex      =   2
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   40
         Min             =   -20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "deg C"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   12
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Outside AirTemperature"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "deg C"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "deg C"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "ER Ambient AirTemperature"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sea Temperature"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   900
      TabIndex        =   82
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "Instrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim instrp_values(30) As Single
Private set_dummy_focus As Integer


Private Sub Form_Load()

' 24/7/98
Timer1.Enabled = False
Call update_values

set_dummy_focus = True

End Sub

Private Sub Timer1_Timer()

'Parameter(0).Text = Format(instrp_values(0), "#0") & Chr(176) & "C"
   Call update_values
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()

' 24/7/98
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

Parameter(1).Text = Format(VB_ENGINE(81), "##0.0")   '  Tamb_steam

CALVAL.Text = Format(VB_ENGINE(127), "###0.000") & " MJ/kG"

End Sub

Private Sub UpDown1_Change(index As Integer)

' 24/7/98
Call set_instrp(index + 1, Val(Parameter(index)))


Timer1.Enabled = True

set_dummy_focus = True


End Sub
Sub update_values()
Dim i As Integer

' 17/7/98
Parameter(0).Text = Int(VB_ENGINE(83))     '  Tsea
''' Moved to timer2...
'''Parameter(1).Text = Format(VB_engine(81), "##0.0")   '  Tamb_steam
Parameter(2).Text = Int(VB_ENGINE(84))     '  Toutside

For i = 3 To 18
  Parameter(i).Text = Int(VB_ENGINE(108 + i))  ' HULL etc etc
Next i

' SLOW SPEED....
'Parameter(3).Text = Int(LSS_reals(557))    ' HULL
'Parameter(4).Text = Int(LSS_reals(558))    ' WIND
'Parameter(5).Text = Int(LSS_reals(559))    ' BALLAST
'Parameter(6).Text = Int(LSS_reals(560))    ' SEA STATE
'Parameter(7).Text = Int(LSS_reals(725))    ' DRAFT ( 7 to 17 metres)
'
'Parameter(8).Text = Int(LSS_reals(84))    ' ASH
'Parameter(9).Text = Int(LSS_reals(85))    ' SULPHUR
'Parameter(10).Text = Int(LSS_reals(563))  ' HO Bunker Water
'
'Parameter(11).Text = Int(LSS_reals(319))   '  GOV_LOW_LIM
'Parameter(12).Text = Int(LSS_reals(318))   '  GOV_HIGH_LIM
'Parameter(13).Text = Int(LSS_reals(327))   '  GOV_PSCAV_LOLIM
'Parameter(14).Text = Int(LSS_reals(326))   '  GOV_PSCAV_HILIM
'
'Parameter(15).Text = Int(LSS_reals(568))    '  GOV_DROOP
'
'Parameter(16).Text = Int(LSS_reals(565))    '  DROOP_percent(1)
'Parameter(17).Text = Int(LSS_reals(566))    '  DROOP_percent(2)
'
'Parameter(18).Text = Int(LSS_reals(564))    ' TACCEL


End Sub
