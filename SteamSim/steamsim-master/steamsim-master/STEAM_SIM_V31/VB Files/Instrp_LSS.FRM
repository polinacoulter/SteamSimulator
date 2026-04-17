VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form Instrp_LSS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructor Parameters"
   ClientHeight    =   11475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11475
   ScaleWidth      =   18810
   Begin VB.Frame Frame12 
      Caption         =   "Constants"
      Height          =   1155
      Left            =   0
      TabIndex        =   237
      Top             =   10200
      Width           =   3735
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Propellor Pitch"
         Height          =   255
         Index           =   48
         Left            =   480
         TabIndex        =   243
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ship Inertia Multiplier"
         Height          =   255
         Index           =   47
         Left            =   480
         TabIndex        =   242
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Shaft Inertia Multiplier"
         Height          =   255
         Index           =   46
         Left            =   480
         TabIndex        =   241
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label CONSTANT 
         Alignment       =   2  'Center
         Caption         =   "C3"
         Height          =   255
         Index           =   3
         Left            =   2100
         TabIndex        =   240
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label CONSTANT 
         Alignment       =   2  'Center
         Caption         =   "C2"
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   239
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label CONSTANT 
         Alignment       =   2  'Center
         Caption         =   "C1"
         Height          =   255
         Index           =   1
         Left            =   2100
         TabIndex        =   238
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "LO CoolerThermostat"
      Height          =   915
      Left            =   0
      TabIndex        =   233
      Top             =   9240
      Width           =   3735
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   43
         Left            =   2520
         TabIndex        =   234
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "TLO_SET_F"
         BuddyDispid     =   196613
         OrigLeft        =   2520
         OrigTop         =   240
         OrigRight       =   2760
         OrigBottom      =   615
         Max             =   250
         Min             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Thermostat Set-Point (deg F)"
         Height          =   255
         Index           =   45
         Left            =   180
         TabIndex        =   236
         Top             =   300
         Width           =   2115
      End
      Begin VB.Label TLO_SET_F 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         TabIndex        =   235
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Emergency Fire Pump"
      Height          =   1155
      Left            =   8220
      TabIndex        =   230
      Top             =   10200
      Width           =   2955
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   42
         Left            =   1140
         TabIndex        =   231
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label Fire_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0FF"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1380
         TabIndex        =   232
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1155
      Left            =   3780
      TabIndex        =   226
      Top             =   10200
      Width           =   4335
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   41
         Left            =   2760
         TabIndex        =   227
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "HFO_ign_lbl"
         BuddyDispid     =   196617
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   250
         Min             =   86
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "HFO Ignition Temp (deg F)"
         Height          =   255
         Index           =   44
         Left            =   480
         TabIndex        =   229
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label HFO_ign_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3000
         TabIndex        =   228
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF00FF&
      Caption         =   "Explosion Sound"
      Height          =   855
      Left            =   15060
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   10080
      Width           =   2235
   End
   Begin VB.Frame Frame8 
      Caption         =   "VIAX"
      Height          =   3315
      Left            =   14700
      TabIndex        =   206
      Top             =   6480
      Width           =   3675
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   35
         Left            =   2460
         TabIndex        =   207
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   36
         Left            =   2460
         TabIndex        =   208
         Top             =   780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   37
         Left            =   2460
         TabIndex        =   209
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   38
         Left            =   2460
         TabIndex        =   210
         Top             =   1740
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   39
         Left            =   2460
         TabIndex        =   211
         Top             =   2220
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   40
         Left            =   2460
         TabIndex        =   212
         Top             =   2700
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Enabled         =   -1  'True
      End
      Begin VB.Label VIAX_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   2700
         TabIndex        =   224
         Top             =   2700
         Width           =   555
      End
      Begin VB.Label VIAX_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   2700
         TabIndex        =   223
         Top             =   2220
         Width           =   555
      End
      Begin VB.Label VIAX_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   2700
         TabIndex        =   222
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label VIAX_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2700
         TabIndex        =   221
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label VIAX_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2700
         TabIndex        =   220
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "LP Axial displacement FWD mm "
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   219
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "HP Axial displacement FWD mm "
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   218
         Top             =   2340
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "LP Axial displacement AFT mm "
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   217
         Top             =   1860
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "LP Vibration ( 0-1 G )"
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   216
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label VIAX_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   2700
         TabIndex        =   215
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "HP Axial displacement AFT mm "
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   214
         Top             =   1380
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "HP Vibration ( 0-1 G )"
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   213
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SSTG Vacuum Trips"
      Height          =   1215
      Index           =   5
      Left            =   0
      TabIndex        =   197
      Top             =   8040
      Width           =   3735
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   33
         Left            =   2280
         TabIndex        =   199
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   240
         Width           =   390
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   34
         Left            =   2280
         TabIndex        =   198
         Text            =   "5"
         Top             =   720
         Width           =   390
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   33
         Left            =   2640
         TabIndex        =   200
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "Parameter(33)"
         BuddyDispid     =   196623
         BuddyIndex      =   33
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   28
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   34
         Left            =   2640
         TabIndex        =   201
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "Parameter(34)"
         BuddyDispid     =   196623
         BuddyIndex      =   34
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   28
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SSTG# 2"
         Height          =   255
         Index           =   42
         Left            =   420
         TabIndex        =   205
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "in Hg"
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   204
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SSTG# 1"
         Height          =   255
         Index           =   43
         Left            =   420
         TabIndex        =   203
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "in Hg"
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   202
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Salinty"
      Height          =   3315
      Left            =   11280
      TabIndex        =   178
      Top             =   6360
      Width           =   3315
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   27
         Left            =   1920
         TabIndex        =   179
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   28
         Left            =   1920
         TabIndex        =   187
         Top             =   780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   29
         Left            =   1920
         TabIndex        =   189
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   30
         Left            =   1920
         TabIndex        =   191
         Top             =   1740
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   31
         Left            =   2400
         TabIndex        =   193
         Top             =   2220
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   32
         Left            =   1920
         TabIndex        =   195
         Top             =   2640
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin VB.Label Salinity_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   196
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Salinity_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   194
         Top             =   2220
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Salinity_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   192
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Salinity_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   190
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Salinity_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   188
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Condenser Hotwell"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   186
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Atmospheric Drain Tank"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   185
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Starboard Boiler Drum"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   184
         Top             =   1380
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Port Boiler Drum"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   183
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Distilling Plant"
         Height          =   255
         Index           =   4
         Left            =   420
         TabIndex        =   182
         Top             =   2280
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Distilled Water Tank"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   181
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label Salinity_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   180
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Shaft Lock"
      Height          =   1215
      Left            =   11280
      TabIndex        =   173
      Top             =   9840
      Width           =   3315
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   26
         Left            =   1380
         TabIndex        =   174
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Note: STOP the shaft then LOCK it"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   41
         Left            =   120
         TabIndex        =   177
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label SHAFT_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0FF"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1620
         TabIndex        =   176
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Shaft Lock"
         Height          =   255
         Index           =   40
         Left            =   180
         TabIndex        =   175
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Deck Electrical Loads (at 440 Volts)"
      Height          =   2835
      Left            =   3780
      TabIndex        =   150
      Top             =   6360
      Width           =   4335
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   21
         Left            =   2400
         TabIndex        =   151
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Deck_lbl(1)"
         BuddyDispid     =   196630
         BuddyIndex      =   1
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
         Index           =   22
         Left            =   2400
         TabIndex        =   154
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Deck_lbl(2)"
         BuddyDispid     =   196630
         BuddyIndex      =   2
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
         Index           =   23
         Left            =   2400
         TabIndex        =   156
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Deck_lbl(3)"
         BuddyDispid     =   196630
         BuddyIndex      =   3
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
         Index           =   24
         Left            =   2400
         TabIndex        =   158
         Top             =   1680
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Deck_lbl(4)"
         BuddyDispid     =   196630
         BuddyIndex      =   4
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "kW/phase"
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   172
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "kW/phase"
         Height          =   255
         Index           =   10
         Left            =   3240
         TabIndex        =   171
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "kW/phase"
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   170
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "kW/phase"
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   169
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deck Load ( 0 to 100 kW )"
         Height          =   255
         Index           =   37
         Left            =   360
         TabIndex        =   168
         Top             =   1740
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deck Load ( 0 to 100 kW )"
         Height          =   255
         Index           =   36
         Left            =   360
         TabIndex        =   167
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deck Load ( 0 to 100 kW )"
         Height          =   255
         Index           =   35
         Left            =   360
         TabIndex        =   166
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label Deck_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   159
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Deck_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   157
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Deck_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   155
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Deck_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   153
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deck Load ( 0 to 100 kW )"
         Height          =   255
         Index           =   17
         Left            =   360
         TabIndex        =   152
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "FO Heater Thermostat"
      Height          =   915
      Left            =   3780
      TabIndex        =   146
      Top             =   9240
      Width           =   4335
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   20
         Left            =   2760
         TabIndex        =   147
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "FO_htr_lbl(0)"
         BuddyDispid     =   196632
         BuddyIndex      =   0
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   250
         Min             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label FO_htr_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   149
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Thermostat Set-Point (deg F)"
         Height          =   255
         Index           =   34
         Left            =   180
         TabIndex        =   148
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Deck Steam Load"
      Height          =   915
      Left            =   8220
      TabIndex        =   142
      Top             =   9240
      Width           =   3015
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   19
         Left            =   2040
         TabIndex        =   143
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Deck_lbl(0)"
         BuddyDispid     =   196630
         BuddyIndex      =   0
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deck Steam ( 0 to 100 % )"
         Height          =   255
         Index           =   33
         Left            =   60
         TabIndex        =   145
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Deck_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   144
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7920
      Top             =   10920
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   7980
      Top             =   10440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alternator Droop"
      Height          =   1755
      Index           =   4
      Left            =   0
      TabIndex        =   45
      Top             =   6240
      Width           =   3735
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   25
         Left            =   2280
         TabIndex        =   161
         Text            =   "5"
         Top             =   780
         Width           =   390
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   16
         Left            =   2280
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   17
         Left            =   2280
         TabIndex        =   46
         Text            =   "5"
         Top             =   1260
         Width           =   375
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   16
         Left            =   2640
         TabIndex        =   48
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(16)"
         BuddyDispid     =   196623
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
         TabIndex        =   49
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(17)"
         BuddyDispid     =   196623
         BuddyIndex      =   17
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
         Index           =   25
         Left            =   2640
         TabIndex        =   160
         Top             =   780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Parameter(25)"
         BuddyDispid     =   196623
         BuddyIndex      =   25
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
         Index           =   7
         Left            =   3000
         TabIndex        =   165
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   164
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SSTG# 2 Droop"
         Height          =   255
         Index           =   39
         Left            =   240
         TabIndex        =   163
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SSTG# 1 Droop"
         Height          =   255
         Index           =   38
         Left            =   240
         TabIndex        =   162
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SSDG # 1 Droop"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   50
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cargo Pump Loading"
      Height          =   2715
      Index           =   3
      Left            =   8220
      TabIndex        =   39
      Top             =   6360
      Width           =   3015
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   18
         Left            =   2820
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "25"
         Top             =   1320
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   15
         Left            =   2820
         TabIndex        =   42
         Text            =   "50"
         Top             =   780
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   14
         Left            =   2820
         TabIndex        =   41
         Text            =   "100"
         Top             =   300
         Visible         =   0   'False
         Width           =   675
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   14
         Left            =   2040
         TabIndex        =   43
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Cargo_lbl(1)"
         BuddyDispid     =   196636
         BuddyIndex      =   1
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
         Left            =   2040
         TabIndex        =   44
         Top             =   780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         BuddyControl    =   "Cargo_lbl(2)"
         BuddyDispid     =   196636
         BuddyIndex      =   2
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
         Index           =   18
         Left            =   2040
         TabIndex        =   96
         Top             =   1320
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "Cargo_lbl(3)"
         BuddyDispid     =   196636
         BuddyIndex      =   3
         OrigLeft        =   1560
         OrigTop         =   360
         OrigRight       =   1800
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Cargo_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   137
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Cargo_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   136
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Cargo_lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   135
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pump # 3 ( 0 to 100 % )"
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   94
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pump #2  ( 0 to 100 % )"
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   93
         Top             =   900
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pump # 1 ( 0 to 100 % )"
         Height          =   255
         Index           =   14
         Left            =   180
         TabIndex        =   40
         Top             =   420
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fuel Analysis (by mass)"
      Height          =   6255
      Index           =   2
      Left            =   3780
      TabIndex        =   32
      Top             =   0
      Width           =   14715
      Begin VB.TextBox Current_fuel_by_analysis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Height          =   345
         Left            =   2700
         TabIndex        =   125
         Text            =   "Diesel Oil"
         Top             =   5700
         Width           =   1575
      End
      Begin VB.TextBox Current_Fuel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2700
         TabIndex        =   102
         Text            =   "Diesel Oil"
         Top             =   5220
         Width           =   1575
      End
      Begin VB.TextBox LCV_current 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2340
         TabIndex        =   80
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         Caption         =   "New Value"
         Height          =   5655
         Left            =   4440
         TabIndex        =   65
         Top             =   240
         Width           =   10215
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enter"
            Height          =   375
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   5040
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enter"
            Height          =   375
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   5040
            Width           =   1095
         End
         Begin VB.TextBox LCV_new 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   1560
            TabIndex        =   83
            Top             =   4260
            Width           =   1815
         End
         Begin VB.TextBox HCV_new 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   1560
            TabIndex        =   82
            Top             =   3840
            Width           =   1815
         End
         Begin VB.TextBox NEW_total 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   1740
            TabIndex        =   79
            Text            =   "TOTAL > 101 %"
            Top             =   3300
            Width           =   1395
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enter"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   5040
            Width           =   1095
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   8
            Left            =   1980
            TabIndex        =   66
            Top             =   420
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            Value           =   750
            OrigLeft        =   1560
            OrigTop         =   360
            OrigRight       =   1800
            OrigBottom      =   735
            Max             =   900
            Min             =   750
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   9
            Left            =   1980
            TabIndex        =   67
            Top             =   900
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            Value           =   100
            OrigLeft        =   2640
            OrigTop         =   840
            OrigRight       =   2880
            OrigBottom      =   1215
            Max             =   150
            Min             =   100
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   10
            Left            =   1980
            TabIndex        =   68
            Top             =   1380
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            OrigLeft        =   1560
            OrigTop         =   360
            OrigRight       =   1800
            OrigBottom      =   735
            Max             =   50
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   11
            Left            =   1980
            TabIndex        =   69
            Top             =   1860
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            OrigLeft        =   1560
            OrigTop         =   360
            OrigRight       =   1800
            OrigBottom      =   735
            Max             =   50
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   12
            Left            =   1980
            TabIndex        =   70
            Top             =   2340
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            OrigLeft        =   2640
            OrigTop         =   840
            OrigRight       =   2880
            OrigBottom      =   1215
            Max             =   70
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Index           =   13
            Left            =   1980
            TabIndex        =   71
            Top             =   2820
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            OrigLeft        =   1560
            OrigTop         =   360
            OrigRight       =   1800
            OrigBottom      =   735
            Max             =   50
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enter"
            Height          =   375
            Left            =   6180
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   5040
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "Heavy FO"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   32
            Left            =   8400
            TabIndex        =   141
            Top             =   4800
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Diesel Oil"
            Height          =   195
            Index           =   31
            Left            =   4260
            TabIndex        =   140
            Top             =   4800
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0000C0C0&
            Caption         =   "Instructor Designed Oil"
            Height          =   195
            Index           =   30
            Left            =   1620
            TabIndex        =   139
            Top             =   4800
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000040C0&
            Caption         =   "Light FO"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   29
            Left            =   6180
            TabIndex        =   138
            Top             =   4800
            Width           =   1095
         End
         Begin VB.Label LCV_HFO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7920
            TabIndex        =   134
            Top             =   4260
            Width           =   1815
         End
         Begin VB.Label HCV_HFO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7920
            TabIndex        =   133
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label LCV_LFO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   5880
            TabIndex        =   132
            Top             =   4260
            Width           =   1815
         End
         Begin VB.Label HCV_LFO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   5880
            TabIndex        =   131
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label LCV_DO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   3780
            TabIndex        =   130
            Top             =   4260
            Width           =   1815
         End
         Begin VB.Label HCV_DO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   3780
            TabIndex        =   129
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100 %"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   8100
            TabIndex        =   128
            Top             =   3300
            Width           =   1395
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100 %"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   6120
            TabIndex        =   127
            Top             =   3300
            Width           =   1395
         End
         Begin VB.Label DO_total 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "100 %"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4020
            TabIndex        =   126
            Top             =   3300
            Width           =   1395
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "86.1"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   30
            Left            =   8520
            TabIndex        =   123
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "11.8"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   31
            Left            =   8520
            TabIndex        =   122
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "Heavy FO"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   28
            Left            =   8280
            TabIndex        =   121
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   32
            Left            =   8520
            TabIndex        =   120
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   33
            Left            =   8520
            TabIndex        =   119
            Top             =   1860
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2.1"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   34
            Left            =   8520
            TabIndex        =   118
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   35
            Left            =   8520
            TabIndex        =   117
            Top             =   2820
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "86.2"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   20
            Left            =   6480
            TabIndex        =   115
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "12.4"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   21
            Left            =   6480
            TabIndex        =   114
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000040C0&
            Caption         =   "Light FO"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   27
            Left            =   6240
            TabIndex        =   113
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   22
            Left            =   6480
            TabIndex        =   112
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   23
            Left            =   6480
            TabIndex        =   111
            Top             =   1860
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1.4"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   24
            Left            =   6480
            TabIndex        =   110
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   25
            Left            =   6480
            TabIndex        =   109
            Top             =   2820
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   4440
            TabIndex        =   106
            Top             =   2820
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.9"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   4440
            TabIndex        =   105
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   4440
            TabIndex        =   104
            Top             =   1860
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   4440
            TabIndex        =   103
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0000C0C0&
            Caption         =   "Instructor Designed Oil"
            Height          =   195
            Index           =   26
            Left            =   1500
            TabIndex        =   100
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "Diesel Oil"
            Height          =   195
            Index           =   16
            Left            =   4200
            TabIndex        =   99
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "12.8"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   4440
            TabIndex        =   98
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "86.3"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   97
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Hydrogen ( 10 to 15 % )"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   92
            Top             =   960
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Oxygen ( 0 to 5 % )"
            Height          =   255
            Index           =   23
            Left            =   420
            TabIndex        =   91
            Top             =   1500
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nitrogen ( 0 to 5 % )"
            Height          =   255
            Index           =   22
            Left            =   360
            TabIndex        =   90
            Top             =   1920
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sulphur ( 0 to 7 % )"
            Height          =   255
            Index           =   21
            Left            =   360
            TabIndex        =   89
            Top             =   2400
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Ash ( 0 to 5 % )"
            Height          =   255
            Index           =   20
            Left            =   420
            TabIndex        =   88
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "New Total"
            Height          =   255
            Left            =   660
            TabIndex        =   87
            Top             =   3420
            Width           =   1035
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "New  LCV "
            Height          =   195
            Left            =   480
            TabIndex        =   86
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "New  HCV "
            Height          =   195
            Left            =   480
            TabIndex        =   85
            Top             =   3960
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Carbon ( 75 to 90% )"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   84
            Top             =   540
            Width           =   1635
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.5"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   13
            Left            =   2220
            TabIndex        =   77
            Top             =   2820
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.2"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   12
            Left            =   2220
            TabIndex        =   76
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   11
            Left            =   2220
            TabIndex        =   75
            Top             =   1860
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   10
            Left            =   2220
            TabIndex        =   74
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "11.2"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   9
            Left            =   2220
            TabIndex        =   73
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Fuel_lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "85.0"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   8
            Left            =   2220
            TabIndex        =   72
            Top             =   420
            Width           =   495
         End
      End
      Begin VB.TextBox Current_total 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2340
         TabIndex        =   63
         Top             =   3420
         Width           =   1155
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   13
         Left            =   2760
         TabIndex        =   60
         Text            =   "0"
         Top             =   2940
         Width           =   615
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   12
         Left            =   2760
         TabIndex        =   59
         Text            =   "0"
         Top             =   2460
         Width           =   615
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   56
         Text            =   "0"
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox HCV_current 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2340
         TabIndex        =   52
         Top             =   4020
         Width           =   1155
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   10
         Left            =   2760
         TabIndex        =   35
         Text            =   "0"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   9
         Left            =   2760
         TabIndex        =   34
         Text            =   "0"
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox Parameter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   8
         Left            =   2760
         TabIndex        =   33
         Text            =   "0"
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Current Fuel [BY FUEL ANALYSIS]"
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   5760
         Width           =   2475
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Current Fuel  [AT THE BURNERS]"
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   5280
         Width           =   2475
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Lower Calorific Value LCV "
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   4500
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CurrentTotal"
         Height          =   255
         Left            =   1260
         TabIndex        =   64
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Current Value"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   62
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ash ( 0 to 5 % )"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   61
         Top             =   2940
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sulphur ( 0 to 7 % )"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nitrogen ( 0 to 5 % )"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   57
         Top             =   2100
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Higher Calorific Value HCV "
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   4080
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Carbon ( 75 to 90% )"
         Height          =   255
         Index           =   10
         Left            =   720
         TabIndex        =   38
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Hydrogen ( 10 to 15 % )"
         Height          =   255
         Index           =   9
         Left            =   720
         TabIndex        =   37
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Oxygen ( 0 to 5 % )"
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   36
         Top             =   1620
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ship and Environment "
      Height          =   3615
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   2580
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
         BuddyDispid     =   196623
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
         BuddyDispid     =   196623
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
         BuddyDispid     =   196623
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
         BuddyDispid     =   196623
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
         Value           =   7
         BuddyControl    =   "Parameter(7)"
         BuddyDispid     =   196623
         BuddyIndex      =   7
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   735
         Max             =   17
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "metres"
         Height          =   255
         Index           =   17
         Left            =   2940
         TabIndex        =   54
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Draft"
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
      Begin VB.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Head"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ballast"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ambient Temperatures"
      Height          =   2595
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
         Text            =   "1"
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
         Text            =   "1"
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
         BuddyDispid     =   196623
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
         Top             =   900
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Parameter(1)"
         BuddyDispid     =   196623
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
         BuddyDispid     =   196623
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
      Left            =   300
      TabIndex        =   55
      Top             =   300
      Width           =   435
   End
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   435
      Left            =   960
      Top             =   9540
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   767
   End
End
Attribute VB_Name = "Instrp_LSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim instrp_values(30) As Single
Private set_dummy_focus As Integer
Private NEWtotal As Single
Private Currenttotal As Single
Private C12 As Single
Private H1 As Single
Private O2 As Single
Private S32 As Single
Private HCV As Single
Private LCV As Single


Private Sub Command1_Click()

' 26/7/00
VB_fuel(1) = Val(Fuel_lbl(8))
VB_fuel(2) = Val(Fuel_lbl(9))
VB_fuel(3) = Val(Fuel_lbl(10))
VB_fuel(4) = Val(Fuel_lbl(11))
VB_fuel(5) = Val(Fuel_lbl(12))
VB_fuel(6) = Val(Fuel_lbl(13))

' spares
VB_fuel(7) = 0#
VB_fuel(8) = 0#
VB_fuel(9) = 0#
VB_fuel(10) = 0#


If NEWtotal < 98.9999 Or NEWtotal > 101.0001 Then
  HCV_new.text = "NOT entered"
  LCV_new.text = "New Total is incorrect"
Else

  Call set_FUEL(VB_fuel(1))
  
  ' Tell the model
  Fuel_Type = 3
  Call set_EXECI_LSS(40, Fuel_Type)


End If



' reset timer2
Timer2.Enabled = False
Timer2.Interval = 1500
Timer2.Enabled = True

End Sub

Private Sub Command2_Click()


' 2/8/00
VB_fuel(1) = Val(Fuel_lbl(0))
VB_fuel(2) = Val(Fuel_lbl(1))
VB_fuel(3) = Val(Fuel_lbl(2))
VB_fuel(4) = Val(Fuel_lbl(3))
VB_fuel(5) = Val(Fuel_lbl(4))
VB_fuel(6) = Val(Fuel_lbl(5))

' spares
VB_fuel(7) = 0#
VB_fuel(8) = 0#
VB_fuel(9) = 0#
VB_fuel(10) = 0#


'If NEWtotal < 98.9999 Or NEWtotal > 101.0001 Then
'  HCV_new.text = "NOT entered"
'  LCV_new.text = "New Total is incorrect"
'Else

  Call set_FUEL(VB_fuel(1))
  
'End If

  ' Tell the model
  Fuel_Type = 0
  Call set_EXECI_LSS(40, Fuel_Type)


' reset timer2
Timer2.Enabled = False
Timer2.Interval = 1500
Timer2.Enabled = True


End Sub

Private Sub Command3_Click()

' 2/8/00
VB_fuel(1) = Val(Fuel_lbl(20))
VB_fuel(2) = Val(Fuel_lbl(21))
VB_fuel(3) = Val(Fuel_lbl(22))
VB_fuel(4) = Val(Fuel_lbl(23))
VB_fuel(5) = Val(Fuel_lbl(24))
VB_fuel(6) = Val(Fuel_lbl(25))

' spares
VB_fuel(7) = 0#
VB_fuel(8) = 0#
VB_fuel(9) = 0#
VB_fuel(10) = 0#


'If NEWtotal < 98.9999 Or NEWtotal > 101.0001 Then
'  HCV_new.text = "NOT entered"
'  LCV_new.text = "New Total is incorrect"
'Else

  Call set_FUEL(VB_fuel(1))
  
'End If

  ' Tell the model
  Fuel_Type = 1
  Call set_EXECI_LSS(40, Fuel_Type)


' reset timer2
Timer2.Enabled = False
Timer2.Interval = 1500
Timer2.Enabled = True


End Sub

Private Sub Command4_Click()

' 2/8/00
VB_fuel(1) = Val(Fuel_lbl(30))
VB_fuel(2) = Val(Fuel_lbl(31))
VB_fuel(3) = Val(Fuel_lbl(32))
VB_fuel(4) = Val(Fuel_lbl(33))
VB_fuel(5) = Val(Fuel_lbl(34))
VB_fuel(6) = Val(Fuel_lbl(35))

' spares
VB_fuel(7) = 0#
VB_fuel(8) = 0#
VB_fuel(9) = 0#
VB_fuel(10) = 0#


'If NEWtotal < 98.9999 Or NEWtotal > 101.0001 Then
'  HCV_new.text = "NOT entered"
'  LCV_new.text = "New Total is incorrect"
'Else

  Call set_FUEL(VB_fuel(1))
  
'End If

  ' Tell the model
  Fuel_Type = 2
  Call set_EXECI_LSS(40, Fuel_Type)


' reset timer2
Timer2.Enabled = False
Timer2.Interval = 1500
Timer2.Enabled = True

End Sub

Private Sub Command5_Click()

' 14/12/00
Explosion_count = 0


End Sub

Private Sub Form_Load()
Dim i As Integer

' 16/7/98
Timer1.Enabled = False
Call update_values

set_dummy_focus = True

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

' 25/7/00 Initialise some up-downs
For i = 8 To 13
  UpDown1(i).Value = Val(Fuel_lbl(i)) * 10
Next i


' 2/8/00  Calculation for Diesel Oil
    C12 = Val(Fuel_lbl(0).Caption) * 0.01
    H1 = Val(Fuel_lbl(1).Caption) * 0.01
    O2 = Val(Fuel_lbl(2).Caption) * 0.01
    S32 = Val(Fuel_lbl(4).Caption) * 0.01

  '  Higher Calorific Value ( MJ/kg )
    HCV = 33.7 * C12 + 144# * (H1 - O2 / 8#) + 9.3 * S32
  '   Lower Calorific Value ( MJ/kg )
    LCV = HCV - (2.442 * 9# * H1)


  HCV_DO.Caption = Format(HCV, "##0.00") & " MJ/kg"
  LCV_DO.Caption = Format(LCV, "##0.00") & " MJ/kg"

' 2/8/00  Calculation for Light Fuel Oil
    C12 = Val(Fuel_lbl(20).Caption) * 0.01
    H1 = Val(Fuel_lbl(21).Caption) * 0.01
    O2 = Val(Fuel_lbl(22).Caption) * 0.01
    S32 = Val(Fuel_lbl(24).Caption) * 0.01

  '  Higher Calorific Value ( MJ/kg )
    HCV = 33.7 * C12 + 144# * (H1 - O2 / 8#) + 9.3 * S32
  '   Lower Calorific Value ( MJ/kg )
    LCV = HCV - (2.442 * 9# * H1)


  HCV_LFO.Caption = Format(HCV, "##0.00") & " MJ/kg"
  LCV_LFO.Caption = Format(LCV, "##0.00") & " MJ/kg"

' 2/8/00  Calculation for Heavy Fuel Oil
    C12 = Val(Fuel_lbl(30).Caption) * 0.01
    H1 = Val(Fuel_lbl(31).Caption) * 0.01
    O2 = Val(Fuel_lbl(32).Caption) * 0.01
    S32 = Val(Fuel_lbl(34).Caption) * 0.01

  '  Higher Calorific Value ( MJ/kg )
    HCV = 33.7 * C12 + 144# * (H1 - O2 / 8#) + 9.3 * S32
  '   Lower Calorific Value ( MJ/kg )
    LCV = HCV - (2.442 * 9# * H1)


  HCV_HFO.Caption = Format(HCV, "##0.00") & " MJ/kg"
  LCV_HFO.Caption = Format(LCV, "##0.00") & " MJ/kg"


End Sub

Private Sub Timer1_Timer()

'Parameter(0).Text = Format(instrp_values(0), "#0") & Chr(176) & "C"

   Call update_values
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
Dim i As Integer


' 9/8/00
If Force_Instrp_update Then
   Force_Instrp_update = False
    Timer1.Enabled = True
End If


' 17/7/98
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


Parameter(1).text = Format(LSS_reals(58), "##0.0")   '  Tamb_steam


'equivalence(STEAM_logicals(14),HEAVY_OIL)
If VB_steam_logicals(14) Then

  Current_Fuel.text = "Heavy Oil"
  Current_Fuel.BackColor = &H40C0&      ' Dark Brown
  Current_Fuel.ForeColor = &HFFFFFF     ' White
  
Else

  Current_Fuel.text = "Diesel Oil"
  Current_Fuel.BackColor = &H80FF&      ' Orange
  Current_Fuel.ForeColor = &H0&         ' Black

End If


If LSS_reals(564) > 2.5 Then

  Current_fuel_by_analysis.text = "Instructor Designed"
  
  Current_fuel_by_analysis.BackColor = &HC0C0&        ' Greeny Brown
  Current_fuel_by_analysis.ForeColor = &H0&           ' Black

ElseIf LSS_reals(564) > 1.5 Then

  Current_fuel_by_analysis.text = "Heavy FO"
  Current_fuel_by_analysis.BackColor = &H4080&      ' Dark Brown
  Current_fuel_by_analysis.ForeColor = &HFFFFFF     ' White
  
ElseIf LSS_reals(564) > 0.5 Then

  Current_fuel_by_analysis.text = "Light FO"
  Current_fuel_by_analysis.BackColor = &H40C0&      ' Light Brown
  Current_fuel_by_analysis.ForeColor = &HFFFFFF     ' White
  
Else

  Current_fuel_by_analysis.text = "Diesel Oil"
  Current_fuel_by_analysis.BackColor = &H80FF&         ' Orange
  Current_fuel_by_analysis.ForeColor = &H0&            ' Black
  
End If



  Currenttotal = 0#

For i = 0 To 5
  Parameter(8 + i).text = Format(LSS_reals(129 + i), "##0.0")
  Currenttotal = Currenttotal + LSS_reals(129 + i)
Next i
  Current_total.text = Format(Currenttotal, "###0.0") & " %"

'      PC_REALS(135) = HCV
'      PC_REALS(136) = LCV

HCV_current.text = Format(LSS_reals(135), "##0.00") & " MJ/kg"
LCV_current.text = Format(LSS_reals(136), "##0.00") & " MJ/kg"


' 24/7/00
  NEWtotal = 0#
For i = 8 To 13
  NEWtotal = NEWtotal + UpDown1(i).Value * 0.1
Next i

If NEWtotal > 101.0001 Then

  NEW_total.text = "TOTAL > 101.0 %"
  NEW_total.ForeColor = &HFF&         'RED
  HCV_new.text = "n/a"
  LCV_new.text = "n/a"
  
ElseIf NEWtotal < 98.9999 Then

  NEW_total.text = "TOTAL < 99.0 %"
  NEW_total.ForeColor = &HFF&         'RED
  HCV_new.text = "n/a"
  LCV_new.text = "n/a"
  
Else

  NEW_total.text = Format(NEWtotal, "###0.0") & " %"
  NEW_total.ForeColor = &H80000008         'BLACK
  


    ' === 24/7/00 ===
    C12 = Val(Fuel_lbl(8).Caption) * 0.01
    H1 = Val(Fuel_lbl(9).Caption) * 0.01
    O2 = Val(Fuel_lbl(10).Caption) * 0.01
    S32 = Val(Fuel_lbl(12).Caption) * 0.01

    'c  Higher Calorific Value ( MJ/kg )
    HCV = 33.7 * C12 + 144# * (H1 - O2 / 8#) + 9.3 * S32
  'c   Lower Calorific Value ( MJ/kg )
    LCV = HCV - (2.442 * 9# * H1)


  HCV_new.text = Format(HCV, "##0.00") & " MJ/kg"
  LCV_new.text = Format(LCV, "##0.00") & " MJ/kg"

End If


Timer2.Enabled = False
Timer2.Interval = 500
Timer2.Enabled = True


End Sub

Private Sub UpDown1_Change(Index As Integer)
Dim real_val As Single

' 25/7/00
real_val = UpDown1(Index).Value * 0.1


' 16/7/98
Select Case Index

 Case 0    ' EXECI(14)    ' Tsea
     Call set_EXECI_LSS(14, Val(Parameter(Index)))
 Case 1    ' EXECI(15)    ' Tamb_steam
     Call set_EXECI_LSS(15, Val(Parameter(Index)))
 Case 2    ' EXECI(17)    ' Toutside
     Call set_EXECI_LSS(17, Val(Parameter(Index)))
     
 Case 3    ' EXECI(11) HULL_LOAD
     Call set_EXECI_LSS(11, Val(Parameter(Index)))
 Case 4    ' EXECI(12) WIND_LOAD
     Call set_EXECI_LSS(12, Val(Parameter(Index)))
 Case 5    ' EXECI(13) BALLAST_LOAD
     Call set_EXECI_LSS(13, Val(Parameter(Index)))
 Case 6    ' EXECI(16) SEA_STATE
     Call set_EXECI_LSS(16, Val(Parameter(Index)))
 Case 7    ' EXECI(48)  DRAFT
     Call set_EXECI_LSS(48, Val(Parameter(Index)))
     
 Case 8
 
   ' ==== 24/7/00    FUEL =========
   '  Call set_EXECI_LSS(50, Val(Parameter(Index)))
     Fuel_lbl(Index).Caption = Format(real_val, "##0.0")
     
 Case 9
   '  Call set_EXECI_LSS(51, Val(Parameter(Index)))
     Fuel_lbl(Index).Caption = Format(real_val, "##0.0")
 
 Case 10
   '  Call set_EXECI_LSS(10, Val(Parameter(Index)))
     Fuel_lbl(Index).Caption = Format(real_val, "##0.0")
   
 Case 11    ' EXECI(52) = GOV_LOW_LIM
   '  Call set_EXECI_LSS(52, Val(Parameter(Index)))
     Fuel_lbl(Index).Caption = Format(real_val, "##0.0")
      
Case 12    ' EXECI(53) = GOV_HIGH_LIM
    ' Call set_EXECI_LSS(53, Val(Parameter(Index)))
     Fuel_lbl(Index).Caption = Format(real_val, "##0.0")
      
Case 13    ' EXECI(54) = GOV_PSCAV_LOLIM
    ' Call set_EXECI_LSS(54, Val(Parameter(Index)))
     Fuel_lbl(Index).Caption = Format(real_val, "##0.0")
   
   ' ==== 24/7/00    FUEL end =====
    
 Case 14    ' EXECI(55) = CARGO # 1 LOADING  ( 1/8/00 )
     Call set_EXECI_LSS(55, Val(Cargo_lbl(1)))
     
 Case 15    ' EXECI(56) = CARGO # 2 LOADING  ( 1/8/00 )
     Call set_EXECI_LSS(56, Val(Cargo_lbl(2)))
     
     
 Case 16   ' EXECI(57) = DROOP_PERCENT(1)   ( D/A 1 DROOP )
     Call set_EXECI_LSS(57, Val(Parameter(Index)))
     
 ' 4/9/00
 Case 17   ' EXECI(59) = DROOP_PERCENT(2)   ( SSTG 2 DROOP )
     Call set_EXECI_LSS(59, Val(Parameter(Index)))
     
 Case 18   ' EXECI(  ) = CARGO # 3 LOADING ( 1/8/00 )
      
      Call set_EXECI_LSS(58, Val(Cargo_lbl(3)))
    
 Case 19   ' EXECI(  ) = DECK LOADING ( 8/8/00 )
      
      Call set_EXECI_LSS(54, Val(Deck_lbl(0)))
     
     
 Case 20   ' EXECI(  ) = TFO_heater_set_point_F  ( 9/8/00 )
      
      Call set_EXECI_LSS(53, Val(FO_htr_lbl(0)))
 
 ' 6/9/00
 Case 21   ' EXECI(61) = DECK ELEC LOADING
      Call set_EXECI_LSS(61, Val(Deck_lbl(1)))
 Case 22   ' EXECI(62) = DECK ELEC LOADING
      Call set_EXECI_LSS(62, Val(Deck_lbl(2)))
 Case 23   ' EXECI(63) = DECK ELEC LOADING
      Call set_EXECI_LSS(63, Val(Deck_lbl(3)))
 Case 24   ' EXECI(64) = DECK ELEC LOADING
      Call set_EXECI_LSS(64, Val(Deck_lbl(4)))
     
     
 ' 4/9/00
 Case 25   ' EXECI(60) = DROOP_PERCENT(3)   ( SSTG 1 DROOP )
     Call set_EXECI_LSS(60, Val(Parameter(Index)))
 
 
 Case 26   ' EXECI(65) = Shaft Lock ON/OFF
    If SHAFT_lbl.Caption = "ON" Then
      Call set_EXECI_LSS(65, 0)
    Else
      Call set_EXECI_LSS(65, 1)
    End If
     
 Case 27   ' EXECI(66) = SALINITY_value(1)
     Salinity_lbl(0).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(66, Val(Salinity_lbl(0)) * 10)
         
 Case 28   ' EXECI(67) = SALINITY_value(2)
     Salinity_lbl(1).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(67, Val(Salinity_lbl(1)) * 10)
     
 Case 29   ' EXECI(68) = SALINITY_value(3)
     Salinity_lbl(2).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(68, Val(Salinity_lbl(2)) * 10)
     
 Case 30   ' EXECI(69) = SALINITY_value(4)
     Salinity_lbl(3).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(69, Val(Salinity_lbl(3)) * 10)
     
 Case 31   ' EXECI(70) = SALINITY_value(5)
     Salinity_lbl(4).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(70, Val(Salinity_lbl(4)) * 10)
     
 Case 32   ' EXECI(71) = SALINITY_value(6)
     Salinity_lbl(5).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(71, Val(Salinity_lbl(5)) * 10)
     
' 15/9/00
 Case 33   ' EXECI(72) = SSTG # 1 VAC TRIP
     Call set_EXECI_LSS(72, Val(Parameter(Index)))
' 15/9/00
 Case 34   ' EXECI(73) = SSTG # 2 VAC TRIP
     Call set_EXECI_LSS(73, Val(Parameter(Index)))


 Case 35   ' EXECI(74) = VIAX
     VIAX_lbl(0).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(74, Val(VIAX_lbl(0)) * 10)
     
 Case 36   ' EXECI(75) = VIAX
     VIAX_lbl(1).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(75, Val(VIAX_lbl(1)) * 10)
 Case 37   ' EXECI(76) = VIAX
     VIAX_lbl(2).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(76, Val(VIAX_lbl(2)) * 10)
 Case 38   ' EXECI(77) = VIAX
     VIAX_lbl(3).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(77, Val(VIAX_lbl(3)) * 10)
 Case 39   ' EXECI(78) = VIAX
     VIAX_lbl(4).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(78, Val(VIAX_lbl(4)) * 10)
 Case 40   ' EXECI(79) = VIAX
     VIAX_lbl(5).Caption = Format(real_val, "##0.0")
     Call set_EXECI_LSS(79, Val(VIAX_lbl(5)) * 10)

    ' === 19/11/01 ===
 Case 41   ' EXECI(80) = HFO_Ign_Temp (deg F )
      Call set_EXECI_LSS(80, Val(HFO_ign_lbl))
     
 Case 42   ' EXECI(81) = Emergency Fire Pump ON/OFF
    If Fire_lbl.Caption = "ON" Then
      Call set_EXECI_LSS(81, 0)
    Else
      Call set_EXECI_LSS(81, 1)
    End If
     
    ' === 23/01/02 ===
 Case 43   ' EXECI(82) = TLO_SET_F (deg F )
      Call set_EXECI_LSS(82, Val(TLO_SET_F))
     
End Select

Timer1.Enabled = True

set_dummy_focus = True

End Sub

Sub update_values()
Dim i As Integer


' === 04/03/02 ===
Constant(1).Caption = Format(Shaft_Inertia_multiplier, "##0.0")
Constant(2).Caption = Format(Ship_Inertia_multiplier, "##0.0")
Constant(3).Caption = Format(Pitch_multiplier, "##0.0")



' === 19/11/01 ===
HFO_ign_lbl.Caption = Int(LSS_reals(824))

' === 23/02/01 ===
TLO_SET_F.Caption = Int(LSS_reals(836))




' 14/9/00
'              PC_REALS(579) = SALINITY_value(1)
'              PC_REALS(580) = SALINITY_value(2)
'              PC_REALS(581) = SALINITY_value(3)
'              PC_REALS(582) = SALINITY_value(4)
'              PC_REALS(583) = SALINITY_value(5)
'              PC_REALS(584) = SALINITY_value(6)
  
For i = 0 To 5

  Salinity_lbl(i).Caption = Format(LSS_reals(579 + i), "##0.0")
  
  
  ' 27/11/00
  VIAX_lbl(i).Caption = Format(LSS_reals(804 + i), "##0.0")

Next i


' 28/11/00
'         PC_REALS(804) = EXECI(74) * 0.1
'         PC_REALS(805) = EXECI(75) * 0.1
'         PC_REALS(806) = EXECI(76) * 0.1
'         PC_REALS(807) = EXECI(77) * 0.1
'         PC_REALS(808) = EXECI(78) * 0.1
'         PC_REALS(809) = EXECI(79) * 0.1







' 2/8/00
'          PC_REALS(549) = Cargo_pump_load(1) * 100#
'          PC_REALS(550) = Cargo_pump_load(2) * 100#
'          PC_REALS(551) = Cargo_pump_load(3) * 100#
Cargo_lbl(1).Caption = Int(LSS_reals(549))
Cargo_lbl(2).Caption = Int(LSS_reals(550))
Cargo_lbl(3).Caption = Int(LSS_reals(551))

'          PC_REALS(151) = Deck_stm_load(1)
Deck_lbl(0).Caption = Int(LSS_reals(151))


' 6/9/00
'      PC_REALS(746) = EXECI(61)
'      PC_REALS(747) = EXECI(62)
'      PC_REALS(748) = EXECI(63)
'      PC_REALS(749) = EXECI(64)
Deck_lbl(1).Caption = Int(LSS_reals(746))
Deck_lbl(2).Caption = Int(LSS_reals(747))
Deck_lbl(3).Caption = Int(LSS_reals(748))
Deck_lbl(4).Caption = Int(LSS_reals(749))


' 8/9/00
If VB_steam_logicals(132) Then
  SHAFT_lbl.Caption = "ON"
Else
  SHAFT_lbl.Caption = "OFF"
End If

' === 12/12/01 ===
'c       === 12/12/01  14.104 Emergency Fire Pump ===
'             LSS_LOGICALS(92) = DO(941)
If LSS_LOGICALS(92) Then
  Fire_lbl.Caption = "ON"
Else
  Fire_lbl.Caption = "OFF"
End If


' 9/8/00
FO_htr_lbl(0).Caption = Int(LSS_reals(153))


' 17/7/98
Parameter(0).text = Int(LSS_reals(59))     '  Tsea
''' Moved to timer2...
'''Parameter(1).Text = Format(LSS_reals(58), "##0.0")   '  Tamb_steam
Parameter(2).text = Int(LSS_reals(57))     '  Toutside

Parameter(3).text = Int(LSS_reals(557))    ' HULL

Parameter(4).text = Int(LSS_reals(558))    ' WIND
' === 26/11/02 ===
If (Int(LSS_reals(558)) > 0) Then
    Label2(4).Caption = "Head"
ElseIf (Int(LSS_reals(558)) < 0) Then
    Label2(4).Caption = "Tail"
Else
    Label2(4).Caption = ""
End If


Parameter(5).text = Int(LSS_reals(559))    ' BALLAST
Parameter(6).text = Int(LSS_reals(560))    ' SEA STATE
Parameter(7).text = Int(LSS_reals(725))    ' DRAFT ( 7 to 17 metres )


'    PC_REALS(129) = C12
'    PC_REALS(130) = H1
'    PC_REALS(131) = O2
'    PC_REALS(132) = n2
'    PC_REALS(133) = S32
'    PC_REALS(134) = ASH_in_fuel
'
'      PC_REALS(135) = HCV
'      PC_REALS(136) = LCV


' 26/7/00
' MOVED to timer2
'For i = 0 To 5
'  Parameter(8 + i).text = Format(LSS_reals(129 + i), "##0.00")
'Next i

' OLD LSS STUFF...
'Parameter(9).text = Int(LSS_reals(85))    ' SULPHUR
'Parameter(10).text = Int(LSS_reals(563))  ' HO Bunker Water
'Parameter(11).text = Int(LSS_reals(319))   '  GOV_LOW_LIM
'Parameter(12).text = Int(LSS_reals(318))   '  GOV_HIGH_LIM
'Parameter(13).text = Int(LSS_reals(327))   '  GOV_PSCAV_LOLIM


Parameter(14).text = Int(LSS_reals(326))   '  GOV_PSCAV_HILIM

Parameter(15).text = Int(LSS_reals(568))    '  GOV_DROOP

Parameter(16).text = Int(LSS_reals(565))    '  DROOP_percent(1)
Parameter(17).text = Int(LSS_reals(566))    '  DROOP_percent(2)

' 4/9/00
Parameter(25).text = Int(LSS_reals(567))    '  DROOP_percent(3)


' 15/9/00
'c    15/9/00EXECI(72)   reserved for SSTG # 1 Vac Trip
'     PC_REALS(727) = EXECI(72)
'     PC_REALS(728) = EXECI(73)
Parameter(33).text = Int(LSS_reals(727))    '  DROOP_percent(1)
Parameter(34).text = Int(LSS_reals(728))    '  DROOP_percent(1)


End Sub
