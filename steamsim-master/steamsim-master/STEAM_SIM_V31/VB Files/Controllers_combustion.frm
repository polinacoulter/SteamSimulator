VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form Controllers_combustion 
   Caption         =   "Combustion Control"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   9900
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7515
      Left            =   0
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   13256
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         Height          =   435
         Left            =   4080
         TabIndex        =   56
         Top             =   6900
         Width           =   675
      End
      Begin VB.CommandButton Command4 
         Height          =   255
         Left            =   7260
         TabIndex        =   52
         Top             =   6480
         Width           =   195
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   6600
         TabIndex        =   50
         Top             =   6480
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   6180
         TabIndex        =   48
         Top             =   5880
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   4860
         TabIndex        =   6
         Top             =   3300
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H000000FF&
         Caption         =   "      Under Development"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   720
         TabIndex        =   59
         Top             =   5160
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   58
         Top             =   4260
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   5
         Left            =   6180
         Picture         =   "Controllers_combustion.frx":0000
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Line Line80 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6420
         Y1              =   4260
         Y2              =   4080
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Man"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5700
         TabIndex        =   57
         Top             =   4500
         Width           =   495
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5940
         TabIndex        =   55
         Top             =   6240
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "0.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   54
         Top             =   6240
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port Master SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7500
         TabIndex        =   53
         Top             =   6480
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Master SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6000
         TabIndex        =   51
         Top             =   6480
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   7140
         X2              =   7140
         Y1              =   3660
         Y2              =   4080
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   49
         Top             =   3480
         Width           =   495
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6720
         Y1              =   4260
         Y2              =   4800
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   495
         Left            =   6540
         Top             =   5700
         Width           =   975
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   7440
         X2              =   7740
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   4
         Left            =   7620
         Picture         =   "Controllers_combustion.frx":030A
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   255
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   6660
         X2              =   6660
         Y1              =   1860
         Y2              =   1380
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   7740
         X2              =   7740
         Y1              =   1860
         Y2              =   1380
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   6960
         X2              =   6960
         Y1              =   3120
         Y2              =   2820
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   6660
         X2              =   6960
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Line Line25 
         BorderWidth     =   2
         X1              =   7380
         X2              =   7380
         Y1              =   6420
         Y2              =   6180
      End
      Begin VB.Line Line24 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6720
         Y1              =   6420
         Y2              =   6180
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   10
         Left            =   6900
         Picture         =   "Controllers_combustion.frx":0614
         Stretch         =   -1  'True
         Top             =   5460
         Width           =   255
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   47
         Top             =   5760
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   6660
         X2              =   6660
         Y1              =   2820
         Y2              =   2280
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   7740
         X2              =   7740
         Y1              =   2820
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   4440
         X2              =   4440
         Y1              =   3360
         Y2              =   4680
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   3420
         X2              =   4440
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3180
         TabIndex        =   46
         Top             =   3540
         Width           =   915
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3180
         TabIndex        =   45
         Top             =   4380
         Width           =   915
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   9
         Left            =   2220
         Picture         =   "Controllers_combustion.frx":091E
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   8
         Left            =   960
         Picture         =   "Controllers_combustion.frx":0C28
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   1
         Left            =   4200
         Picture         =   "Controllers_combustion.frx":0F32
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   0
         Left            =   4200
         Picture         =   "Controllers_combustion.frx":123C
         Stretch         =   -1  'True
         Top             =   4200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   1080
         X2              =   2520
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line22 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2520
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         Height          =   555
         Left            =   2520
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   44
         Top             =   1140
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         Height          =   435
         Left            =   780
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I/E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   43
         Top             =   2040
         Width           =   435
      End
      Begin VB.Line Line21 
         BorderWidth     =   2
         X1              =   1080
         X2              =   1080
         Y1              =   1980
         Y2              =   1500
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   42
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15 mA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   1620
         Width           =   915
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         Height          =   435
         Left            =   780
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   40
         Top             =   1140
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         Height          =   435
         Left            =   2040
         Top             =   1080
         Width           =   555
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   435
         Left            =   2040
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I/E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   39
         Top             =   2040
         Width           =   435
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2340
         Y1              =   1980
         Y2              =   1500
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   38
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15 mA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   37
         Top             =   1620
         Width           =   915
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2340
         Y1              =   4260
         Y2              =   2400
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   1080
         X2              =   1080
         Y1              =   4440
         Y2              =   2400
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   6
         Left            =   960
         Picture         =   "Controllers_combustion.frx":1546
         Stretch         =   -1  'True
         Top             =   1740
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   5
         Left            =   2220
         Picture         =   "Controllers_combustion.frx":1850
         Stretch         =   -1  'True
         Top             =   1740
         Width           =   255
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "875 psi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1740
         TabIndex        =   36
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "875 psi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   17
         Top             =   1020
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Pressure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   34
         Top             =   180
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   33
         Top             =   480
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   480
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2820
         OleObjectBlob   =   "Controllers_combustion.frx":1B5A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Greater_than_48.doc"
         TabIndex        =   31
         Top             =   3960
         Width           =   315
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "0.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   30
         Top             =   4560
         Width           =   735
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         Height          =   435
         Left            =   6900
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   29
         Top             =   4200
         Width           =   315
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7140
         TabIndex        =   28
         Top             =   4140
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   2
         Left            =   6180
         Picture         =   "Controllers_combustion.frx":6B72
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   3
         Left            =   6540
         Picture         =   "Controllers_combustion.frx":6E7C
         Stretch         =   -1  'True
         Top             =   1620
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   2
         Left            =   6840
         Picture         =   "Controllers_combustion.frx":7186
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   1
         Left            =   7320
         Picture         =   "Controllers_combustion.frx":7490
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   0
         Left            =   7020
         Picture         =   "Controllers_combustion.frx":779A
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   7
         Left            =   6600
         Picture         =   "Controllers_combustion.frx":7AA4
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "6.0 kg/s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   7440
         TabIndex        =   27
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15 mA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   26
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   25
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I/E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7500
         TabIndex        =   24
         Top             =   1920
         Width           =   435
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         Height          =   435
         Left            =   7440
         Top             =   1860
         Width           =   555
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   435
         Left            =   7440
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   23
         Top             =   1020
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7380
         TabIndex        =   22
         Top             =   300
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   435
         Left            =   6360
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "6.0 kg/s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15 mA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   20
         Top             =   1500
         Width           =   915
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   19
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I/E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6420
         TabIndex        =   18
         Top             =   1920
         Width           =   435
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   435
         Left            =   6360
         Top             =   1860
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6300
         TabIndex        =   16
         Top             =   300
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6900
         OleObjectBlob   =   "Controllers_combustion.frx":7DAE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Sigma_48.doc"
         TabIndex        =   15
         Top             =   4860
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Flow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         Top             =   0
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   7440
         X2              =   7440
         Y1              =   3120
         Y2              =   2820
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8.0 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6900
         TabIndex        =   13
         Top             =   5340
         Width           =   915
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7260
         TabIndex        =   12
         Top             =   3780
         Width           =   915
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5940
         X2              =   6420
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "7.5 V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   3780
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   10
         Top             =   3240
         Width           =   315
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   7140
         X2              =   7380
         Y1              =   3540
         Y2              =   3180
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   555
         Left            =   6720
         Top             =   3120
         Width           =   1035
      End
      Begin VB.OLE OLE11 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6900
         OleObjectBlob   =   "Controllers_combustion.frx":CDC6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Sigma_18.doc"
         TabIndex        =   9
         Top             =   3180
         Width           =   195
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   7020
         X2              =   7020
         Y1              =   5700
         Y2              =   5280
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "875 psi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2700
         TabIndex        =   8
         Top             =   3300
         Width           =   915
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   3300
         X2              =   4440
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   495
         Left            =   6540
         Top             =   4800
         Width           =   915
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   7140
         X2              =   7140
         Y1              =   4800
         Y2              =   4500
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Plant Master"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SP"
         Height          =   255
         Left            =   2940
         TabIndex        =   5
         Top             =   3660
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4140
         TabIndex        =   4
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4140
         TabIndex        =   3
         Top             =   3480
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   4440
         X2              =   6000
         Y1              =   4680
         Y2              =   4080
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   4440
         X2              =   6000
         Y1              =   3360
         Y2              =   4080
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4545
         OleObjectBlob   =   "Controllers_combustion.frx":11DDE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Percent_sign_18.doc"
         TabIndex        =   2
         Top             =   3840
         Width           =   255
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         OleObjectBlob   =   "Controllers_combustion.frx":16DF6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Delta_sign_18.doc"
         TabIndex        =   1
         Top             =   3840
         Width           =   255
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4920
         OleObjectBlob   =   "Controllers_combustion.frx":1BE0E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Integral_sign_18.doc"
         TabIndex        =   0
         Top             =   3840
         Width           =   255
      End
   End
End
Attribute VB_Name = "Controllers_combustion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Controller_PLANT_MASTER.ZOrder

End Sub

Private Sub Command3_Click()

' === 5/03/02 ===
Controller_Stbd.ZOrder

End Sub

Private Sub Command5_Click()

Key_form.ZOrder

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_combustion_controller)


' === 05/03/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_combustion_controller)
    HavenStretch1.Visible = True
End If

End Sub
