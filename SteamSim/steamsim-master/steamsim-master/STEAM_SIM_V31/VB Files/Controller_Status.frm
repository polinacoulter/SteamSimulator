VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form Controller_Status 
   Caption         =   "Controller Status / Controller Gauge Scales"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   11595
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4755
      Left            =   0
      Top             =   60
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8387
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   10380
         Top             =   3240
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   9
         X1              =   180
         X2              =   6480
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   8
         X1              =   180
         X2              =   6480
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   14
         X1              =   6480
         X2              =   6480
         Y1              =   3600
         Y2              =   4440
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "0 to 104 Bar"
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
         Index           =   10
         Left            =   4860
         TabIndex        =   59
         Top             =   4080
         Width           =   1635
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   11
         X1              =   10980
         X2              =   10980
         Y1              =   120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   4
         X1              =   180
         X2              =   10980
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   12
         X1              =   6480
         X2              =   6480
         Y1              =   120
         Y2              =   3120
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   9
         Left            =   9300
         TabIndex        =   58
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   8
         Left            =   9300
         TabIndex        =   57
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   7
         Left            =   9300
         TabIndex        =   56
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   6
         Left            =   9300
         TabIndex        =   55
         Top             =   1500
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   4
         Left            =   4860
         TabIndex        =   54
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   3
         Left            =   4860
         TabIndex        =   53
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   2
         Left            =   4860
         TabIndex        =   52
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   1
         Left            =   4860
         TabIndex        =   51
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "0 to 104 Bar"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   50
         Top             =   4080
         Width           =   1515
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   9
         Left            =   7500
         TabIndex        =   49
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   8
         Left            =   7500
         TabIndex        =   48
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   7
         Left            =   7500
         TabIndex        =   47
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   6
         Left            =   7500
         TabIndex        =   46
         Top             =   1500
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   4
         Left            =   3060
         TabIndex        =   45
         Top             =   2760
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   3
         Left            =   3060
         TabIndex        =   44
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   2
         Left            =   3060
         TabIndex        =   43
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   1
         Left            =   3060
         TabIndex        =   42
         Top             =   1500
         Width           =   1635
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "42"
         Height          =   255
         Index           =   9
         Left            =   7140
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "40"
         Height          =   255
         Index           =   8
         Left            =   7140
         TabIndex        =   40
         Top             =   2340
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "38"
         Height          =   255
         Index           =   7
         Left            =   7140
         TabIndex        =   39
         Top             =   1920
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "36"
         Height          =   255
         Index           =   6
         Left            =   7140
         TabIndex        =   38
         Top             =   1560
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "34"
         Height          =   255
         Index           =   5
         Left            =   7140
         TabIndex        =   37
         Top             =   1140
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "41"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "39"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   35
         Top             =   2340
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "37"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   34
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   9
         Left            =   6720
         TabIndex        =   33
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   8
         Left            =   6720
         TabIndex        =   32
         Top             =   2340
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   7
         Left            =   6720
         TabIndex        =   31
         Top             =   1980
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   6
         Left            =   6720
         TabIndex        =   30
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   4
         Left            =   2340
         TabIndex        =   29
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   3
         Left            =   2340
         TabIndex        =   28
         Top             =   2340
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   2
         Left            =   2340
         TabIndex        =   27
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   1
         Left            =   2340
         TabIndex        =   26
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "35"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   25
         Top             =   1500
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label I_lbl 
         Alignment       =   2  'Center
         Caption         =   "33"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   24
         Top             =   1140
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   16
         X1              =   2100
         X2              =   6480
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   7
         X1              =   2100
         X2              =   2100
         Y1              =   120
         Y2              =   3120
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Man/Auto"
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
         Index           =   8
         Left            =   2220
         TabIndex        =   23
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "LH Scale"
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
         Index           =   7
         Left            =   3720
         TabIndex        =   22
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "RH Scale"
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
         Index           =   6
         Left            =   5520
         TabIndex        =   21
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "A"
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
         Left            =   2340
         TabIndex        =   20
         Top             =   4080
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   15
         X1              =   2100
         X2              =   2100
         Y1              =   3600
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   13
         X1              =   180
         X2              =   180
         Y1              =   4020
         Y2              =   4440
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   0
         Left            =   4860
         TabIndex        =   19
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   " -230 to +230 mm "
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
         Index           =   0
         Left            =   3060
         TabIndex        =   18
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   0
         Left            =   2340
         TabIndex        =   17
         Top             =   1140
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "RH Scale"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "LH Scale"
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
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Man/Auto"
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
         Index           =   0
         Left            =   2220
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label RH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   5
         Left            =   9300
         TabIndex        =   13
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label LH_scale 
         Alignment       =   2  'Center
         Caption         =   "-9 to +9 inches"
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
         Index           =   5
         Left            =   7500
         TabIndex        =   12
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Man_Auto_lbl 
         Alignment       =   2  'Center
         Caption         =   "M"
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
         Index           =   5
         Left            =   6720
         TabIndex        =   11
         Top             =   1140
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   10
         X1              =   2100
         X2              =   10980
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   6
         X1              =   180
         X2              =   180
         Y1              =   1080
         Y2              =   3120
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "RH Scale"
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
         Index           =   5
         Left            =   9840
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "LH Scale"
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
         Index           =   4
         Left            =   8040
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Man/Auto"
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
         Index           =   3
         Left            =   6540
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Port"
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
         Index           =   0
         Left            =   7860
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   5
         X1              =   180
         X2              =   10980
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         X1              =   180
         X2              =   10980
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Plant Master"
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
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Master"
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
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   2760
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Air"
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
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   2340
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fuel"
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
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   180
         X2              =   10980
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   180
         X2              =   10980
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   180
         X2              =   10980
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Stbd"
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
         Index           =   2
         Left            =   3420
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Steam Temperature"
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
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Drum Level"
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
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   1140
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Controller_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub work_it_out(steam_integers_no As Integer, LH_text As String, RH_text As String, Man_mode As Boolean)
Dim Auto_unit_str As String


Select Case steam_integers_no
    Case 33, 34    ' Drum Level
        If SI_units Then
            Auto_unit_str = " -230 to +230 mm "
        Else
            Auto_unit_str = " -9 to +9 inches "
        End If
        
    Case 35, 36     ' Steam Temperature
    
        ' === 13/12/02 17.012 ===
        'If SI_units Then
        '    Auto_unit_str = " 0 to 1020" & Chr(176) & "C"
        'Else
        '    Auto_unit_str = " 32 to 1870" & Chr(176) & "F"
        'End If
        If SI_units Then
            Auto_unit_str = " 93 to 815" & Chr(176) & "C"
        Else
            Auto_unit_str = " 200 to 1500" & Chr(176) & "F"
        End If
        
    Case 37, 38    ' Fuel Flow
        If SI_units Then
            Auto_unit_str = " 0 to 2.02 kg/s "
        Else
            Auto_unit_str = " 0 to 4.45 lb/s "
        End If
        
    Case 39, 40     ' Air Flow
        If SI_units Then
            ' === 27/03/03 ===
            'Auto_unit_str = " 0 to 36 kg/s "
             Auto_unit_str = " 0 to 25 kg/s "
        Else
            ' === 27/03/03 ===
            'Auto_unit_str = " 0 to 80 lb/s "
             Auto_unit_str = " 0 to 55 lb/s "
        End If
        
    Case 41, 42   ' Master
        If SI_units Then
            Auto_unit_str = " 0 to 130 Bar "
        Else
            Auto_unit_str = " 0 to 1885 psi "
        End If

End Select

If VB_steam_integers(steam_integers_no) = 2 Or VB_steam_integers(steam_integers_no) = 3 Then
       ' AUTO or SP selected...
        LH_text = Auto_unit_str
        Man_mode = False
Else
        LH_text = " 0 to 100% "
        Man_mode = True
End If
        RH_text = Auto_unit_str

End Sub


Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_controller_status)

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_controller_status)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim i_steam_int_no As Integer
Dim LH_caption As String
Dim RH_caption As String
Dim Manual_mode As Boolean

For i = 0 To 9
    i_steam_int_no = Val(I_lbl(i))
    Call work_it_out(i_steam_int_no, LH_caption, RH_caption, Manual_mode)
    LH_scale(i).Caption = LH_caption
    RH_scale(i).Caption = RH_caption
    If Manual_mode Then
        Man_Auto_lbl(i).Caption = "M"
    Else
        Man_Auto_lbl(i).Caption = "A"
    End If
Next i

' Plant Master...
If SI_units Then
            LH_scale(10).Caption = " 0 to 104 Bar "
            RH_scale(10).Caption = " 0 to 104 Bar "
Else
            LH_scale(10).Caption = " 0 to 1500 psi "
            RH_scale(10).Caption = " 0 to 1500 psi "
End If

End Sub
