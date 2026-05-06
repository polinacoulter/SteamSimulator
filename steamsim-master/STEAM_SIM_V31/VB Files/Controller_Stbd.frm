VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form Controller_Stbd 
   Caption         =   "Stbd Control"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   14565
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   10095
      Left            =   0
      Top             =   60
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   17806
      Begin VB.TextBox SP_auto_fuel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   10020
         TabIndex        =   102
         Text            =   "SP_auto_fuel"
         Top             =   3540
         Width           =   1035
      End
      Begin VB.TextBox PID_out_fuel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   12300
         TabIndex        =   101
         Text            =   "PID_out_fuel"
         Top             =   5820
         Width           =   975
      End
      Begin VB.TextBox SP_man_air 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   100
         Text            =   "SP_man_air"
         Top             =   4860
         Width           =   1035
      End
      Begin VB.TextBox SP_auto_air 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   99
         Text            =   "SP_auto_air"
         Top             =   3960
         Width           =   1035
      End
      Begin VB.TextBox SP_AIR_low_fire 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3660
         TabIndex        =   98
         Text            =   "SP_AIR_low_fire"
         Top             =   60
         Width           =   1035
      End
      Begin VB.TextBox PID_out_air 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2940
         TabIndex        =   97
         Text            =   "PID_out_air"
         Top             =   5700
         Width           =   975
      End
      Begin VB.TextBox mv_air 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   96
         Text            =   "mv_air"
         Top             =   4440
         Width           =   1035
      End
      Begin VB.TextBox mv_fuel 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   9120
         TabIndex        =   95
         Text            =   "mv_fuel"
         Top             =   4860
         Width           =   1035
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   5880
         Top             =   3840
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         Height          =   435
         Left            =   7680
         TabIndex        =   84
         Top             =   9420
         Width           =   675
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Left            =   11520
         TabIndex        =   47
         Top             =   3600
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   2700
         TabIndex        =   17
         Top             =   3600
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   900
         TabIndex        =   13
         Top             =   720
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   3300
         TabIndex        =   0
         Top             =   900
         Width           =   255
      End
      Begin VB.Line Line80 
         X1              =   4080
         X2              =   4080
         Y1              =   4920
         Y2              =   4740
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   1260
         X2              =   1260
         Y1              =   3060
         Y2              =   4140
      End
      Begin VB.Line Air_Man_line 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4260
         Y1              =   4560
         Y2              =   4740
      End
      Begin VB.Line Line32 
         BorderWidth     =   2
         X1              =   1260
         X2              =   1260
         Y1              =   4620
         Y2              =   5460
      End
      Begin VB.Line Line86 
         X1              =   1080
         X2              =   1320
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Line Line85 
         X1              =   9600
         X2              =   9600
         Y1              =   4620
         Y2              =   4980
      End
      Begin VB.Line Line84 
         X1              =   1080
         X2              =   1320
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line83 
         X1              =   3840
         X2              =   4560
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Shape Air_Contact_shape 
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   1
         Left            =   5880
         Top             =   600
         Width           =   75
      End
      Begin VB.Shape Fuel_Contact_shape 
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   1
         Left            =   13140
         Top             =   7020
         Width           =   75
      End
      Begin VB.Label Label67 
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
         Left            =   360
         TabIndex        =   94
         Top             =   7620
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   435
         Left            =   1320
         Top             =   1680
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   17
         Left            =   960
         Picture         =   "Controller_Stbd.frx":0000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   14
         Left            =   8340
         Picture         =   "Controller_Stbd.frx":030A
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   13
         Left            =   8520
         Picture         =   "Controller_Stbd.frx":0614
         Stretch         =   -1  'True
         Top             =   5460
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   12
         Left            =   8160
         Picture         =   "Controller_Stbd.frx":091E
         Stretch         =   -1  'True
         Top             =   5460
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   16
         Left            =   7800
         Picture         =   "Controller_Stbd.frx":0C28
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   7800
         Picture         =   "Controller_Stbd.frx":0F32
         Stretch         =   -1  'True
         Top             =   5340
         Width           =   255
      End
      Begin VB.Label Label66 
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
         Left            =   12360
         TabIndex        =   93
         Top             =   4860
         Width           =   495
      End
      Begin VB.Label Label65 
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
         Left            =   3900
         TabIndex        =   92
         Top             =   2220
         Width           =   495
      End
      Begin VB.Line Line82 
         BorderWidth     =   2
         X1              =   4980
         X2              =   4740
         Y1              =   1860
         Y2              =   1680
      End
      Begin VB.Line Line81 
         BorderWidth     =   2
         X1              =   13380
         X2              =   13140
         Y1              =   4560
         Y2              =   4380
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
         Left            =   3540
         TabIndex        =   91
         Top             =   4620
         Width           =   495
      End
      Begin VB.Line Air_Auto_line 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4320
         Y1              =   4560
         Y2              =   4380
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   15
         Left            =   5100
         Picture         =   "Controller_Stbd.frx":123C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         Height          =   555
         Left            =   5700
         Top             =   1800
         Width           =   855
      End
      Begin VB.Line Line79 
         BorderWidth     =   2
         X1              =   6120
         X2              =   5940
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOW Fire Indicator Lamp"
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
         Left            =   9000
         TabIndex        =   90
         Top             =   8640
         Width           =   2595
      End
      Begin VB.Shape Logic_alarm 
         BorderWidth     =   2
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   1
         Left            =   11640
         Shape           =   3  'Circle
         Top             =   8640
         Width           =   255
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "24 V (Purge)"
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
         Left            =   900
         TabIndex        =   89
         Top             =   6900
         Width           =   1575
      End
      Begin VB.Line Line78 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   840
         X2              =   4140
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Boiler Steam Pressure LOW"
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
         Left            =   6180
         TabIndex        =   88
         Top             =   7500
         Width           =   2595
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1st Burner in Service 10 sec"
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
         Left            =   6180
         TabIndex        =   87
         Top             =   6840
         Width           =   2595
      End
      Begin VB.Line Line77 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   11760
         X2              =   11460
         Y1              =   7620
         Y2              =   7620
      End
      Begin VB.Line Line76 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   11760
         X2              =   11760
         Y1              =   8580
         Y2              =   7620
      End
      Begin VB.Line Line75 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5280
         X2              =   5280
         Y1              =   5340
         Y2              =   5160
      End
      Begin VB.Line Line74 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5280
         X2              =   5280
         Y1              =   8460
         Y2              =   5640
      End
      Begin VB.Line Line73 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5700
         X2              =   5280
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line72 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5280
         X2              =   5280
         Y1              =   3180
         Y2              =   4920
      End
      Begin VB.Line Line71 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   5280
         X2              =   5280
         Y1              =   2940
         Y2              =   780
      End
      Begin VB.Line Line70 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   11760
         X2              =   5280
         Y1              =   8460
         Y2              =   8460
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   10980
         X2              =   10980
         Y1              =   7260
         Y2              =   7980
      End
      Begin VB.Line Line69 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   6360
         X2              =   10980
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line68 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   10800
         X2              =   10800
         Y1              =   7440
         Y2              =   7200
      End
      Begin VB.Line Line60 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   10980
         X2              =   10800
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line59 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   9900
         X2              =   6300
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Shape Shape20 
         BorderWidth     =   2
         Height          =   375
         Left            =   9900
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "NOT"
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
         Left            =   10080
         TabIndex        =   86
         Top             =   7020
         Width           =   375
      End
      Begin VB.Line Line58 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   12960
         X2              =   10620
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "OR"
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
         Left            =   11040
         TabIndex        =   85
         Top             =   7500
         Width           =   375
      End
      Begin VB.Shape Shape18 
         BorderWidth     =   2
         Height          =   495
         Left            =   10980
         Shape           =   3  'Circle
         Top             =   7380
         Width           =   495
      End
      Begin VB.Label Fuel_pos_stbd 
         Alignment       =   2  'Center
         Caption         =   "100 %"
         Height          =   255
         Index           =   2
         Left            =   12000
         TabIndex        =   83
         Top             =   9780
         Width           =   615
      End
      Begin VB.Label FD_position_label 
         Alignment       =   2  'Center
         Caption         =   "100 %"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   82
         Top             =   9780
         Width           =   615
      End
      Begin VB.Line Line63 
         BorderWidth     =   2
         Index           =   1
         X1              =   5820
         X2              =   5820
         Y1              =   180
         Y2              =   660
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Fuel Valve"
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
         Left            =   11040
         TabIndex        =   81
         Top             =   9480
         Width           =   1875
      End
      Begin VB.Line Line63 
         BorderWidth     =   2
         Index           =   3
         X1              =   13080
         X2              =   12720
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Low Fire Relay FUEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   2
         Left            =   13560
         TabIndex        =   79
         Top             =   6900
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   14
         Left            =   12780
         Picture         =   "Controller_Stbd.frx":1546
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   255
      End
      Begin VB.Line Line62 
         BorderWidth     =   2
         Index           =   1
         X1              =   13260
         X2              =   13260
         Y1              =   7920
         Y2              =   7680
      End
      Begin VB.Label Label50 
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
         Index           =   2
         Left            =   11520
         TabIndex        =   80
         Top             =   6360
         Width           =   1155
      End
      Begin VB.Shape Circ_shape_left 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   13020
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   135
      End
      Begin VB.Shape Circ_shape_right 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   13320
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   135
      End
      Begin VB.Shape Shape16 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   13020
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   135
      End
      Begin VB.Shape Shape17 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   13320
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   135
      End
      Begin VB.Shape Fuel_Contact_shape 
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   0
         Left            =   13260
         Top             =   7020
         Width           =   75
      End
      Begin VB.Line Line66 
         BorderWidth     =   2
         Index           =   2
         X1              =   13080
         X2              =   13080
         Y1              =   7380
         Y2              =   7680
      End
      Begin VB.Line Line63 
         BorderWidth     =   2
         Index           =   2
         X1              =   13080
         X2              =   13080
         Y1              =   6480
         Y2              =   7080
      End
      Begin VB.Shape Shape19 
         BorderStyle     =   3  'Dot
         Height          =   555
         Index           =   2
         Left            =   12960
         Top             =   6960
         Width           =   555
      End
      Begin VB.Line Line65 
         BorderWidth     =   2
         Index           =   2
         X1              =   13380
         X2              =   13080
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Line Line67 
         BorderWidth     =   2
         Index           =   2
         X1              =   13380
         X2              =   13380
         Y1              =   7440
         Y2              =   7680
      End
      Begin VB.Line Line64 
         BorderWidth     =   2
         Index           =   2
         X1              =   13380
         X2              =   13380
         Y1              =   5100
         Y2              =   7080
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "E/I"
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
         Index           =   1
         Left            =   13020
         TabIndex        =   78
         Top             =   7980
         Width           =   435
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         Index           =   1
         X1              =   13260
         X2              =   13260
         Y1              =   8580
         Y2              =   8340
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I/P"
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
         Index           =   1
         Left            =   13020
         TabIndex        =   77
         Top             =   8640
         Width           =   435
      End
      Begin VB.Shape Shape15 
         BorderWidth     =   2
         Height          =   435
         Index           =   1
         Left            =   12960
         Top             =   8580
         Width           =   555
      End
      Begin VB.Label Label54 
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
         Index           =   1
         Left            =   12000
         TabIndex        =   76
         Top             =   8340
         Width           =   915
      End
      Begin VB.Label Label58 
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
         Index           =   1
         Left            =   12300
         TabIndex        =   75
         Top             =   7740
         Width           =   615
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         Height          =   435
         Index           =   1
         Left            =   12960
         Top             =   7920
         Width           =   555
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15 psi"
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
         Left            =   12000
         TabIndex        =   74
         Top             =   9060
         Width           =   915
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         Index           =   1
         X1              =   13260
         X2              =   13260
         Y1              =   9420
         Y2              =   9000
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   11
         Left            =   13140
         Picture         =   "Controller_Stbd.frx":1850
         Stretch         =   -1  'True
         Top             =   9120
         Width           =   255
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         Height          =   555
         Left            =   4140
         Top             =   5940
         Width           =   855
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   10
         Left            =   4320
         Picture         =   "Controller_Stbd.frx":1B5A
         Stretch         =   -1  'True
         Top             =   9120
         Width           =   255
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Air Damper"
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
         Left            =   2040
         TabIndex        =   73
         Top             =   9480
         Width           =   1875
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3960
         OleObjectBlob   =   "Controller_Stbd.frx":1E64
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   72
         Top             =   9420
         Width           =   915
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         Index           =   0
         X1              =   4440
         X2              =   4440
         Y1              =   9420
         Y2              =   9000
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "15 psi"
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
         Left            =   3180
         TabIndex        =   71
         Top             =   9060
         Width           =   915
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Index           =   0
         Left            =   13020
         OleObjectBlob   =   "Controller_Stbd.frx":3E7C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   70
         Top             =   9420
         Width           =   510
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         Height          =   435
         Index           =   0
         Left            =   4140
         Top             =   7920
         Width           =   555
      End
      Begin VB.Label Label58 
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
         Index           =   0
         Left            =   3420
         TabIndex        =   69
         Top             =   7680
         Width           =   615
      End
      Begin VB.Label Label54 
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
         Index           =   0
         Left            =   3180
         TabIndex        =   68
         Top             =   8340
         Width           =   915
      End
      Begin VB.Shape Shape15 
         BorderWidth     =   2
         Height          =   435
         Index           =   0
         Left            =   4140
         Top             =   8580
         Width           =   555
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "I/P"
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
         Index           =   0
         Left            =   4200
         TabIndex        =   67
         Top             =   8640
         Width           =   435
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         Index           =   0
         X1              =   4440
         X2              =   4440
         Y1              =   8580
         Y2              =   8340
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "E/I"
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
         Index           =   0
         Left            =   4200
         TabIndex        =   66
         Top             =   7980
         Width           =   435
      End
      Begin VB.Line Line24 
         BorderWidth     =   2
         X1              =   1260
         X2              =   6120
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   5700
         X2              =   4980
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line48 
         BorderWidth     =   2
         X1              =   6300
         X2              =   6300
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Shape Shape19 
         BorderStyle     =   3  'Dot
         Height          =   555
         Index           =   1
         Left            =   5700
         Top             =   540
         Width           =   555
      End
      Begin VB.Line Line67 
         BorderWidth     =   2
         Index           =   1
         X1              =   6120
         X2              =   6120
         Y1              =   1020
         Y2              =   1260
      End
      Begin VB.Line Line66 
         BorderWidth     =   2
         Index           =   1
         X1              =   5820
         X2              =   5820
         Y1              =   960
         Y2              =   1260
      End
      Begin VB.Line Line65 
         BorderWidth     =   2
         Index           =   1
         X1              =   6120
         X2              =   5820
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line64 
         BorderWidth     =   2
         Index           =   1
         X1              =   6120
         X2              =   6120
         Y1              =   360
         Y2              =   660
      End
      Begin VB.Shape Air_Contact_shape 
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   0
         Left            =   6000
         Top             =   600
         Width           =   75
      End
      Begin VB.Shape Shape17 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6060
         Shape           =   3  'Circle
         Top             =   900
         Width           =   135
      End
      Begin VB.Shape Shape16 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5760
         Shape           =   3  'Circle
         Top             =   900
         Width           =   135
      End
      Begin VB.Shape Circ_shape_right 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6060
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Circ_shape_left 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5760
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Low Fire Relay AIR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   6300
         TabIndex        =   65
         Top             =   360
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line61 
         BorderWidth     =   2
         Index           =   1
         X1              =   5820
         X2              =   4680
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   4
         Left            =   5280
         Picture         =   "Controller_Stbd.frx":C494
         Stretch         =   -1  'True
         Top             =   60
         Width           =   255
      End
      Begin VB.Line Line64 
         BorderWidth     =   2
         Index           =   0
         X1              =   4560
         X2              =   4560
         Y1              =   6480
         Y2              =   7080
      End
      Begin VB.Line Line67 
         BorderWidth     =   2
         Index           =   0
         X1              =   4560
         X2              =   4560
         Y1              =   7440
         Y2              =   7680
      End
      Begin VB.Line Line65 
         BorderWidth     =   2
         Index           =   0
         X1              =   4560
         X2              =   4260
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Shape Shape19 
         BorderStyle     =   3  'Dot
         Height          =   555
         Index           =   0
         Left            =   4140
         Top             =   6960
         Width           =   555
      End
      Begin VB.Line Line63 
         BorderWidth     =   2
         Index           =   0
         X1              =   4260
         X2              =   4260
         Y1              =   6720
         Y2              =   7080
      End
      Begin VB.Line Line66 
         BorderWidth     =   2
         Index           =   0
         X1              =   4260
         X2              =   4260
         Y1              =   7380
         Y2              =   7680
      End
      Begin VB.Shape Contact_shape 
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   0
         Left            =   4440
         Top             =   7020
         Width           =   75
      End
      Begin VB.Shape Shape17 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   135
      End
      Begin VB.Shape Shape16 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4200
         Shape           =   3  'Circle
         Top             =   7320
         Width           =   135
      End
      Begin VB.Shape Circ_shape_right 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   135
      End
      Begin VB.Shape Circ_shape_left 
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4200
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   135
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Purge Relay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2640
         TabIndex        =   64
         Top             =   7260
         Width           =   1455
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10 V (Purge)"
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
         Left            =   2400
         TabIndex        =   63
         Top             =   6600
         Width           =   1155
      End
      Begin VB.Line Line62 
         BorderWidth     =   2
         Index           =   0
         X1              =   4440
         X2              =   4440
         Y1              =   7920
         Y2              =   7680
      End
      Begin VB.Line Line61 
         BorderWidth     =   2
         Index           =   0
         X1              =   4260
         X2              =   3660
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   13
         Left            =   3720
         Picture         =   "Controller_Stbd.frx":C79E
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   255
      End
      Begin VB.Line Line57 
         BorderWidth     =   2
         X1              =   4500
         X2              =   4500
         Y1              =   6060
         Y2              =   6360
      End
      Begin VB.OLE OLE13 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   675
         Left            =   4320
         OleObjectBlob   =   "Controller_Stbd.frx":CAA8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Text\Less_than_48.doc"
         TabIndex        =   62
         Top             =   5880
         Width           =   495
      End
      Begin VB.Line Line34 
         BorderWidth     =   2
         X1              =   8280
         X2              =   8280
         Y1              =   5700
         Y2              =   3900
      End
      Begin VB.Line Line56 
         BorderWidth     =   2
         X1              =   10260
         X2              =   10260
         Y1              =   4500
         Y2              =   4140
      End
      Begin VB.Line Line55 
         BorderWidth     =   2
         X1              =   10260
         X2              =   10260
         Y1              =   6480
         Y2              =   4740
      End
      Begin VB.Line Line54 
         BorderWidth     =   2
         X1              =   9960
         X2              =   10260
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line53 
         BorderWidth     =   2
         X1              =   8460
         X2              =   9060
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         Height          =   555
         Left            =   9060
         Top             =   6180
         Width           =   915
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "f(x)"
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
         Left            =   9240
         TabIndex        =   61
         Top             =   6300
         Width           =   495
      End
      Begin VB.Line Line52 
         BorderWidth     =   2
         X1              =   8460
         X2              =   8460
         Y1              =   6480
         Y2              =   6240
      End
      Begin VB.Line Line51 
         BorderWidth     =   2
         X1              =   8640
         X2              =   8640
         Y1              =   5100
         Y2              =   5700
      End
      Begin VB.Line Line50 
         BorderWidth     =   2
         X1              =   8400
         X2              =   8640
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Line49 
         BorderWidth     =   2
         X1              =   4980
         X2              =   8160
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Line47 
         BorderWidth     =   2
         X1              =   4980
         X2              =   4980
         Y1              =   5100
         Y2              =   3240
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         Height          =   555
         Left            =   8040
         Top             =   5700
         Width           =   855
      End
      Begin VB.Label Label30 
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
         Left            =   1380
         TabIndex        =   60
         Top             =   1740
         Width           =   435
      End
      Begin VB.Label Label47 
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
         Left            =   420
         TabIndex        =   59
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label46 
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
         Left            =   12660
         TabIndex        =   58
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label41 
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
         Left            =   12240
         TabIndex        =   57
         Top             =   4620
         Width           =   615
      End
      Begin VB.Line Line46 
         BorderWidth     =   2
         X1              =   11100
         X2              =   11100
         Y1              =   3660
         Y2              =   4980
      End
      Begin VB.Line Line45 
         BorderWidth     =   2
         X1              =   10080
         X2              =   11100
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Label Label40 
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
         Left            =   10080
         TabIndex        =   56
         Top             =   3840
         Width           =   675
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
         Left            =   10200
         TabIndex        =   55
         Top             =   4680
         Width           =   675
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   12
         Left            =   12900
         Picture         =   "Controller_Stbd.frx":118C0
         Stretch         =   -1  'True
         Top             =   4620
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   11
         Left            =   10860
         Picture         =   "Controller_Stbd.frx":11BCA
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   10
         Left            =   10860
         Picture         =   "Controller_Stbd.frx":11ED4
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   9
         Left            =   12900
         Picture         =   "Controller_Stbd.frx":121DE
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Line Line44 
         BorderWidth     =   2
         X1              =   12600
         X2              =   13080
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label38 
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
         Left            =   12660
         TabIndex        =   54
         Top             =   4080
         Width           =   795
      End
      Begin VB.Line Line43 
         BorderWidth     =   2
         X1              =   10260
         X2              =   11100
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Fuel"
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
         Left            =   10440
         TabIndex        =   53
         Top             =   3180
         Width           =   2295
      End
      Begin VB.Label Label34 
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
         Left            =   10800
         TabIndex        =   52
         Top             =   4260
         Width           =   255
      End
      Begin VB.Label Label33 
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
         Left            =   10800
         TabIndex        =   51
         Top             =   3780
         Width           =   255
      End
      Begin VB.Line Line42 
         BorderWidth     =   2
         X1              =   11100
         X2              =   12660
         Y1              =   4980
         Y2              =   4380
      End
      Begin VB.Line Line41 
         BorderWidth     =   2
         X1              =   11100
         X2              =   12660
         Y1              =   3660
         Y2              =   4380
      End
      Begin VB.OLE OLE12 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11205
         OleObjectBlob   =   "Controller_Stbd.frx":124E8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Percent_sign_18.doc"
         TabIndex        =   50
         Top             =   4140
         Width           =   255
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
         Left            =   11820
         OleObjectBlob   =   "Controller_Stbd.frx":17500
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Delta_sign_18.doc"
         TabIndex        =   49
         Top             =   4140
         Width           =   255
      End
      Begin VB.OLE OLE9 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   495
         Left            =   11580
         OleObjectBlob   =   "Controller_Stbd.frx":1C518
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Integral_sign_18.doc"
         TabIndex        =   48
         Top             =   4140
         Width           =   255
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   13380
         X2              =   13380
         Y1              =   5220
         Y2              =   4560
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   9
         Left            =   13260
         Picture         =   "Controller_Stbd.frx":21530
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Line Line39 
         BorderWidth     =   2
         X1              =   8400
         X2              =   10320
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   7800
         Picture         =   "Controller_Stbd.frx":2183A
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   8580
         Picture         =   "Controller_Stbd.frx":21B44
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   255
      End
      Begin VB.Line Line33 
         BorderWidth     =   2
         X1              =   1260
         X2              =   8280
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label Label29 
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
         Left            =   5400
         TabIndex        =   46
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label Label23 
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
         Left            =   6480
         TabIndex        =   45
         Top             =   1500
         Width           =   795
      End
      Begin VB.Line Line38 
         BorderWidth     =   2
         X1              =   8160
         X2              =   7200
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   8
         Left            =   9780
         Picture         =   "Controller_Stbd.frx":21E4E
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Line Line40 
         BorderWidth     =   2
         X1              =   9900
         X2              =   9900
         Y1              =   3120
         Y2              =   4620
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   5
         Left            =   8160
         Picture         =   "Controller_Stbd.frx":22158
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Line Line37 
         BorderWidth     =   2
         X1              =   7200
         X2              =   7200
         Y1              =   1440
         Y2              =   4620
      End
      Begin VB.Line Line28 
         BorderWidth     =   2
         X1              =   8280
         X2              =   8280
         Y1              =   2580
         Y2              =   2160
      End
      Begin VB.Label Label22 
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
         Left            =   7260
         TabIndex        =   44
         Top             =   4020
         Width           =   915
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Air Flow"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   43
         Top             =   4020
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   4
         Left            =   4440
         Picture         =   "Controller_Stbd.frx":22462
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   3
         Left            =   4440
         Picture         =   "Controller_Stbd.frx":2276C
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Line Line36 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   5940
         Y2              =   5580
      End
      Begin VB.Line Line35 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   5340
         Y2              =   4560
      End
      Begin VB.Label Label27 
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
         Left            =   9600
         TabIndex        =   42
         Top             =   900
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   435
         Left            =   9600
         Top             =   1740
         Width           =   555
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
         Left            =   9660
         TabIndex        =   41
         Top             =   1800
         Width           =   435
      End
      Begin VB.Line Line31 
         BorderWidth     =   2
         X1              =   9900
         X2              =   9900
         Y1              =   1740
         Y2              =   1260
      End
      Begin VB.Label Label25 
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
         Left            =   8880
         TabIndex        =   40
         Top             =   2220
         Width           =   915
      End
      Begin VB.Label Label24 
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
         Left            =   8820
         TabIndex        =   39
         Top             =   1380
         Width           =   915
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   435
         Left            =   9600
         Top             =   840
         Width           =   555
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   2
         Left            =   9780
         Picture         =   "Controller_Stbd.frx":22A76
         Stretch         =   -1  'True
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2.0 lb/s"
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
         Left            =   9360
         TabIndex        =   38
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "f(t)"
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
         Left            =   9660
         TabIndex        =   37
         Top             =   2700
         Width           =   495
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   555
         Left            =   9480
         Top             =   2580
         Width           =   915
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   9900
         X2              =   9900
         Y1              =   2580
         Y2              =   2160
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Flow"
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
         Left            =   9480
         TabIndex        =   36
         Top             =   180
         Width           =   1395
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   555
         Left            =   7860
         Top             =   2580
         Width           =   915
      End
      Begin VB.OLE OLE8 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   675
         Left            =   8220
         OleObjectBlob   =   "Controller_Stbd.frx":22D80
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Text\Less_than_48.doc"
         TabIndex        =   35
         Top             =   5640
         Width           =   495
      End
      Begin VB.Line Line29 
         BorderWidth     =   2
         X1              =   8280
         X2              =   8280
         Y1              =   3360
         Y2              =   3120
      End
      Begin VB.Label Label18 
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
         Left            =   7920
         TabIndex        =   34
         Top             =   3480
         Width           =   315
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   555
         Left            =   7860
         Top             =   3360
         Width           =   915
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "f(t)"
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
         Left            =   8040
         TabIndex        =   33
         Top             =   2700
         Width           =   495
      End
      Begin VB.OLE OLE10 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   495
         Left            =   8220
         OleObjectBlob   =   "Controller_Stbd.frx":27B98
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Square_root_18.doc"
         TabIndex        =   32
         Top             =   3420
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12 in H2O"
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
         Left            =   7620
         TabIndex        =   31
         Top             =   540
         Width           =   1095
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   6
         Left            =   8160
         Picture         =   "Controller_Stbd.frx":2CBB0
         Stretch         =   -1  'True
         Top             =   1500
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   435
         Left            =   7980
         Top             =   840
         Width           =   555
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
         Left            =   7200
         TabIndex        =   30
         Top             =   1380
         Width           =   915
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
         Left            =   7260
         TabIndex        =   29
         Top             =   2220
         Width           =   915
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   8280
         X2              =   8280
         Y1              =   1740
         Y2              =   1260
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
         Left            =   8040
         TabIndex        =   28
         Top             =   1800
         Width           =   435
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         Height          =   435
         Left            =   7980
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Label15 
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
         Left            =   7980
         TabIndex        =   27
         Top             =   900
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Air Pressure"
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
         TabIndex        =   26
         Top             =   180
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   6540
         Picture         =   "Controller_Stbd.frx":2CEBA
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   1
         Left            =   5820
         Picture         =   "Controller_Stbd.frx":2D1C4
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label14 
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
         Left            =   6180
         TabIndex        =   25
         Top             =   2400
         Width           =   795
      End
      Begin VB.Line Line25 
         BorderWidth     =   2
         X1              =   6120
         X2              =   6120
         Y1              =   3060
         Y2              =   2340
      End
      Begin VB.OLE OLE7 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         OleObjectBlob   =   "Controller_Stbd.frx":2D4CE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Integral_sign_18.doc"
         TabIndex        =   24
         Top             =   4140
         Width           =   255
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
         Left            =   3000
         OleObjectBlob   =   "Controller_Stbd.frx":324E6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Delta_sign_18.doc"
         TabIndex        =   23
         Top             =   4140
         Width           =   255
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2385
         OleObjectBlob   =   "Controller_Stbd.frx":374FE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Percent_sign_18.doc"
         TabIndex        =   22
         Top             =   4140
         Width           =   255
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   2280
         X2              =   3840
         Y1              =   3660
         Y2              =   4380
      End
      Begin VB.Line Line22 
         BorderWidth     =   2
         X1              =   2280
         X2              =   3840
         Y1              =   4980
         Y2              =   4380
      End
      Begin VB.Label Label13 
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
         Left            =   1980
         TabIndex        =   21
         Top             =   3780
         Width           =   255
      End
      Begin VB.Label Label12 
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
         Left            =   1980
         TabIndex        =   20
         Top             =   4260
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Air"
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
         Left            =   1620
         TabIndex        =   19
         Top             =   3180
         Width           =   2295
      End
      Begin VB.Line Line21 
         BorderWidth     =   2
         X1              =   1260
         X2              =   2280
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   3780
         X2              =   4260
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   8
         Left            =   4080
         Picture         =   "Controller_Stbd.frx":3C516
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   7
         Left            =   2040
         Picture         =   "Controller_Stbd.frx":3C820
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   6
         Left            =   2040
         Picture         =   "Controller_Stbd.frx":3CB2A
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   5
         Left            =   4080
         Picture         =   "Controller_Stbd.frx":3CE34
         Stretch         =   -1  'True
         Top             =   4620
         Width           =   255
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   1260
         X2              =   2280
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   2280
         X2              =   2280
         Y1              =   3660
         Y2              =   4980
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
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
         Left            =   3840
         TabIndex        =   18
         Top             =   4020
         Width           =   495
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   7200
         X2              =   6300
         Y1              =   1440
         Y2              =   1440
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
         Left            =   6000
         OleObjectBlob   =   "Controller_Stbd.frx":3D13E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Greater_than_48.doc"
         TabIndex        =   16
         Top             =   1680
         Width           =   315
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   5940
         X2              =   5940
         Y1              =   1800
         Y2              =   1260
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   4500
         X2              =   4680
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   4980
         X2              =   4980
         Y1              =   2880
         Y2              =   360
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
         Left            =   0
         TabIndex        =   15
         Top             =   1440
         Width           =   915
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         Height          =   435
         Left            =   240
         Top             =   1680
         Width           =   555
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
         Left            =   240
         TabIndex        =   14
         Top             =   1740
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
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
         Left            =   960
         TabIndex        =   12
         Top             =   60
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   780
         X2              =   1320
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   1260
         X2              =   1260
         Y1              =   1440
         Y2              =   780
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   7
         Left            =   6000
         Picture         =   "Controller_Stbd.frx":42156
         Stretch         =   -1  'True
         Top             =   2700
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   0
         Left            =   1140
         Picture         =   "Controller_Stbd.frx":42460
         Stretch         =   -1  'True
         Top             =   840
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
         Left            =   3360
         OleObjectBlob   =   "Controller_Stbd.frx":4276A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Integral_sign_18.doc"
         TabIndex        =   11
         Top             =   1440
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
         Left            =   3600
         OleObjectBlob   =   "Controller_Stbd.frx":47782
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Delta_sign_18.doc"
         TabIndex        =   10
         Top             =   1440
         Width           =   255
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
         Left            =   2985
         OleObjectBlob   =   "Controller_Stbd.frx":4C79A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Percent_sign_18.doc"
         TabIndex        =   9
         Top             =   1440
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2880
         X2              =   4440
         Y1              =   960
         Y2              =   1680
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   2880
         X2              =   4440
         Y1              =   2280
         Y2              =   1680
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
         Left            =   2580
         TabIndex        =   8
         Top             =   1080
         Width           =   255
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
         Left            =   2580
         TabIndex        =   7
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Master"
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
         Left            =   2220
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   1260
         X2              =   2880
         Y1              =   1440
         Y2              =   1440
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
         Left            =   4440
         TabIndex        =   5
         Top             =   1380
         Width           =   795
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   4440
         X2              =   4680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   2
         Left            =   4500
         Picture         =   "Controller_Stbd.frx":517B2
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   0
         Left            =   2640
         Picture         =   "Controller_Stbd.frx":51ABC
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   1
         Left            =   2640
         Picture         =   "Controller_Stbd.frx":51DC6
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   3
         Left            =   4440
         Picture         =   "Controller_Stbd.frx":520D0
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
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
         Left            =   1980
         TabIndex        =   4
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
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
         Left            =   1980
         TabIndex        =   3
         Top             =   1140
         Width           =   615
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   1860
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   2880
         X2              =   2880
         Y1              =   960
         Y2              =   2280
      End
      Begin VB.Label Label55 
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
         Left            =   3780
         TabIndex        =   2
         Top             =   1980
         Width           =   615
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
         Left            =   4440
         TabIndex        =   1
         Top             =   1140
         Width           =   495
      End
   End
End
Attribute VB_Name = "Controller_Stbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Controller_MASTER_STBD.ZOrder

End Sub

Private Sub Command2_Click()

Controller_AIR_STBD.ZOrder

End Sub

Private Sub Command3_Click()


' === 5/03/02 ===
Controllers_combustion.ZOrder

End Sub

Private Sub Command4_Click()

Controller_FUEL_STBD.ZOrder

End Sub

Private Sub Command5_Click()

Key_form.ZOrder

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_controller_stbd)


' === 05/03/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_controller_stbd)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Timer1_Timer()


' === 26/09/02  ===
If VB_steam_integers(39) = 2 Or VB_steam_integers(39) = 4 Then
    ' AUTO or VLV is chosen..
    Air_Auto_line.Visible = True
    Air_Man_line.Visible = False
Else
    ' MAN or SP is chosen..
    Air_Auto_line.Visible = False
    Air_Man_line.Visible = True
End If


' === 26/09/02  ===
'  SP_AIR_low_fire(1) = ALARM_HILO(2, 250) * 0.01
SP_AIR_low_fire = Format(ALARM_HILO(2, 250) * 0.1, "##0.00") & " V"


' === 01/10/02 ===
'PC_REALS(523) = SP_auto_fuel(1) * 255#
SP_auto_fuel.text = Format(LSS_reals(523) / 25.5, "##0.00") & " V"


' === 26/09/02  ===
'PC_REALS(527) = SP_auto_air(1) * 255#
SP_auto_air.text = Format(LSS_reals(527) / 25.5, "##0.00") & " V"

'PC_REALS(529) = SP_man_air(1) * 255#
SP_man_air.text = Format(LSS_reals(529) / 25.5, "##0.00") & " V"


' PC_REALS(834) = PID_out_air(1)
PID_out_air.text = Format(LSS_reals(834) * 10#, "##0.00") & " V"

' === 01/10/02 ===
'PC_REALS(835) = PID_out_fuel(1)
PID_out_fuel.text = Format(LSS_reals(835) * 10#, "##0.00") & " V"


'          PC_REALS(832) = mV_fuel(1)
'          PC_REALS(833) = mV_air(1)
mv_fuel.text = Format(LSS_reals(832) * 10#, "##0.00") & " V"
mv_air.text = Format(LSS_reals(833) * 10#, "##0.00") & " V"


'equivalence(STEAM_logicals(109),LOW_FIRE_FUEL(1))
If VB_steam_logicals(109) Then
    Fuel_Contact_shape(0).Visible = False
    Fuel_Contact_shape(1).Visible = True
Else
    Fuel_Contact_shape(0).Visible = True
    Fuel_Contact_shape(1).Visible = False
End If


If VB_steam_logicals(130) Then
  'For i = 0 To 4
  i = 1
    If VB_steam_logicals(61 + i) Then
        Logic_alarm(i).FillColor = &HFF&            ' BRIGHT RED
        Air_Contact_shape(0).Visible = False
        Air_Contact_shape(1).Visible = True
    Else
        Logic_alarm(i).FillColor = &H80&            ' DARK RED
        Air_Contact_shape(0).Visible = True
        Air_Contact_shape(1).Visible = False
    End If
  'Next i
Else
  'For i = 0 To 4
  i = 1
        Logic_alarm(i).FillColor = &H80&            ' DARK RED
  'Next i
  
        Air_Contact_shape(0).Visible = True
        Air_Contact_shape(1).Visible = False

End If



End Sub
