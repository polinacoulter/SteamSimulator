VERSION 5.00
Begin VB.Form frm_power 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Type 42 Power Distribution"
   ClientHeight    =   11115
   ClientLeft      =   -1920
   ClientTop       =   1170
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "T42BUS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Begin VB.PictureBox PicClip2 
      Height          =   975
      Index           =   7
      Left            =   16080
      Picture         =   "T42BUS.frx":030A
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   21
      Top             =   11760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicClip2 
      Height          =   975
      Index           =   6
      Left            =   16080
      Picture         =   "T42BUS.frx":0B8C
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   20
      Top             =   10680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicClip2 
      Height          =   975
      Index           =   5
      Left            =   14880
      Picture         =   "T42BUS.frx":140E
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   19
      Top             =   11760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer power_system_timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   600
   End
   Begin VB.PictureBox PicClip2 
      Height          =   975
      Index           =   4
      Left            =   14880
      Picture         =   "T42BUS.frx":1C90
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   10680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicClip2 
      Height          =   1455
      Index           =   3
      Left            =   15600
      Picture         =   "T42BUS.frx":2512
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PicClip2 
      Height          =   1575
      Index           =   2
      Left            =   15600
      Picture         =   "T42BUS.frx":3794
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PicClip2 
      Height          =   1455
      Index           =   1
      Left            =   14160
      Picture         =   "T42BUS.frx":4A16
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PicClip2 
      Height          =   1575
      Index           =   0
      Left            =   14160
      Picture         =   "T42BUS.frx":5C98
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   115
      Visible         =   0   'False
      X1              =   2400
      X2              =   2400
      Y1              =   3120
      Y2              =   1920
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   43
      X1              =   5280
      X2              =   5280
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   41
      X1              =   5085
      X2              =   5085
      Y1              =   4125
      Y2              =   3360
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   39
      X1              =   4890
      X2              =   4890
      Y1              =   3915
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   6840
      X2              =   2880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   33
      X1              =   2685
      X2              =   6840
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   18
      X1              =   2475
      X2              =   4920
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   27
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":6F1A
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   26
      Left            =   4680
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":779C
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   25
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":801E
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   24
      Left            =   12600
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":88A0
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   23
      Left            =   12600
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":9122
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   22
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":99A4
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   21
      Left            =   4680
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":A226
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   20
      Left            =   2280
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":AAA8
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   19
      Left            =   2280
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":B32A
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   18
      Left            =   6840
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":BBAC
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   17
      Left            =   10800
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":C42E
      Stretch         =   -1  'True
      Top             =   9360
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   16
      Left            =   15240
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":CCB0
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   15
      Left            =   15240
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":D532
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   14
      Left            =   10800
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":DDB4
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   13
      Left            =   6840
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":E636
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   12
      Left            =   4680
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":EEB8
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   11
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":F73A
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   10
      Left            =   12600
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":FFBC
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   9
      Left            =   12600
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":1083E
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   8
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":110C0
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   114
      X1              =   15435
      X2              =   15435
      Y1              =   5640
      Y2              =   8160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   113
      X1              =   15645
      X2              =   15645
      Y1              =   8160
      Y2              =   5640
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   112
      X1              =   15840
      X2              =   15840
      Y1              =   5640
      Y2              =   8160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   111
      X1              =   2880
      X2              =   2880
      Y1              =   5640
      Y2              =   8160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   110
      X1              =   2685
      X2              =   2685
      Y1              =   8160
      Y2              =   5640
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   109
      X1              =   2475
      X2              =   2475
      Y1              =   5640
      Y2              =   8160
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   6
      Left            =   4680
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":11942
      Stretch         =   -1  'True
      Top             =   11160
      Width           =   855
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   5
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":12BC4
      Stretch         =   -1  'True
      Top             =   11160
      Width           =   855
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   4
      Left            =   12600
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":13E46
      Stretch         =   -1  'True
      Top             =   11160
      Width           =   855
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   3
      Left            =   12600
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":150C8
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   108
      X1              =   11400
      X2              =   12840
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   107
      X1              =   13200
      X2              =   13200
      Y1              =   11160
      Y2              =   10800
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   106
      X1              =   13005
      X2              =   13005
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   105
      X1              =   12795
      X2              =   12795
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   104
      X1              =   9240
      X2              =   9240
      Y1              =   11160
      Y2              =   10800
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   103
      X1              =   9045
      X2              =   9045
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   1
      Left            =   4680
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":1634A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   102
      X1              =   8835
      X2              =   8835
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   101
      X1              =   4890
      X2              =   4890
      Y1              =   9960
      Y2              =   10200
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   100
      X1              =   5085
      X2              =   5085
      Y1              =   10200
      Y2              =   9765
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   99
      X1              =   5280
      X2              =   5280
      Y1              =   10200
      Y2              =   9840
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   98
      X1              =   13200
      X2              =   13200
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   97
      X1              =   13005
      X2              =   13005
      Y1              =   8280
      Y2              =   8520
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   96
      X1              =   12795
      X2              =   12795
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   95
      X1              =   9045
      X2              =   9045
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   94
      X1              =   9240
      X2              =   9240
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   93
      X1              =   8835
      X2              =   8835
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   92
      X1              =   5280
      X2              =   5280
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   91
      X1              =   5085
      X2              =   5085
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   90
      X1              =   4890
      X2              =   4890
      Y1              =   8520
      Y2              =   8280
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   89
      X1              =   15435
      X2              =   15435
      Y1              =   4320
      Y2              =   5040
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   88
      X1              =   15645
      X2              =   15645
      Y1              =   5040
      Y2              =   4125
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   87
      X1              =   15840
      X2              =   15840
      Y1              =   3915
      Y2              =   5040
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   86
      X1              =   15435
      X2              =   11400
      Y1              =   9540
      Y2              =   9540
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   85
      X1              =   10800
      X2              =   7440
      Y1              =   9540
      Y2              =   9540
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   84
      X1              =   11400
      X2              =   15645
      Y1              =   9765
      Y2              =   9765
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   83
      X1              =   7440
      X2              =   10800
      Y1              =   9765
      Y2              =   9765
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   82
      X1              =   11400
      X2              =   12840
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   81
      X1              =   7440
      X2              =   8880
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   80
      X1              =   5280
      X2              =   5280
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   79
      X1              =   5085
      X2              =   5085
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   78
      X1              =   4890
      X2              =   4890
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   77
      X1              =   9240
      X2              =   9240
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   76
      X1              =   9045
      X2              =   9045
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   75
      X1              =   8835
      X2              =   8835
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   74
      X1              =   13200
      X2              =   13200
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   73
      X1              =   13005
      X2              =   13005
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   72
      X1              =   12810
      X2              =   12810
      Y1              =   5520
      Y2              =   5160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   71
      X1              =   13200
      X2              =   13200
      Y1              =   2835
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   70
      X1              =   13005
      X2              =   13005
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   54
      X1              =   12810
      X2              =   12810
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   53
      X1              =   9240
      X2              =   9240
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   51
      X1              =   9045
      X2              =   9045
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   49
      X1              =   8835
      X2              =   8835
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   47
      X1              =   11400
      X2              =   15640
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   46
      X1              =   7440
      X2              =   10800
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   31
      X1              =   7440
      X2              =   8880
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   26
      X1              =   15435
      X2              =   11400
      Y1              =   4305
      Y2              =   4305
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   8
      X1              =   10800
      X2              =   7440
      Y1              =   4305
      Y2              =   4305
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   615
      Index           =   7
      Left            =   4680
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":175CC
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   69
      X1              =   13200
      X2              =   13200
      Y1              =   4560
      Y2              =   4320
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   68
      X1              =   13005
      X2              =   13005
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   67
      X1              =   12810
      X2              =   12810
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   66
      X1              =   9240
      X2              =   9240
      Y1              =   4560
      Y2              =   4320
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   65
      X1              =   9045
      X2              =   9045
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   64
      X1              =   8835
      X2              =   8835
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   63
      X1              =   13320
      X2              =   15840
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   62
      X1              =   13200
      X2              =   13200
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   61
      X1              =   13005
      X2              =   13005
      Y1              =   4125
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   60
      X1              =   9360
      X2              =   10800
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   59
      X1              =   12810
      X2              =   12810
      Y1              =   3915
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   57
      X1              =   9240
      X2              =   9240
      Y1              =   3960
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   58
      X1              =   8835
      X2              =   8835
      Y1              =   3915
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   56
      X1              =   9045
      X2              =   9045
      Y1              =   4125
      Y2              =   3480
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   55
      X1              =   5400
      X2              =   6840
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   52
      X1              =   5085
      X2              =   5085
      Y1              =   9360
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   50
      X1              =   4890
      X2              =   4890
      Y1              =   9360
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   48
      X1              =   4890
      X2              =   4890
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   45
      X1              =   5085
      X2              =   5085
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   44
      X1              =   5280
      X2              =   5280
      Y1              =   4560
      Y2              =   4320
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   42
      X1              =   5280
      X2              =   5280
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   40
      X1              =   5085
      X2              =   5085
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   38
      X1              =   4890
      X2              =   4890
      Y1              =   2880
      Y2              =   2400
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EDC K2"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   11
      Left            =   12120
      TabIndex        =   18
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EDC K1"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   10
      Left            =   8160
      TabIndex        =   17
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EDC G"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   9
      Left            =   4320
      TabIndex        =   16
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EDC E"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   8
      Left            =   12120
      TabIndex        =   15
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EDC N2"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   7
      Left            =   8160
      TabIndex        =   14
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   37
      X1              =   2475
      X2              =   2475
      Y1              =   3915
      Y2              =   5040
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   36
      X1              =   2685
      X2              =   2685
      Y1              =   5040
      Y2              =   4125
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   35
      X1              =   2880
      X2              =   2880
      Y1              =   4305
      Y2              =   5040
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alternator M2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   6
      Left            =   11400
      TabIndex        =   13
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shore Supply Aft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   5
      Left            =   7320
      TabIndex        =   12
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alternator M1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   3720
      TabIndex        =   11
      Top             =   840
      Width           =   2895
   End
   Begin VB.Shape SRect 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1215
      Index           =   5
      Left            =   12000
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Shape SRect 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1215
      Index           =   4
      Left            =   12000
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   34
      X1              =   2685
      X2              =   2685
      Y1              =   9765
      Y2              =   8760
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   32
      X1              =   2880
      X2              =   2865
      Y1              =   8760
      Y2              =   9540
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EDC N"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Shape SRect 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1215
      Index           =   3
      Left            =   8040
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Shape SRect 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1215
      Index           =   2
      Left            =   4200
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Shape SRect 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1215
      Index           =   1
      Left            =   8040
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      Height          =   855
      Index           =   2
      Left            =   8640
      MousePointer    =   1  'Arrow
      Picture         =   "T42BUS.frx":17E4E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   30
      X1              =   9045
      X2              =   9045
      Y1              =   9360
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   29
      X1              =   9240
      X2              =   9240
      Y1              =   9540
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   28
      X1              =   8835
      X2              =   8835
      Y1              =   9360
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   17
      X1              =   13200
      X2              =   13200
      Y1              =   9600
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   16
      X1              =   2685
      X2              =   6840
      Y1              =   9765
      Y2              =   9765
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   15
      X1              =   2480
      X2              =   4990
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   27
      X1              =   15435
      X2              =   15435
      Y1              =   9540
      Y2              =   8760
   End
   Begin VB.Shape bg_indicator 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   240
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alternator J1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Top             =   12240
      Width           =   3495
   End
   Begin VB.Shape SRect 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1215
      Index           =   0
      Left            =   4200
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   12
      X1              =   13005
      X2              =   13005
      Y1              =   9765
      Y2              =   10200
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   12795
      X2              =   12795
      Y1              =   9360
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   25
      X1              =   2475
      X2              =   2475
      Y1              =   8760
      Y2              =   9960
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alternator J2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   11160
      TabIndex        =   3
      Top             =   12360
      Width           =   3615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<<<< DO NOT DELETE THIS LINE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   23
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OFF LINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   21
      Left            =   1770
      TabIndex        =   1
      Top             =   10680
      Width           =   1905
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shore Supply Ford"
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   3
      Left            =   7440
      TabIndex        =   0
      Top             =   12360
      Width           =   3285
   End
   Begin VB.Shape power_indicator 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   360
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   24
      X1              =   15645
      X2              =   15645
      Y1              =   9765
      Y2              =   8760
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   7
      X1              =   12795
      X2              =   12795
      Y1              =   9960
      Y2              =   10320
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   5
      X1              =   9045
      X2              =   9045
      Y1              =   9765
      Y2              =   10200
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   10
      X1              =   15840
      X2              =   15840
      Y1              =   9960
      Y2              =   8760
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   22
      X1              =   5280
      X2              =   5280
      Y1              =   11160
      Y2              =   10800
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   9
      X1              =   9240
      X2              =   9240
      Y1              =   10200
      Y2              =   9840
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   14
      X1              =   13005
      X2              =   13005
      Y1              =   9120
      Y2              =   9360
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   23
      X1              =   5280
      X2              =   5280
      Y1              =   9600
      Y2              =   9120
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   10800
      X2              =   9360
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Image SImage 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1230
      Index           =   0
      Left            =   0
      Picture         =   "T42BUS.frx":190D0
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   13320
      X2              =   15840
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   4
      X1              =   9240
      X2              =   9240
      Y1              =   9540
      Y2              =   9600
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   19
      X1              =   4890
      X2              =   4890
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   11
      X1              =   13200
      X2              =   13200
      Y1              =   10200
      Y2              =   9840
   End
   Begin VB.Line SLine 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   20
      X1              =   5085
      X2              =   5085
      Y1              =   11160
      Y2              =   10800
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   21
      X1              =   5400
      X2              =   6840
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Line SLine 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   13
      X1              =   6840
      X2              =   2880
      Y1              =   9540
      Y2              =   9540
   End
   Begin VB.Line SLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   8835
      X2              =   8835
      Y1              =   9960
      Y2              =   10200
   End
End
Attribute VB_Name = "frm_power"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' IMO 30/7/96 Form Resize Declarations
Dim orig_form_width As Integer ' Original width of the form
Dim orig_form_height As Integer ' Original height of the form
Dim width_to_height_ratio As Single ' Width/Height ration of form

Private Type sizes
    Top As Integer
    Left As Integer
    Height As Integer
    Width As Integer
End Type

' 28/1/97
'Const max_SLine = 27 ' SLine 0..27
'Const max_SLine = 30 ' SLine 0..30
Const max_SLine = 115 ' SLine 0..115
' 29/1/97
Const max_RLine = 2 ' 0 to 2
Const max_BLine = 2 ' 0 to 2
Const max_YLine = 2 ' 0 to 2

'Const max_SImage = 17 ' SImage 0..17
Const max_SImage = 27 ' SImage 0..27

Const max_bg_indicator = 1 ' bg_indicator 0..1
Const max_power_indicator = 1 ' power_indicator 0..1
Const max_SScroll = 2 ' SScroll 0..2
Const max_ind_motor_load = 0 ' ind_motor_load 0
Const max_label_val = 2 'label_val 0..2
Const max_Label = 26 ' Label 0..26

' 27/1/97
'Const max_SRect = 2 ' SRect 0..2
Const max_SRect = 5 ' SRect 0..5


Dim orig_label(max_Label) As sizes
Dim orig_label_val(max_label_val) As sizes
Dim orig_ind_motor_load(max_ind_motor_load) As sizes
Dim orig_SScroll(max_SScroll) As sizes
Dim orig_power_indicator(max_power_indicator) As sizes
Dim orig_bg_indicator(max_power_indicator) As sizes
Dim orig_SImage(max_SImage) As sizes
Dim orig_SRect(max_SRect) As sizes

Dim orig_SLine(max_SLine) As sizes ' End of resize declarations
' 29/1/97
Dim orig_RLine(max_RLine) As sizes ' End of resize declarations
Dim orig_BLine(max_BLine) As sizes ' End of resize declarations
Dim orig_YLine(max_YLine) As sizes ' End of resize declarations

Dim form_loading As Boolean ' Used to stop resizing of controls during the loading of the form






Sub BRN_lines13(L1 As Integer, L2 As Integer, L3 As Integer, L4 As Integer, L5 As Integer, L6 As Integer, L7 As Integer, L8 As Integer, L9 As Integer, L10 As Integer, L11 As Integer, L12 As Integer, L13 As Integer)

              SLine(L1).BorderColor = &H808080    ' DULL BROWN
              SLine(L2).BorderColor = &H808080    ' DULL BROWN
              SLine(L3).BorderColor = &H808080    ' DULL BROWN
              SLine(L4).BorderColor = &H808080    ' DULL BROWN
              SLine(L5).BorderColor = &H808080    ' DULL BROWN
              SLine(L6).BorderColor = &H808080    ' DULL BROWN
              SLine(L7).BorderColor = &H808080    ' DULL BROWN
              SLine(L8).BorderColor = &H808080    ' DULL BROWN
              SLine(L9).BorderColor = &H808080    ' DULL BROWN
              SLine(L10).BorderColor = &H808080    ' DULL BROWN
              SLine(L11).BorderColor = &H808080    ' DULL BROWN
              SLine(L12).BorderColor = &H808080    ' DULL BROWN
              SLine(L13).BorderColor = &H808080    ' DULL BROWN

End Sub


Sub BRN_lines3(L1 As Integer, L2 As Integer, L3 As Integer)


              SLine(L1).BorderColor = &H808080    ' DULL BROWN
              SLine(L2).BorderColor = &H808080    ' DULL BROWN
              SLine(L3).BorderColor = &H808080    ' DULL BROWN
         
End Sub


















Sub new_resize_this__form()
Const min_form_size = 4500 ' Minumun size the form can have

'Debug.Print "NEW Resize this form"



If Me.WindowState = 0 Then ' normal
'If Sounds_On Then WaveMixPump

'   If Me.Width < min_form_size Then Me.Width = min_form_size
    If Mywidth < min_form_size Then
            Me.Width = min_form_size
    Else
            Me.Width = Mywidth
    End If
   
 ' 29/1/97
        Mywidth = Me.Width
   
'If Sounds_On Then WaveMixPump

'   If Me.Height <> Int(width_to_height_ratio * Me.Width) Then
    If Myheight <> Int(width_to_height_ratio * Me.Width) Then
    
'If Sounds_On Then WaveMixPump
        Me.Height = Int(width_to_height_ratio * Me.Width) ' Keep original proportion of form

' 29/1/97
           Myheight = Me.Height
       Myscalewidth = Me.ScaleWidth
      Myscaleheight = Me.ScaleHeight

'If Sounds_On Then WaveMixPump
        If Not form_loading Then ' Only resize the controls if it is safe to do so, i.e. the width/height is all setup on the form
'If Sounds_On Then WaveMixPump

            Me.Visible = False
            
'If Sounds_On Then WaveMixPump


                Resize_This_Form
 
            
'If Sounds_On Then WaveMixPump
            Me.Visible = True
            
'If Sounds_On Then WaveMixPump
        End If
'If Sounds_On Then WaveMixPump
    End If
    
'If Sounds_On Then WaveMixPump

End If



End Sub

Sub RBY_lines13(L1 As Integer, L2 As Integer, L3 As Integer, L4 As Integer, L5 As Integer, L6 As Integer, L7 As Integer, L8 As Integer, L9 As Integer, L10 As Integer, L11 As Integer, L12 As Integer, L13 As Integer)

              SLine(L1).BorderColor = &HFF          ' RED
              SLine(L2).BorderColor = &HFF          ' RED
              SLine(L3).BorderColor = &HFF          ' RED
              SLine(L4).BorderColor = &HFF          ' RED
              SLine(L5).BorderColor = &HFF          ' RED
              
              SLine(L6).BorderColor = &HFF0000      ' BLUE
              SLine(L7).BorderColor = &HFF0000      ' BLUE
              SLine(L8).BorderColor = &HFF0000      ' BLUE
              SLine(L9).BorderColor = &HFF0000      ' BLUE
              
              SLine(L10).BorderColor = &HFFFF&       ' yellow
              SLine(L11).BorderColor = &HFFFF&       ' yellow
              SLine(L12).BorderColor = &HFFFF&       ' yellow
              SLine(L13).BorderColor = &HFFFF&       ' yellow
              


End Sub

Sub RBY_lines3(L1 As Integer, L2 As Integer, L3 As Integer)

              SLine(L1).BorderColor = &HFF          ' RED
              SLine(L2).BorderColor = &HFF0000      ' BLUE
              SLine(L3).BorderColor = &HFFFF&       ' yellow

End Sub

Sub Resize_Initialise()
Dim obj As sizes
Dim valid As Boolean
Dim req_twips_x As Integer
Dim req_twips_y As Integer
Dim i As Object


' 28/1/97 TEST
'GoTo skip_init:
' Caused a STACK over-flow error in form resize

form_loading = True ' Not yet safe to resize the controls

' 27/1/97
'width_to_height_ratio = 6450 / 4590 ' Design time height/width
'Me.Width = 4590 ' The height is set proportionaly to width in Form_Resize

'width_to_height_ratio = 12375 / 15600 ' Design time height/width
'Me.Width = 15600 ' The height is set proportionaly to width in Form_Resize

'width_to_height_ratio = 9945 / 13395 ' Design time height/width
'Me.Width = 13395 ' The height is set proportionaly to width in Form_Resize

'width_to_height_ratio = 11910 / 15150 ' Design time height/width
'Me.Width = 11910 ' The height is set proportionaly to width in Form_Resize

width_to_height_ratio = 13665 / 18135 ' Design time height/width
Me.Width = 18135 ' The height is set proportionaly to width in Form_Resize

'width_to_height_ratio = Me.Height / Me.Width ' Design time height/width

orig_form_width = Me.ScaleWidth
orig_form_height = Me.ScaleHeight

' 29/1/97
    Myscalewidth = Me.ScaleWidth
    Myscaleheight = Me.ScaleHeight

For Each i In Me
    Select Case i.Name
        Case "SImage"
            valid = True
        Case "SLine"
            valid = True
        Case "RLine"
            valid = True
        Case "BLine"
            valid = True
        Case "YLine"
            valid = True
        Case "Label"
            valid = True
        Case "bg_indicator"
            valid = True
        Case "SScroll"
            valid = True
        Case "label_val"
            valid = True
        Case "ind_motor_load"
            valid = True
        Case "SRect"
            valid = True
        Case "power_indicator"
            valid = True
        Case Else
            valid = False
    End Select
    
    If valid Then
        If i.Name = "SLine" Then
            With obj
                    .Top = i.Y1
                    .Left = i.X1
                    .Height = i.Y2
                    .Width = i.X2
            End With
        ElseIf i.Name = "RLine" Then
            With obj
                    .Top = i.Y1
                    .Left = i.X1
                    .Height = i.Y2
                    .Width = i.X2
            End With
        ElseIf i.Name = "BLine" Then
            With obj
                    .Top = i.Y1
                    .Left = i.X1
                    .Height = i.Y2
                    .Width = i.X2
            End With
        ElseIf i.Name = "YLine" Then
            With obj
                    .Top = i.Y1
                    .Left = i.X1
                    .Height = i.Y2
                    .Width = i.X2
            End With
        Else
            With obj
                .Height = i.Height
                .Width = i.Width
                .Top = i.Top
                .Left = i.Left
            End With
        End If

        Select Case i.Name
            Case "SImage"
                orig_SImage(i.Index) = obj
            Case "SLine"
                orig_SLine(i.Index) = obj
' 29/1/97
            Case "RLine"
                orig_RLine(i.Index) = obj
            Case "BLine"
                orig_BLine(i.Index) = obj
            Case "YLine"
                orig_YLine(i.Index) = obj
                
            Case "Label"
                orig_label(i.Index) = obj
            Case "bg_indicator"
                orig_bg_indicator(i.Index) = obj
            Case "SScroll"
                orig_SScroll(i.Index) = obj
            Case "label_val"
                orig_label_val(i.Index) = obj
            Case "ind_motor_load"
                orig_ind_motor_load(i.Index) = obj
            Case "power_indicator"
                orig_power_indicator(i.Index) = obj
            Case "SRect"
                orig_SRect(i.Index) = obj
        End Select
    End If
Next i

' Now size the form to the correct size for this resolution

' 27/1/97
'req_twips_x = (Screen.Width / Screen.TwipsPerPixelX) * (Me.Width / 800)
'req_twips_x = (Screen.Width / Screen.TwipsPerPixelX) * (Me.Width / 1024)
req_twips_x = (Screen.Width / Screen.TwipsPerPixelX) * (Me.Width / 1280)

form_loading = False ' Okay to resize the controls from now on
Me.Width = req_twips_x ' Again the height is set proportionaly to the width in Form_Resize

' Place this form in the bottom left corner
'Me.Left = 0
'Me.Top = mdi_main.ScaleHeight - Me.Height - 900

skip_init:

End Sub

Sub Resize_This_Form()
Dim i As Object
Dim obj As sizes
Dim scalex!, scaley!
Dim valid As Boolean
Dim item_name As String
Dim calc_temp As Integer


'Debug.Print "Resize this form"

' 28/1/97 TEST RWJ
 'GoTo skip_resize:
 
 ' 28/1/97
 'The next line does not work because the mouse is NOT DOWN when you re-size a form
 ' So remove the next line:
 'If mouse_down = True Then GoTo skip_resize:
 
 
 
 
 
' Note this form STILL flickered between this form and frm_host when the screen was re-sized.
' (Obviously it wouldn't actually re-size with the above goto statement , but the point is there was STILL
' screen flicker when the mouse was used to re-size the form.
' There-fore the flicker has NOTHING to do with the code in this subroutine.

' Note - WaveMixPump statements keep the sound buffer full of data, so during
' long operations ie. resizing, it can keep playing sound and not stop so much
' 5/9/96 IMO

'If Sounds_On Then WaveMixPump


        scalex = Me.ScaleWidth / orig_form_width
        scaley = Me.ScaleHeight / orig_form_height
' 29/1/97
'        scalex = Myscalewidth / orig_form_width
'        scaley = Myscaleheight / orig_form_height
       
        For Each i In Me
'If Sounds_On Then WaveMixPump
            Select Case i.Name
            
                Case "SImage"
                    obj = orig_SImage(i.Index)
                    valid = True
                Case "SLine"
                    obj = orig_SLine(i.Index)
                    valid = True
                    
                Case "RLine"
                    obj = orig_RLine(i.Index)
                    valid = True
                 Case "BLine"
                    obj = orig_BLine(i.Index)
                    valid = True
                 Case "YLine"
                    obj = orig_YLine(i.Index)
                    valid = True
                    
               Case "Label"
                    obj = orig_label(i.Index)
                    valid = True
                Case "bg_indicator"
                    obj = orig_bg_indicator(i.Index)
                    valid = True
                Case "Label"
                    obj = orig_label(i.Index)
                    valid = True
                Case "SScroll"
                    obj = orig_SScroll(i.Index)
                valid = True
                Case "label_val"
                    obj = orig_label_val(i.Index)
                    valid = True
                Case "ind_motor_load"
                    obj = orig_ind_motor_load(i.Index)
                    valid = True
                Case "SRect"
                    obj = orig_SRect(i.Index)
                    valid = True
                Case "power_indicator"
                    obj = orig_power_indicator(i.Index)
                    valid = True
                Case Else
                    valid = False
                End Select
'If Sounds_On Then WaveMixPump
                If valid Then
'If Sounds_On Then WaveMixPump
                    If i.Name = "SLine" Then ' Lines use x1...rather than left!
                        With obj
                            If i.Y1 <> Int(scaley * .Top) Then i.Y1 = scaley * .Top
                            If i.Y1 <> Int(scalex * .Left) Then i.X1 = scalex * .Left
                            If i.X2 <> Int(scalex * .Width) Then i.X2 = scalex * .Width
                            If i.Y2 <> Int(scaley * .Height) Then i.Y2 = scaley * .Height
                        End With
                    ElseIf i.Name = "RLine" Then ' Lines use x1...rather than left!
                        With obj
                            If i.Y1 <> Int(scaley * .Top) Then i.Y1 = scaley * .Top
                            If i.Y1 <> Int(scalex * .Left) Then i.X1 = scalex * .Left
                            If i.X2 <> Int(scalex * .Width) Then i.X2 = scalex * .Width
                            If i.Y2 <> Int(scaley * .Height) Then i.Y2 = scaley * .Height
                        End With
                    ElseIf i.Name = "BLine" Then ' Lines use x1...rather than left!
                        With obj
                            If i.Y1 <> Int(scaley * .Top) Then i.Y1 = scaley * .Top
                            If i.Y1 <> Int(scalex * .Left) Then i.X1 = scalex * .Left
                            If i.X2 <> Int(scalex * .Width) Then i.X2 = scalex * .Width
                            If i.Y2 <> Int(scaley * .Height) Then i.Y2 = scaley * .Height
                        End With
                    ElseIf i.Name = "YLine" Then ' Lines use x1...rather than left!
                        With obj
                            If i.Y1 <> Int(scaley * .Top) Then i.Y1 = scaley * .Top
                            If i.Y1 <> Int(scalex * .Left) Then i.X1 = scalex * .Left
                            If i.X2 <> Int(scalex * .Width) Then i.X2 = scalex * .Width
                            If i.Y2 <> Int(scaley * .Height) Then i.Y2 = scaley * .Height
                        End With
                        
                    Else
                        With obj
                            If i.Top <> Int(scaley * .Top) Then i.Top = scaley * .Top
                            If i.Left <> Int(scalex * .Left) Then i.Left = scalex * .Left
'If Sounds_On Then WaveMixPump
                            If (Left$(i.Name, 5) = "Label") Then
                                ' if it is a line, rather than making the box bigger, take account of the
                                ' extra size by centering the box in the extra space
                                
' 28/1/97 RWJ comment out the next 4 lines keeps the correct (and re-sized) x,y, cords for labels...
'                                calc_temp = i.Top + (((scaley * .Height) - .Height) / 2)
'                                If i.Top <> calc_temp Then i.Top = calc_temp
'                                calc_temp = i.Left + (((scalex * .Width) - .Width) / 2)
'                                If i.Left <> calc_temp Then i.Left = calc_temp

                            End If
                            
'If Sounds_On Then WaveMixPump

' 28/1/97 TRY COMMENTING OUT (for label text re-sizing)
' Did not solve the problem of incorrect positioning of labels so put back in:
                            If (Left$(i.Name, 14) <> "ind_motor_load") Then ' Don't change a combo_box width/height
                                If (Left$(i.Name, 9) <> "label_val") Then ' nor resize these labels
                                    i.Height = scaley * .Height
                                    i.Width = scalex * .Width
                                End If
                            End If
                            
'If Sounds_On Then WaveMixPump
                        End With
'If Sounds_On Then WaveMixPump
                    End If
'If Sounds_On Then WaveMixPump
            End If
'If Sounds_On Then WaveMixPump
        Next i
        
skip_resize:
End Sub

Sub scan_cross_bkr()
Dim i As Integer, ilast As Integer


For i = 1 To 8

         If i = 1 Then
            ilast = 8
         Else
            ilast = i - 1
         End If
                
' NB the next line is very important:
          Call scan_Xlive
          
          If Not Xlive(ilast) And Not Xlive(i) Then
              Cross_BKR_closed(ilast) = False
          End If
       
Next i


        
End Sub

Sub scan_Xlive()
Dim j As Integer

If Not ON_LINE Then

   ' Scan forward to determine status of bus-bars (Xlive(1 to 8)
   For j = 1 To 7
     If Xlive(j) And Cross_BKR_closed(j) Then Xlive(j + 1) = True
   Next j
     If Xlive(8) And Cross_BKR_closed(8) Then Xlive(1) = True
     
   ' Scan backwards to determine status of bus-bars  (Xlive(1 to 8)
   For j = 8 To 2 Step -1
     If Xlive(j) And Cross_BKR_closed(j - 1) Then Xlive(j - 1) = True
   Next j
     If Xlive(1) And Cross_BKR_closed(8) Then Xlive(8) = True
     
End If

End Sub

Private Sub Form_Activate()
Dim ret As Integer

' if disabled, enable the form as it is now visible
If frm_power.Enabled = False Then frm_power.Enabled = True



End Sub

Private Sub Form_Load()
Const min_form_size = 4500 ' Minumun size the form can have
Dim i As Integer

' 30/1/97
For i = 0 To max_SLine
              SLine(i).BorderColor = &H808080    ' DULL BROWN
Next i


' 28/1/97
resize_flag = False


'Debug.Print "Resize init"

Resize_Initialise

'Debug.Print "Resize init end"


'=== 29/1/97 ===================
If Me.WindowState = 0 Then ' normal

    If Me.Height <> Int(width_to_height_ratio * Me.Width) Then
        Me.Height = Int(width_to_height_ratio * Me.Width) ' Keep original proportion of form
     End If
     Mywidth = Me.Width
     Myheight = Me.Height
End If
'======================

' 29/1/97
   resize_flag = True
   power_system_timer.Enabled = True
   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


' 28/1/97
'mouse_down = True
'Label(21).Caption = "Down"
Debug.Print "frm_power mouse_down = true "

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 28/1/97
'mouse_down = False
'Label(21).Caption = "up"
Debug.Print "frm_power mouse_down = false "



End Sub

Private Sub Form_Paint()
' Note - WaveMixPump statements keep the sound buffer full of data, so during
' long operations ie. resizing, it can keep playing sound and not stop so much
' 5/9/96 IMO

'24/1/97
'If Sounds_On Then WaveMixPump

End Sub

Private Sub Form_Resize()
Const min_form_size = 4500 ' Minumun size the form can have


'===== 29/1/97=====
     resize_flag = True
     Mywidth = Me.Width
     Myheight = Me.Height
    Myscalewidth = Me.ScaleWidth
    Myscaleheight = Me.ScaleHeight
    


'   Debug.Print "frm_power resize "

'   GoTo end_resize:
     
'==================


' Note - WaveMixPump statements keep the sound buffer full of data, so during
' long operations ie. resizing, it can keep playing sound and not stop so much
' 5/9/96 IMO


' 24/1/97
'If Sounds_On Then WaveMixPump
If Me.WindowState = 0 Then ' normal


   If Me.Width < min_form_size Then Me.Width = min_form_size
    
   If Me.Height <> Int(width_to_height_ratio * Me.Width) Then
    
        Me.Height = Int(width_to_height_ratio * Me.Width) ' Keep original proportion of form


         
        If Not form_loading Then ' Only resize the controls if it is safe to do so, i.e. the width/height is all setup on the form


            Me.Visible = False
            
'If Sounds_On Then WaveMixPump


               Resize_This_Form
 
            
'If Sounds_On Then WaveMixPump
            Me.Visible = True
            
'If Sounds_On Then WaveMixPump
        End If

'If Sounds_On Then WaveMixPump

    End If
    
'If Sounds_On Then WaveMixPump


'If Sounds_On Then WaveMixPump

End If


end_resize:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
' 27/1/97
'    mdi_main.mnuPower.Checked = False
'    form_open(4) = False

End Sub





Private Sub Label_Click(Index As Integer)
Dim i As Integer

 If Index = 21 Then
 
   ON_LINE = Not ON_LINE
   
   If ON_LINE Then
   
      Label(21).Caption = " ON LINE "
      Label(21).ForeColor = &HFF00&
      
' 3/2/97 Save the OFF-LINE STATUS
      For i = 1 To 6
         old_gen_ON(i) = gen_ON(i)
         old_Main_BKR_closed(i) = Main_BKR_closed(i)
         old_EDC_BKR_closed(i) = EDC_BKR_closed(i)
      Next i
      For i = 1 To 8
         old_Xlive(i) = Xlive(i)
         old_Cross_BKR_closed(i) = Cross_BKR_closed(i)
      Next i
      
    Else
    
      Label(21).Caption = " OFF LINE "
      Label(21).ForeColor = &HFF&
      
' 3/2/97 RESTORE the OFF-LINE STATUS
      For i = 1 To 6
         gen_ON(i) = old_gen_ON(i)
         Main_BKR_closed(i) = old_Main_BKR_closed(i)
         EDC_BKR_closed(i) = old_EDC_BKR_closed(i)
      Next i
      For i = 1 To 8
         Xlive(i) = old_Xlive(i)
         Cross_BKR_closed(i) = old_Cross_BKR_closed(i)
      Next i
    End If
    
 End If
 

End Sub

Private Sub SImage_Click(Index As Integer)
Dim Ngen As Integer, Nbkr As Integer

' ====  Alternators =============================================

If Index < 7 Then

        If Not ON_LINE Then
          gen_ON(Index) = Not gen_ON(Index)
          Debug.Print "not ON_line gen_ON =  " & gen_ON(Index)
        End If
End If
     
     
' ======= Main breakers =========================================

If Index > 6 And Index < 13 Then

            Ngen = Index - 6
            
       If Not ON_LINE Then
             If gen_ON(Ngen) Then
                 Main_BKR_closed(Ngen) = Not Main_BKR_closed(Ngen)
             Else
                 Main_BKR_closed(Ngen) = False
             End If
       End If
             
End If

' ====== Cross-over breakers ====================================
If Index > 12 And Index < 21 Then

            Nbkr = Index - 12
            
       If Not ON_LINE Then
                 Cross_BKR_closed(Nbkr) = Not Cross_BKR_closed(Nbkr)
       End If
        
End If

' ======= EDC load breakers ====================================
If Index > 20 And Index < 27 Then

            Nbkr = Index - 20
            
       If Not ON_LINE Then
                 EDC_BKR_closed(Nbkr) = Not EDC_BKR_closed(Nbkr)
       End If
       
End If

 

End Sub
















Private Sub power_system_timer_Timer()
Dim i As Integer
Dim ilast As Integer, istart As Integer

' 28/1/97
If resize_flag Then
'           Me.Visible = False
'        Debug.Print " New resize "
'
'         new_resize_this__form

            resize_flag = False
'           Me.Visible = True
End If


If ON_LINE Then
     For i = 1 To 8
     
' 3/2/97 NB Do NOT Rx T42BKRL( ) directly
' VB agrees with FORTRAN that 1 = true
' but normal definition of a BYTE variable is 255 = true and 0 = false
' So receive from FORTRAN changed....
' 3/2/97      Xlive(i) = T42BKRL(24 + i)   'get from model DLL
          If T42BKRL(i) Then               '  FORTRAN BUSLIVE(1 to 16)
              Xlive(i) = True
          Else
              Xlive(i) = False
          End If
     Next i
Else
      For i = 1 To 8
           Xlive(i) = False
      Next i
End If

' ====  Alternators running =============================================

For i = 1 To 6

       If ON_LINE Then
       
'          Debug.Print "ON_line gen_ON "

' 3/2/97       gen_ON(i) = T42BKRL(16 + i)      '(N.B. get from Model DLL)
            If T42BKRL(16 + i) Then
                gen_ON(i) = True
            Else
                gen_ON(i) = False
            End If
            
       End If

        If gen_ON(i) Then
        
           If i < 4 Then
            SImage(i).Picture = PicClip2(3).Picture        ' Green GEN ON picture (inverted)
           Else
            SImage(i).Picture = PicClip2(1).Picture        ' Green GEN ON picture (normal)
           End If
           
           If i = 1 Then Call RBY_lines3(38, 40, 42)       ' M1
           If i = 2 Then Call RBY_lines3(49, 51, 53)       ' S/S Aft
           If i = 3 Then Call RBY_lines3(54, 70, 71)       ' M2
           If i = 4 Then Call RBY_lines3(105, 106, 107)    ' J2
           If i = 5 Then Call RBY_lines3(102, 103, 104)    ' S/S Ford
           If i = 6 Then Call RBY_lines3(19, 20, 22)       ' J1
                    
        Else
        
           If i < 4 Then
             SImage(i).Picture = PicClip2(2).Picture       ' Grey GEN OFF picture (inverted)
           Else
             SImage(i).Picture = PicClip2(0).Picture       ' Grey GEN OFF picture (normal)
           End If
        
           If i = 1 Then Call BRN_lines3(38, 40, 42)       ' M1
           If i = 2 Then Call BRN_lines3(49, 51, 53)       ' S/S Aft
           If i = 3 Then Call BRN_lines3(54, 70, 71)       ' M2
           If i = 4 Then Call BRN_lines3(105, 106, 107)    ' J2
           If i = 5 Then Call BRN_lines3(102, 103, 104)    ' S/S Ford
           If i = 6 Then Call BRN_lines3(19, 20, 22)       ' J1
                              
        End If
Next i

' ======= Main breakers ==================================================
For i = 1 To 6

       If ON_LINE Then
' 3/2/97      Main_BKR_closed(i) = T42BUSL(i)      ' (N.B get from model DLL)
              If T42BUSL(i) Then
                 Main_BKR_closed(i) = True
              Else
                 Main_BKR_closed(i) = False
              End If
       Else
          If Not gen_ON(i) Then
              Main_BKR_closed(i) = False
          End If
       End If
       
       If Not Main_BKR_closed(i) Then
          
            SImage(i + 6).Picture = PicClip2(4).Picture    ' Vertical breaker OPEN
                                   
       Else
          
            SImage(i + 6).Picture = PicClip2(5).Picture    ' Vertical breaker CLOSED
                        
           If Not ON_LINE Then
              If i = 1 Then Xlive(1) = True
              If i = 2 Then Xlive(2) = True
              If i = 3 Then Xlive(3) = True
              If i = 4 Then Xlive(5) = True
              If i = 5 Then Xlive(6) = True
              If i = 6 Then Xlive(7) = True
           End If
            
            
       End If
Next i

' ===== Cross-over breakers ==============================================


If ON_LINE Then

    For i = 1 To 8
' 3/2/97     Cross_BKR_closed(i) = T42BUSL(24 + i)  '(N.B get from model DLL)
             If T42BUSL(24 + i) Then
                 Cross_BKR_closed(i) = True
             Else
                 Cross_BKR_closed(i) = False
             End If
    Next i

Else
' Scan the Cross-over Breakers (trip them if the bus-lines go dead)
        Call scan_cross_bkr
        
' Scan the bus-bars to determine the status of Xlive(1 to 8)
        Call scan_Xlive
End If


' Draw the Breaker Pictures
For i = 1 To 8
        If i = 1 Or i = 2 Or i = 5 Or i = 6 Then
                If Not Cross_BKR_closed(i) Then
                       SImage(i + 12).Picture = PicClip2(6).Picture    ' Cross over Horizontal breaker OPEN
                Else
                       SImage(i + 12).Picture = PicClip2(7).Picture    ' Cross over Horizontal breaker CLOSED
                End If
        Else
            If Not Cross_BKR_closed(i) Then
                       SImage(i + 12).Picture = PicClip2(4).Picture    ' Cross over Vertical breaker OPEN
            Else
                       SImage(i + 12).Picture = PicClip2(5).Picture    ' Cross over Vertical breaker CLOSED
            End If
        End If
Next i

              
' Paint the bus-bar lines:
For i = 1 To 8
        If Xlive(i) Then
            If i = 1 Then Call RBY_lines13(39, 18, 37, 48, 55, 41, 33, 36, 45, 43, 3, 35, 44)     ' M1 BKR OUTPUT
            If i = 2 Then Call RBY_lines13(31, 58, 64, 60, 115, 46, 56, 65, 115, 8, 57, 66, 115)  ' S/S Aft BKR OUTPUT
            If i = 3 Then Call RBY_lines13(108, 59, 67, 63, 87, 47, 61, 68, 88, 26, 62, 69, 89)   ' M2 BKR OUTPUT
            If i = 4 Then Call RBY_lines3(112, 113, 114)                                          '  Xlive(4)
            If i = 5 Then Call RBY_lines13(10, 0, 82, 2, 7, 24, 84, 14, 12, 27, 86, 17, 11)       ' J2 BKR OUTPUT
            If i = 6 Then Call RBY_lines13(1, 81, 28, 6, 115, 83, 30, 5, 115, 85, 29, 9, 115)     ' S/S Ford BKR OUTPUT
            If i = 7 Then Call RBY_lines13(21, 15, 25, 50, 101, 16, 34, 52, 100, 13, 32, 23, 99)  ' J1 BKR OUTPUT
            If i = 8 Then Call RBY_lines3(109, 110, 111)                                          '  Xlive(8)
       Else
            If i = 1 Then Call BRN_lines13(39, 18, 37, 48, 55, 41, 33, 36, 45, 43, 3, 35, 44)     ' M1 BKR OUTPUT
            If i = 2 Then Call BRN_lines13(31, 58, 64, 60, 115, 46, 56, 65, 115, 8, 57, 66, 115)  ' S/S Aft BKR OUTPUT
            If i = 3 Then Call BRN_lines13(108, 59, 67, 63, 87, 47, 61, 68, 88, 26, 62, 69, 89)   ' M2 BKR OUTPUT
            If i = 4 Then Call BRN_lines3(112, 113, 114)                                          '  Xlive(4)
            If i = 5 Then Call BRN_lines13(10, 0, 82, 2, 7, 24, 84, 14, 12, 27, 86, 17, 11)       ' J2 BKR OUTPUT
            If i = 6 Then Call BRN_lines13(1, 81, 28, 6, 115, 83, 30, 5, 115, 85, 29, 9, 115)     ' S/S Ford BKR OUTPUT
            If i = 7 Then Call BRN_lines13(21, 15, 25, 50, 101, 16, 34, 52, 100, 13, 32, 23, 99)  ' J1 BKR OUTPUT
            If i = 8 Then Call BRN_lines3(109, 110, 111)                                          '  Xlive(8)
        End If
Next i




' ===== EDC breakers ==============================================

If ON_LINE Then

      For i = 1 To 6
' 3/2/97     EDC_BKR_closed(i) = T42BUSL(12 + i)    '(N.B get from model DLL)
             If T42BUSL(12 + i) Then
                  EDC_BKR_closed(i) = True
             Else
                  EDC_BKR_closed(i) = False
             End If
      Next i

Else
          If Not Xlive(1) Then EDC_BKR_closed(1) = False
          If Not Xlive(2) Then EDC_BKR_closed(2) = False
          If Not Xlive(3) Then EDC_BKR_closed(3) = False
          If Not Xlive(5) Then EDC_BKR_closed(4) = False
          If Not Xlive(6) Then EDC_BKR_closed(5) = False
          If Not Xlive(7) Then EDC_BKR_closed(6) = False
End If


For i = 1 To 6

  If i = 1 Then istart = 78
  If i = 2 Then istart = 75
  If i = 3 Then istart = 72
  If i = 4 Then istart = 96
  If i = 5 Then istart = 93
  If i = 6 Then istart = 90
  
  If EDC_BKR_closed(i) Then
                 SImage(i + 20).Picture = PicClip2(5).Picture    ' Vertical breaker CLOSED
           Call RBY_lines3(istart, istart + 1, istart + 2)       ' EDC bus line colours
  Else
                 SImage(i + 20).Picture = PicClip2(4).Picture    ' Vertical breaker OPEN
           Call BRN_lines3(istart, istart + 1, istart + 2)       ' EDC bus line colours
  End If
             
Next i


  
End Sub










