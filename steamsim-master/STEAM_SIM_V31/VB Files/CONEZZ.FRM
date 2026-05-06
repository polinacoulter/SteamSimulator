VERSION 5.00
Begin VB.Form cone 
   Appearance      =   0  'Flat
   BackColor       =   &H00400040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ENGINE CONTROLS"
   ClientHeight    =   6870
   ClientLeft      =   -1020
   ClientTop       =   1305
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ONpump 
      BackColor       =   &H00C0C0C0&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4200
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   5700
      UseMaskColor    =   -1  'True
      Width           =   1395
   End
   Begin VB.PictureBox Panel3D2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2175
      Left            =   6480
      ScaleHeight     =   2145
      ScaleWidth      =   2505
      TabIndex        =   67
      Top             =   3720
      Width           =   2535
      Begin VB.PictureBox Command3D10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   75
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Command3D9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   74
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Command3D8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   705
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   73
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Command3D7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   705
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   72
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Command3D6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   71
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Command3D5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1890
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   70
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox Command3D4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   700
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   69
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Command3D3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   700
         ScaleHeight     =   225
         ScaleWidth      =   465
         TabIndex        =   68
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER FAN  4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   1320
         TabIndex        =   79
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER FAN  3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   78
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER FAN  2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   1320
         TabIndex        =   77
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ER FAN  1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   205
         Left            =   120
         TabIndex        =   76
         Top             =   900
         Width           =   1095
      End
      Begin VB.Image Image18 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   3
         Left            =   1290
         Picture         =   "conezz.frx":0000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image18 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   2
         Left            =   120
         Picture         =   "conezz.frx":030A
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image18 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   1
         Left            =   1290
         Picture         =   "conezz.frx":0614
         Stretch         =   -1  'True
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image18 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "conezz.frx":091E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox Panel3D1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1695
      Left            =   6480
      ScaleHeight     =   1665
      ScaleWidth      =   1665
      TabIndex        =   66
      Top             =   1920
      Width           =   1695
      Begin VB.TextBox RUDDERANG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   62
         Text            =   " "
         Top             =   390
         Width           =   375
      End
      Begin VB.PictureBox Spin1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   1185
         TabIndex        =   80
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.PictureBox GroupPush3D2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   81
      Top             =   6240
      Width           =   855
   End
   Begin VB.PictureBox GroupPush3D1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   82
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   49
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   48
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   47
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox Command3D2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   46
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Command3D19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   45
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Command3D13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1800
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   44
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   43
      Text            =   "conezz.frx":0C28
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Command3D27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4140
      MousePointer    =   1  'Arrow
      ScaleHeight     =   465
      ScaleWidth      =   1425
      TabIndex        =   0
      Top             =   6060
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   39
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "conezz.frx":0D04
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   36
      Text            =   "conezz.frx":0D16
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "conezz.frx":0D28
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   34
      Text            =   "conezz.frx":0D40
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "conezz.frx":0D51
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "conezz.frx":0D5E
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   31
      Text            =   "conezz.frx":0D6F
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "conezz.frx":0D7F
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "conezz.frx":0D8D
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox enginerpm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4440
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Panel3D12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3975
      Left            =   3360
      ScaleHeight     =   3945
      ScaleWidth      =   2865
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   61
         Top             =   3120
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   60
         Top             =   2880
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   59
         Top             =   2640
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   58
         Top             =   2400
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   57
         Top             =   2160
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   56
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   55
         Top             =   1680
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   54
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   53
         Top             =   1200
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   52
         Top             =   960
         Width           =   615
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   42
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   41
         Top             =   3360
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   40
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " FULL"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   990
         Width           =   525
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " HALF"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1240
         Width           =   540
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " SLOW"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " DEAD SLOW"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1730
         Width           =   1185
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " AHEAD"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STOP"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   1950
         Width           =   510
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " FULL"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2910
         Width           =   525
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " HALF"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   2670
         Width           =   540
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " SLOW"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2430
         Width           =   615
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " DEAD SLOW"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " ASTERN"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   825
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " FULL AWAY"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 10"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 8 "
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2400
         TabIndex        =   8
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 6 "
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 4 "
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 2 "
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 0 "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 2 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 4 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 6 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 8 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " 10 "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   3
         Top             =   3120
         Width           =   345
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   2280
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.PictureBox Panel3D7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   975
      Left            =   3960
      ScaleHeight     =   945
      ScaleWidth      =   1665
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      Begin VB.Shape Shape25 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   360
         Shape           =   3  'Circle
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape Shape26 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7560
      Top             =   240
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMERGENCY RUN"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   65
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMERGENCY STOP"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   64
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING INDICATORS"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   51
      Top             =   2040
      Width           =   2115
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      Height          =   2415
      Left            =   120
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   1215
      Left            =   120
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "AUXILIARY BLOWER"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   50
      Top             =   5640
      Width           =   1845
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "revs/min"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   28
      Top             =   1200
      Width           =   750
   End
End
Attribute VB_Name = "cone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Dim aivopen

Private Sub Command3D1_Click()

Screen.MousePointer = 11

 'temp return
 carsy = 0: If xclick = 1 Then Exit Sub: xclick = 1
 
 

 If pageno$ = "sy010" Then Call xmover: sy010.Show: Unload cone
 If pageno$ = "sy011" Then Call xmover: sy011.Show: Unload cone
 If pageno$ = "sy012" Then Call xmover: sy012.Show: Unload cone
 If pageno$ = "sy013" Then Call xmover: sy013.Show: Unload cone
 If pageno$ = "sy020" Then Call xmover: sy020.Show: Unload cone
 If pageno$ = "sy021" Then Call xmover: sy021.Show: Unload cone
 If pageno$ = "sy022" Then Call xmover: sy022.Show: Unload cone
 If pageno$ = "sy023" Then Call xmover: sy023.Show: Unload cone
 If pageno$ = "sy023H" Then Call xmover: sy023h.Show: Unload cone
 If pageno$ = "sy030" Then Call xmover: sy030.Show: Unload cone
 If pageno$ = "sy030C" Then Call xmover: sy030C.Show: Unload cone
 If pageno$ = "sy035" Then Call xmover: sy035.Show: Unload cone
 If pageno$ = "sy035C" Then Call xmover: sy035C.Show: Unload cone
 
 If pageno$ = "sy041" Then Call xmover: sy041.Show: Unload cone
 If pageno$ = "sy043" Then Call xmover: sy043.Show: Unload cone
 If pageno$ = "sy046" Then Call xmover: sy046.Show: Unload cone
 If pageno$ = "sy046z" Then Call xmover: sy046z.Show: Unload cone
 
 If pageno$ = "sy051" Then Call xmover: sy051.Show: Unload cone
 If pageno$ = "sy053" Then Call xmover: sy053.Show: Unload cone
 If pageno$ = "sy054" Then Call xmover: sy054.Show: Unload cone
 
 If pageno$ = "sy061" Then Call xmover: sy061.Show: Unload cone
 If pageno$ = "sy064" Then Call xmover: sy064.Show: Unload cone
 If pageno$ = "sy064" Then Call xmover: sy064.Show: Unload cone
 
 If pageno$ = "sy071" Then Call xmover: sy071.Show: Unload cone
 If pageno$ = "sy072" Then Call xmover: sy072.Show: Unload cone
 If pageno$ = "sy073" Then Call xmover: sy073.Show: Unload cone

 If pageno$ = "sy080" Then Call xmover: sy080.Show: Unload cone
 If pageno$ = "sy085" Then Call xmover: sy085.Show: Unload cone
 If pageno$ = "sy087" Then Call xmover: sy087.Show: Unload cone
 
 If pageno$ = "sy090" Then Call xmover: sy090.Show: Unload cone
 If pageno$ = "sy090C" Then Call xmover: sy090C.Show: Unload cone
 If pageno$ = "sy091" Then Call xmover: sy091.Show: Unload cone
 If pageno$ = "sy096" Then Call xmover: sy096.Show: Unload cone
 If pageno$ = "sy096c" Then Call xmover: sy096c.Show: Unload cone
 
 If pageno$ = "sy101" Then Call xmover: sy101.Show: Unload cone
 If pageno$ = "sy104" Then Call xmover: sy104.Show: Unload cone

 If pageno$ = "sy121" Then Call xmover: sy121.Show: Unload cone
 If pageno$ = "sy123" Then Call xmover: sy123.Show: Unload cone
 'If pageno$ = "sy124" Then Call xmover: sy124.Show : Unload cone
 If pageno$ = "sy130" Then Call xmover: sy130.Show: Unload cone
 If pageno$ = "sy131" Then Call xmover: sy131.Show: Unload cone
 If pageno$ = "sy137" Then Call xmover: sy137.Show: Unload cone



End Sub

Private Sub Command3D10_Click()
'2133 ER Fan4 start
 j = 2133
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If

End Sub

Private Sub Command3D13_Click()

'137 aux blower standby button
 j = 217
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 Else
   Call warnppfail
 End If


End Sub

Private Sub Command3D19_Click()

'137 aux blower start button
 j = 223
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 Else
   Call warnppfail
 End If


End Sub

Private Sub Command3D2_Click()

'137 aux blower stop button
 j = 208
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 Else
   Call warnppfail
 End If


End Sub

Private Sub Command3D27_Click()

 'air start button
 j = 72
 pcdi.Text = Format$(j, "0"): If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 
End Sub

Private Sub Command3D3_Click()
'2134 ER Fan1 stop
 j = 2134
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If

End Sub

Private Sub Command3D4_Click()

'2126 ER1 Fan start
 j = 2126
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If

End Sub

Private Sub Command3D5_Click()
'2127 ER Fan2 stop
 j = 2127
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If

End Sub

Private Sub Command3D6_Click()
'2135 ER2 Fan start
 j = 2135
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If
End Sub

Private Sub Command3D7_Click()
'2121 ER Fan3 stop
 j = 2121
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If

End Sub

Private Sub Command3D8_Click()
'2129 ER Fan3 start
 j = 2129
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If
End Sub

Private Sub Command3D9_Click()
'2125 ER Fan4 stop
 j = 2125
 If ch(3361) = False Then
   If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 End If

End Sub

Private Sub Form_Load()
Dim temp As Integer
 exit_req = False
 exit_ena = False
 Text36.Text = pageno$

' moveno$ = pageno$
' pageno$ = "cone"

 Timer1.Interval = timercall
 Timer1.Enabled = True




'-----------------
 For j = 0 To 5119
   oldch(j) = -150
 Next

 oldconstat = -150
'-----------------

 
 'set fo lever to previous value
 'oldfolever
 
 If oldfolever = 20 Then
   Picture3(0).BackColor = &HFFFF&
 ElseIf oldfolever = 18 Then
   Picture3(1).BackColor = &HFFFF&
 ElseIf oldfolever = 16 Then
   Picture3(2).BackColor = &HFFFF&
 ElseIf oldfolever = 14 Then
   Picture3(3).BackColor = &HFFFF&
 ElseIf oldfolever = 12 Then
   Picture3(4).BackColor = &HFFFF&
 ElseIf oldfolever = 10 Then
   Picture3(5).BackColor = &HFFFF&
 ElseIf oldfolever = 8 Then
   Picture3(6).BackColor = &HFFFF&
 ElseIf oldfolever = 6 Then
   Picture3(7).BackColor = &HFFFF&
 ElseIf oldfolever = 4 Then
   Picture3(8).BackColor = &HFFFF&
 ElseIf oldfolever = 2 Then
   Picture3(9).BackColor = &HFFFF&
 ElseIf oldfolever = 0 Then
   Picture3(10).BackColor = &HFFFF&
 End If


 'blower on
 If ch(333) = True Then
   Text14.BackColor = &HFF00&
 Else
   Text14.BackColor = &H808080
 End If

 'blower stop
 If ch(369) = True Then
   Text15.BackColor = &HFFFFFF
 Else
   Text15.BackColor = &H808080
 End If

'blower auto
 If ch(381) = True Then
   Text16.BackColor = &HFFFF&
 Else
   Text16.BackColor = &H808080
 End If

 
 
'disable controls if monitoring
 If constat = 1 Then
   Command3D19.Visible = True
   Command3D2.Visible = True
   Command3D13.Visible = True
   Command3D27.Visible = True
 Else
   Command3D19.Visible = False
   Command3D2.Visible = False
   Command3D13.Visible = False
   Command3D27.Visible = False
 End If



'initialize the properties

 If testrun = False Then
 
   On Error GoTo Handler

   pcai00.Text = " "
   Call init_coldlink(pcai00, boxtopic)             'pc remote signal
   pcai16.Text = " "
   Call init_coldlink(pcai16, boxtopic)             'pc remote signal
   
   pcdi.Text = " "
   Call init_coldlink(pcdi, boxtopic)


   'ginit
   
 End If

 GoTo leave

Handler:
    msg1$ = "Server has not responded to a DDE Initiate."
    msg2$ = "Press OK when Server is Ready, or CANCEL to terminate application."
    If (Err = 282) Then
        Select Case MsgBox(msg1$ + Chr$(KEY_RETURN) + msg2$, MB_OKCANCEL)
            Case IDOK
                Resume 0
            Case IDCANCEL
                Resume Terminate_app
            End Select
    Else
   'text24.Text = "True"
        Error Err
        Resume Next
    End If

Terminate_app:
    End

    
' cone.Show
' vscroll1.SetFocus

leave:
End Sub

Private Sub GroupPush3D1_Click(Value As Integer)

'EMERGENCY START BUTTON
 j = 78
 pcdi.Text = Format$(ipcdi, "0000")
 If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))

End Sub

Private Sub GroupPush3D2_Click(Value As Integer)
'EMERGENCY STOP BUTTON
 j = 86
 pcdi.Text = Format$(ipcdi, "0000")
 If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))

End Sub

Private Sub GroupPush3D3_Click(Index As Integer, Value As Integer)

End Sub

Private Sub ONpump_Click(Index As Integer)

 'air start button
 j = 72
 pcdi.Text = Format$(j, "0"): If testrun = False Then Call pokefun(pcdi, "pcdi", Format$(j, "0"))
 
End Sub

Private Sub Picture3_Click(Index As Integer)
 
 aivopen = (10 - Index) * 2
 
 'testbox
 Text10.Text = Format$(aivopen, "0")
 ditxt$ = Format$(aivopen, "0")
 pcai00.Text = ditxt$
 If testrun = False Then Call pokefun(pcai00, "pcai00", ditxt$)
 
 oldfolever = aivopen
 'testbox
 Text11.Text = Format$(oldfolever, "0")
 
' set the color of the box
 For i = 0 To 10
    Picture3(i).BackColor = &HC0C0C0
 Next
 Picture3(Index).BackColor = &HFFFF&
 Text17.Text = Index
End Sub

Private Sub Rudder_Change()

End Sub

Private Sub RUDDERANG_Change()

'rudder angle 114
   If Val(RUDDERANG.Text) > 0 Then               'rudder stbd
     Label16.ForeColor = gaugest
     Label16.Caption = " STBD "
   ElseIf Val(RUDDERANG.Text) < 0 Then           'rudder port
     Label16.ForeColor = gaugepo
     Label16.Caption = " PORT "
   ElseIf Val(RUDDERANG.Text) = 0 Then           'rudder mid
     Label16.ForeColor = &H0&
     Label16.Caption = " MIDSHIPS "
   End If

End Sub

Private Sub Spin1_SpinDown()
If RUDDERANG > -40 Then
    RUDDERANG = RUDDERANG - 1
End If
 If ch(3361) = False And testrun = False Then
   'set and poke rudder value
   j = RUDDERANG
   Call pokefun(pcai16, "pcai16", Format$(j, "0"))
 Else
   Call warnppfail
 End If

End Sub

Private Sub Spin1_SpinUp()
If RUDDERANG < 40 Then
    RUDDERANG = RUDDERANG + 1
End If
 If ch(3361) = False And testrun = False Then
   'set and poke rudder value
   j = RUDDERANG
   Call pokefun(pcai16, "pcai16", Format$(j, "0"))
 Else
   Call warnppfail
 End If

End Sub

Private Sub Timer1_Timer()

    Call get_pcints
' use pump_status(0) to detect the Auxy blower...
    Call pp_status(0, 48)

' 17/5/95.... Status lamps.....
   Call get_ch(376, 133)
   Call get_ch(377, 137)
   Call get_ch(378, 181)
   Call get_ch(379, 191)
   Call get_ch(380, 129)
   Call get_ch(381, 183)
   Call get_ch(382, 187)
   Call get_ch(383, 185)
   Call get_ch(384, 179)
'
' Cam ahead lamp...
  Call get_ch(385, 15)

' Emergency Run code...
  Call get_ch(375, 177)
  If ch(177) <> oldch(177) Then
    oldch(177) = ch(177)
    If ch(177) Then
        GroupPush3D1.BackColor = &HFF&
     Else
        GroupPush3D1.BackColor = &H80&
     End If
   End If

' Emergency Stop code...
  Call get_ch(304, 189)
  If ch(189) <> oldch(189) Then
    oldch(189) = ch(189)
    If ch(189) Then
        GroupPush3D2.BackColor = &HFF&
     Else
        GroupPush3D2.BackColor = &H80&
     End If
   End If

'4/1/96 Engine Room fans
'ER Fan No 1
Call get_ch(411, 3377)
    If ch(3377) <> oldch(3377) Then
        oldch(3377) = ch(3377)
        If ch(3377) Then
            Image18(0).Picture = LoadPicture("c:\cotsim\icon\fan10.ico")
        Else
            Image18(0).Picture = LoadPicture("c:\cotsim\icon\fan15.ico")
        End If
    End If

'ER Fan No 2
Call get_ch(412, 3381)
    If ch(3381) <> oldch(3381) Then
        oldch(3381) = ch(3381)
        If ch(3381) Then
            Image18(1).Picture = LoadPicture("c:\cotsim\icon\fan10.ico")
        Else
            Image18(1).Picture = LoadPicture("c:\cotsim\icon\fan15.ico")
        End If
    End If

'ER Fan No 3
Call get_ch(413, 3333)
    If ch(3333) <> oldch(3333) Then
        oldch(3333) = ch(3333)
        If ch(3333) Then
            Image18(2).Picture = LoadPicture("c:\cotsim\icon\fan10.ico")
        Else
            Image18(2).Picture = LoadPicture("c:\cotsim\icon\fan15.ico")
        End If
    End If

'ER Fan No 4
Call get_ch(414, 3329)
    If ch(3329) <> oldch(3329) Then
        oldch(3329) = ch(3329)
        If ch(3329) Then
            Image18(3).Picture = LoadPicture("c:\cotsim\icon\fan10.ico")
        Else
            Image18(3).Picture = LoadPicture("c:\cotsim\icon\fan15.ico")
        End If
    End If

'=========================================================
'  Get 500 real numbers from d:
'  (costs 100 milli-seconds on a 50Mhz 486 DX2)
   Call get_pcreals


'  N.B. pc_reals(1) is a STRING but it still takes 10 mS
'  to update pml02.text ( on HAL_01 = 486 DX2 )
   enginerpm.Text = Format$(PC_REALS(50), "##0")
   RUDDERANG.Text = Format$(PC_REALS(53), "##0")
   xpiccy

End Sub

Private Sub xmover()
   
   Timer1.Enabled = False
   xclick = 0
   
End Sub

Private Sub xpiccy()

 Text12.Text = moveno$

'engine revs
'fuel lever
'   aivopen = Val(vpos00.Text)
 
 'cam position indicators
 j = 15
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = False Then      'cam astern
        Shape25.BackStyle = 0
        Shape26.BackStyle = 1
    Else
        Shape25.BackStyle = 1
        Shape26.BackStyle = 0
    End If
End If
 'ahead vv
 j = 133
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text1.BackColor = &HFFFFFF
    Else
        Text1.BackColor = &H808080
    End If
End If
 'stop vv
 j = 137
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text2.BackColor = &HFFFFFF
        'text18.Text = "True"
    Else
        Text2.BackColor = &H808080
        'text18.Text = "False"
    End If
End If
 'astern vv
 j = 181
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text3.BackColor = &HFFFFFF
        'text20.Text = "True"
    Else
        Text3.BackColor = &H808080
        'text20.Text = "False"
    End If
End If
 'gov limit
 j = 191
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text4.BackColor = &HFFFFFF
        'text22.Text = "True"
    Else
        Text4.BackColor = &H808080
        'text22.Text = "False"
    End If
End If
 'slow turn
 j = 129
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text5.BackColor = &HFFFFFF
        'text24.Text = "True"
    Else
        Text5.BackColor = &H808080
        'text24.Text = "False"
    End If
End If
 'fail start
 j = 183
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text6.BackColor = &HFFFFFF
        'text26.Text = "True"
    Else
        Text6.BackColor = &H808080
        'text26.Text = "False"
    End If
End If
 'air start vv open
 j = 187
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text7.BackColor = &HFFFFFF
        'text28.Text = "True"
    Else
        Text7.BackColor = &H808080
        'text28.Text = "False"
    End If
End If
 'air start block
 j = 185
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text8.BackColor = &HFFFFFF
        'text30.Text = "True"
    Else
        Text8.BackColor = &H808080
        'text30.Text = "False"
    End If
End If
 'turn gear
 j = 179
 If ch(j) <> oldch(j) Then
    oldch(j) = ch(j)
    If ch(j) = True Then
        Text9.BackColor = &HFFFFFF
        'text32.Text = "True"
    Else
        Text9.BackColor = &H808080
        'text32.Text = "False"
    End If
End If

 'blower on
 If ch(333) = True Then
   Text14.BackColor = &HFF00&
 Else
   Text14.BackColor = &H808080
 End If

 'blower stop
 If ch(369) = True Then
   Text15.BackColor = &HFFFFFF
   'text35.Text = "True"
 Else
   Text15.BackColor = &H808080
   'text35.Text = "False"
 End If

'blower auto
 If ch(381) = True Then
   Text16.BackColor = &HFFFF&
 Else
   Text16.BackColor = &H808080
 End If
'
' 17/5/95  Auxy blower...
   If pump_status(0) = 0 Then
      ch(333) = False
      ch(369) = True
      ch(381) = False
   ElseIf pump_status(0) = 1 Then
      ch(333) = True
      ch(369) = False
      ch(381) = False
   ElseIf pump_status(0) = 2 Then
      ch(333) = False
      ch(369) = True
      ch(381) = True
  Else
      ch(333) = True
      ch(369) = False
      ch(381) = True
  End If


 For j = 0 To 5119
   oldch(j) = ch(j)
 Next

 oldconstat = constat

 Screen.MousePointer = 0

End Sub

