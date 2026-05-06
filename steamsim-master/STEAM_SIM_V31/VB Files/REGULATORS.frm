VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{D6ECA430-4BD6-11D3-AA45-00C04FA979C0}#20.0#0"; "SquareGauge.ocx"
Begin VB.Form REGULATORS 
   Caption         =   "REGULATORS"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8955
      Left            =   180
      Top             =   180
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   15796
      Begin SquareGauge.Gauge Gauge1 
         Height          =   1035
         Left            =   4320
         TabIndex        =   1
         Top             =   1800
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge2 
         Height          =   1035
         Left            =   5640
         TabIndex        =   2
         Top             =   1800
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge3 
         Height          =   1035
         Left            =   4320
         TabIndex        =   3
         Top             =   3000
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge4 
         Height          =   1035
         Left            =   4320
         TabIndex        =   4
         Top             =   4140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge5 
         Height          =   1035
         Left            =   4320
         TabIndex        =   5
         Top             =   5340
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge6 
         Height          =   1035
         Left            =   4320
         TabIndex        =   6
         Top             =   6480
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge7 
         Height          =   1035
         Left            =   4320
         TabIndex        =   7
         Top             =   7620
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge8 
         Height          =   1035
         Left            =   4320
         TabIndex        =   8
         Top             =   720
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1826
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "CONTROL AIR SUPPLY"
         Height          =   315
         Left            =   3300
         TabIndex        =   0
         Top             =   420
         Width           =   1995
      End
      Begin VB.Shape REGULATORS 
         BorderWidth     =   2
         Height          =   8775
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5955
      End
   End
End
Attribute VB_Name = "REGULATORS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
