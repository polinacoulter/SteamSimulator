VERSION 5.00
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form B_form 
   Caption         =   "B"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   0
      Left            =   11460
      TabIndex        =   2
      Top             =   2400
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      OrientVertical  =   0   'False
      EnabledBtmRight =   0   'False
      BackColour      =   12632256
   End
   Begin PumpPrj.Pump Pump1 
      Height          =   615
      Index           =   0
      Left            =   10800
      TabIndex        =   0
      Top             =   2640
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin PumpPrj.Pump Pump1 
      Height          =   615
      Index           =   1
      Left            =   12060
      TabIndex        =   1
      Top             =   2640
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   6
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   1
      Left            =   12900
      TabIndex        =   4
      Top             =   2460
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      OrientVertical  =   0   'False
      EnabledBtmRight =   0   'False
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   2
      Left            =   13860
      TabIndex        =   5
      Top             =   2100
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      OrientVertical  =   0   'False
      EnabledBtmRight =   0   'False
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   3
      Left            =   10920
      TabIndex        =   6
      Top             =   1740
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BottomColour    =   65280
      OrientVertical  =   0   'False
      EnabledTopLeft  =   0   'False
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   4
      Left            =   12420
      TabIndex        =   7
      Top             =   1500
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BottomColour    =   65280
      OrientVertical  =   0   'False
      EnabledTopLeft  =   0   'False
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   5
      Left            =   9960
      TabIndex        =   8
      Top             =   2580
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      BottomColour    =   65280
      OrientVertical  =   0   'False
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   7
      Left            =   9660
      TabIndex        =   9
      Top             =   2040
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   8
      Left            =   10380
      TabIndex        =   10
      Top             =   2040
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   9
      Left            =   2100
      TabIndex        =   11
      Top             =   2640
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   10
      Left            =   1080
      TabIndex        =   12
      Top             =   2700
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   13
      Top             =   2640
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   14
      Top             =   2640
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   13
      Left            =   5700
      TabIndex        =   15
      Top             =   2700
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   14
      Left            =   6660
      TabIndex        =   16
      Top             =   2640
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   15
      Left            =   7740
      TabIndex        =   17
      Top             =   2640
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   16
      Left            =   8700
      TabIndex        =   18
      Top             =   2640
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   315
      Index           =   17
      Left            =   60
      TabIndex        =   19
      Top             =   1380
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
      Height          =   255
      Index           =   18
      Left            =   5100
      TabIndex        =   20
      Top             =   1560
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      TopColour       =   65280
      BottomColour    =   65280
      BackColour      =   12632256
   End
   Begin VB.Line Line32 
      X1              =   7020
      X2              =   7020
      Y1              =   3480
      Y2              =   3720
   End
   Begin VB.Line Line31 
      X1              =   4140
      X2              =   4140
      Y1              =   2700
      Y2              =   3000
   End
   Begin VB.Line Line30 
      X1              =   3060
      X2              =   3060
      Y1              =   2700
      Y2              =   3000
   End
   Begin VB.Line Line29 
      X1              =   2160
      X2              =   2160
      Y1              =   2760
      Y2              =   3000
   End
   Begin VB.Line Line28 
      X1              =   5160
      X2              =   4980
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line27 
      X1              =   4980
      X2              =   4980
      Y1              =   3120
      Y2              =   3060
   End
   Begin VB.Line Line26 
      X1              =   4980
      X2              =   4980
      Y1              =   1200
      Y2              =   2880
   End
   Begin VB.Line Line25 
      X1              =   5160
      X2              =   5160
      Y1              =   1200
      Y2              =   3000
   End
   Begin VB.Line Line24 
      X1              =   120
      X2              =   120
      Y1              =   900
      Y2              =   3120
   End
   Begin VB.Line Line15 
      Index           =   1
      X1              =   1140
      X2              =   1140
      Y1              =   3000
      Y2              =   2880
   End
   Begin VB.Line Line23 
      X1              =   9720
      X2              =   10440
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line22 
      X1              =   10440
      X2              =   10440
      Y1              =   3120
      Y2              =   3060
   End
   Begin VB.Line Line21 
      X1              =   10440
      X2              =   10440
      Y1              =   2880
      Y2              =   1920
   End
   Begin VB.Line Line20 
      X1              =   9720
      X2              =   9720
      Y1              =   3000
      Y2              =   1920
   End
   Begin VB.Line Line19 
      X1              =   10620
      X2              =   10620
      Y1              =   1920
      Y2              =   1680
   End
   Begin VB.Line Line18 
      X1              =   10620
      X2              =   12840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line17 
      X1              =   12840
      X2              =   12840
      Y1              =   2220
      Y2              =   1680
   End
   Begin VB.Line Line16 
      X1              =   12840
      X2              =   12840
      Y1              =   2640
      Y2              =   2340
   End
   Begin VB.Line Line15 
      Index           =   0
      X1              =   11400
      X2              =   11400
      Y1              =   2640
      Y2              =   1920
   End
   Begin VB.Line Line14 
      X1              =   9720
      X2              =   11400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line13 
      X1              =   11520
      X2              =   11820
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line12 
      X1              =   11820
      X2              =   11820
      Y1              =   3120
      Y2              =   2940
   End
   Begin VB.Line Line11 
      X1              =   11820
      X2              =   12120
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line10 
      X1              =   120
      X2              =   10680
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line9 
      X1              =   420
      X2              =   420
      Y1              =   900
      Y2              =   3000
   End
   Begin VB.Line Line8 
      X1              =   11340
      X2              =   11940
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line7 
      X1              =   13440
      X2              =   13440
      Y1              =   2640
      Y2              =   2280
   End
   Begin VB.Line Line6 
      X1              =   12660
      X2              =   13440
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   11940
      X2              =   11940
      Y1              =   2640
      Y2              =   2280
   End
   Begin VB.Line Line4 
      X1              =   14280
      X2              =   14280
      Y1              =   900
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   11940
      X2              =   14280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   420
      X2              =   10800
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14700
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Shape Shape2 
      Height          =   1515
      Left            =   1860
      Top             =   240
      Width           =   195
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   7
      Left            =   8580
      Top             =   1920
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   6
      Left            =   7500
      Top             =   1920
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   5
      Left            =   6420
      Top             =   1920
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   4
      Left            =   5280
      Top             =   2040
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   3
      Left            =   720
      Top             =   2040
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   2
      Left            =   3900
      Top             =   1920
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   1
      Left            =   2820
      Top             =   1920
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   0
      Left            =   1740
      Top             =   1920
      Width           =   915
   End
End
Attribute VB_Name = "B_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
