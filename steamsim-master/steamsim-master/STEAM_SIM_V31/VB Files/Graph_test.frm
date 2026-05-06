VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "GRAPHS32.OCX"
Begin VB.Form Ylabel 
   Caption         =   "Graph Test"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   1875
      Index           =   2
      Left            =   4440
      Max             =   100
      TabIndex        =   30
      Top             =   240
      Value           =   100
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3000
      Top             =   600
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1875
      Index           =   1
      Left            =   1980
      Max             =   100
      TabIndex        =   26
      Top             =   240
      Value           =   100
      Width           =   195
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1875
      Index           =   0
      Left            =   8820
      Max             =   100
      TabIndex        =   13
      Top             =   2760
      Value           =   100
      Width           =   195
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      Left            =   6960
      Max             =   110
      TabIndex        =   12
      Top             =   4440
      Value           =   110
      Width           =   1695
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   3975
      Left            =   780
      TabIndex        =   0
      Top             =   6540
      Width           =   6615
      _Version        =   327680
      _ExtentX        =   11668
      _ExtentY        =   7011
      _StockProps     =   96
      BorderStyle     =   1
      BottomTitle     =   "Main Engine rpm"
      LeftTitle       =   "Governor Position %"
   End
   Begin VB.Line left_line 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   720
      X2              =   3420
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label gov_high_label 
      Caption         =   " %"
      Height          =   255
      Left            =   3900
      TabIndex        =   31
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label gov_low_label 
      Caption         =   " %"
      Height          =   255
      Left            =   1500
      TabIndex        =   29
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label GL 
      Caption         =   "GL"
      Height          =   255
      Left            =   9060
      TabIndex        =   28
      Top             =   3540
      Width           =   915
   End
   Begin VB.Label rpm 
      Caption         =   "rpm"
      Height          =   255
      Left            =   7080
      TabIndex        =   27
      Top             =   4740
      Width           =   1395
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "110"
      Height          =   195
      Index           =   10
      Left            =   5820
      TabIndex        =   25
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "100"
      Height          =   195
      Index           =   9
      Left            =   5340
      TabIndex        =   24
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "90"
      Height          =   195
      Index           =   8
      Left            =   4860
      TabIndex        =   23
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "80"
      Height          =   195
      Index           =   7
      Left            =   4380
      TabIndex        =   22
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "70"
      Height          =   195
      Index           =   6
      Left            =   3900
      TabIndex        =   21
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "60"
      Height          =   195
      Index           =   5
      Left            =   3420
      TabIndex        =   20
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "50"
      Height          =   195
      Index           =   4
      Left            =   2940
      TabIndex        =   19
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "40"
      Height          =   195
      Index           =   3
      Left            =   2460
      TabIndex        =   18
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "30"
      Height          =   195
      Index           =   2
      Left            =   1980
      TabIndex        =   17
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "20"
      Height          =   195
      Index           =   1
      Left            =   1500
      TabIndex        =   16
      Top             =   5520
      Width           =   375
   End
   Begin VB.Line Xtick 
      Index           =   10
      X1              =   5520
      X2              =   5520
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   9
      X1              =   6000
      X2              =   6000
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   8
      X1              =   5040
      X2              =   5040
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   7
      X1              =   4560
      X2              =   4560
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   6
      X1              =   4080
      X2              =   4080
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   5
      X1              =   3600
      X2              =   3600
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   4
      X1              =   3120
      X2              =   3120
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   3
      X1              =   2640
      X2              =   2640
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   2
      X1              =   2160
      X2              =   2160
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Xtick 
      Index           =   1
      X1              =   1680
      X2              =   1680
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Label Xlabel 
      Alignment       =   2  'Center
      Caption         =   "10"
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   15
      Top             =   5520
      Width           =   375
   End
   Begin VB.Line Xtick 
      Index           =   0
      X1              =   1200
      X2              =   1200
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Label Label1 
      Caption         =   "Main Engine rpm"
      Height          =   195
      Left            =   6120
      TabIndex        =   14
      Top             =   5220
      Width           =   1455
   End
   Begin VB.Shape gov 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   660
      Shape           =   3  'Circle
      Top             =   5160
      Width           =   135
   End
   Begin VB.Line Line14 
      X1              =   600
      X2              =   840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line13 
      X1              =   600
      X2              =   840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "100"
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   11
      Top             =   420
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "90"
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   10
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "80"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   9
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "70"
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "60"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "50"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "40"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   3300
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "30"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "20"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Ylabel 
      Alignment       =   1  'Right Justify
      Caption         =   "10"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Governor Position %"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line right_line 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3420
      X2              =   5580
      Y1              =   3240
      Y2              =   2520
   End
   Begin VB.Line Line10 
      X1              =   600
      X2              =   840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line9 
      X1              =   600
      X2              =   840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line8 
      X1              =   600
      X2              =   840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line7 
      X1              =   600
      X2              =   840
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      X1              =   600
      X2              =   840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line5 
      X1              =   600
      X2              =   840
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      X1              =   720
      X2              =   6000
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      X1              =   600
      X2              =   840
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   840
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   720
      Y1              =   480
      Y2              =   5280
   End
End
Attribute VB_Name = "Ylabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim i As Integer


Graph1.NumSets = 2
'RWJ Graph1.NumPoints = 16
Graph1.NumPoints = 11

Graph1.GraphType = 6  ' Line
'Graph1.GraphStyle = 5 'lines +  symbols
Graph1.GraphStyle = 0 'lines only (default)
'Graph1.GraphStyle = 1 'lines only (default)

Graph1.SymbolSize = 100
Graph1.AutoInc = 0

'You need to set up the Min and Max values of the Y
'axis.
' RWJ Graph1.YAxisMin = -10
Graph1.YAxisMin = 0
Graph1.YAxisStyle = 2
Graph1.XAxisStyle = 2
Graph1.YAxisMax = 100
Graph1.YAxisTicks = Graph1.YAxisMax / 10
' RWJ Graph1.XAxisTicks = 4
'Graph1.XAxisTicks = 10
Graph1.XAxisTicks = Graph1.YAxisMax / 10

Graph1.XAxisMax = 100

Graph1.SDKMouse = 1
Graph1.DrawMode = 3
Graph1.LeftTitleStyle = 0 'up

Graph1.ThisSet = 1
For i = 1 To 5
 Graph1.Data(i) = 60
Next i
For i = 6 To 11
 Graph1.Data(i) = 60 + 5 * (i - 6)
Next i

Graph1.ThisSet = 2
Graph1.Data(5) = 50

Graph1.DrawMode = 3

End Sub

Private Sub HScroll1_Change()

eng_rpm = HScroll1.Value

rpm.Caption = HScroll1.Value

End Sub

Private Sub Timer1_Timer()

' Governor position ( RED BLOB )
gov.Left = eng_rpm / 110# * (5940 - 660) + 660
gov.Top = (100 - gov_position) / 100# * (5160 - 380) + 380
     
     
' Governor LOW Limit
left_line.Y1 = (100 - gov_low) * (5280 - 480) / 100# + 480
left_line.Y2 = left_line.Y1
right_line.Y1 = left_line.Y1

' Governor HIGH Limit
right_line.Y2 = (100 - gov_high) * (5280 - 480) / 100# + 480

End Sub

Private Sub VScroll1_Change(index As Integer)

Select Case index
   Case 0
   
      gov_position = VScroll1(0).Value
      
      GL.Caption = VScroll1(0).Value
      
   Case 1
   
     gov_low = VScroll1(1).Value
     gov_low_label.Caption = gov_low
     
   Case 2
   
     gov_high = VScroll1(2).Value
     gov_high_label.Caption = gov_high
     
End Select

End Sub
