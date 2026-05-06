VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "GRAPHS32.OCX"
Begin VB.Form GOVERNOR_LIM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Governor Limits"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7545
   Begin VB.VScrollBar VScroll4 
      Height          =   1215
      Left            =   6240
      Max             =   255
      TabIndex        =   45
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1215
      Left            =   5280
      Max             =   255
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   1215
      Left            =   4380
      Max             =   255
      TabIndex        =   40
      Top             =   6360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox AFR 
      Height          =   345
      Left            =   2760
      TabIndex        =   39
      Top             =   6300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox KDIG 
      Height          =   345
      Left            =   1440
      TabIndex        =   38
      Top             =   6300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox MF 
      Height          =   345
      Left            =   180
      TabIndex        =   35
      Top             =   6300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1875
      Index           =   2
      Left            =   3180
      Max             =   100
      TabIndex        =   30
      Top             =   180
      Value           =   100
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1140
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
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1875
      Index           =   0
      Left            =   6420
      Max             =   100
      TabIndex        =   13
      Top             =   2700
      Value           =   100
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      Left            =   1260
      Max             =   110
      TabIndex        =   12
      Top             =   4320
      Value           =   110
      Visible         =   0   'False
      Width           =   1695
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   1335
      Left            =   5760
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
      _Version        =   327680
      _ExtentX        =   2566
      _ExtentY        =   2355
      _StockProps     =   96
      BorderStyle     =   1
      BottomTitle     =   "Main Engine rpm"
      LeftTitle       =   "Governor Position %"
   End
   Begin VB.Line Line16 
      Visible         =   0   'False
      X1              =   5580
      X2              =   5580
      Y1              =   2520
      Y2              =   5280
   End
   Begin VB.Line Line15 
      Visible         =   0   'False
      X1              =   4620
      X2              =   4620
      Y1              =   3840
      Y2              =   5340
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   3540
      X2              =   3540
      Y1              =   4140
      Y2              =   5280
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   3840
      X2              =   3840
      Y1              =   3240
      Y2              =   5280
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "AI(102)"
      Height          =   195
      Left            =   5940
      TabIndex        =   43
      Top             =   6060
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "AI(101)"
      Height          =   195
      Left            =   4980
      TabIndex        =   42
      Top             =   6060
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "AI(100)"
      Height          =   195
      Left            =   4080
      TabIndex        =   41
      Top             =   6060
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "AFR"
      Height          =   195
      Left            =   2820
      TabIndex        =   37
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "KDIG"
      Height          =   195
      Left            =   1500
      TabIndex        =   36
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labelx 
      Alignment       =   2  'Center
      Caption         =   "MF"
      Height          =   195
      Left            =   180
      TabIndex        =   34
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Line right_line_MED 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3540
      X2              =   4620
      Y1              =   4140
      Y2              =   3840
   End
   Begin VB.Line left_line_MED 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   720
      X2              =   3540
      Y1              =   4140
      Y2              =   4140
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
   Begin VB.Label omega_label 
      Caption         =   "omega"
      Height          =   255
      Left            =   6240
      TabIndex        =   32
      Top             =   5580
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line left_line_LSS 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   720
      X2              =   3840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label gov_high_label 
      Caption         =   " %"
      Height          =   255
      Left            =   2880
      TabIndex        =   31
      Top             =   2280
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label gov_low_label 
      Caption         =   " %"
      Height          =   255
      Left            =   1740
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label GL 
      Caption         =   "GL"
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   3660
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label rpm 
      Caption         =   "rpm"
      Height          =   255
      Left            =   1260
      TabIndex        =   27
      Top             =   4620
      Visible         =   0   'False
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
   Begin VB.Line right_line_LSS 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3840
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
   Begin VB.Shape Restricted_band_shape 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   4080
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Restricted_band_label 
      Alignment       =   2  'Center
      Caption         =   "Restricted Band"
      Height          =   195
      Left            =   3540
      TabIndex        =   33
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "GOVERNOR_LIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim i As Integer


If medium_speed Then

  Restricted_band_shape.Visible = False
  Restricted_band_label.Visible = False
  For i = 0 To 10
    Xlabel(i).Caption = Format((i + 1) * 100, "###0")
  Next i
  
  left_line_LSS.Visible = False
  right_line_LSS.Visible = False
  
Else

  left_line_MED.Visible = False
  right_line_MED.Visible = False
  
End If



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

     
If medium_speed Then

  'For i = 3 To 18
  '  Parameter(i).Text = Int(VB_ENGINE(108 + i))  ' HULL etc etc
  'Next i
     gov_low = VB_ENGINE(119)    '  GOV_LOW_LIM
    gov_high = VB_ENGINE(120)    '  GOV_HIGH_LIM

  'enginerpm.Text = Format(VB_ENGINE(92), "##0")
          eng_rpm = VB_ENGINE(92)
      rpm.Caption = Format(VB_ENGINE(92), "##0.0")

   'gl.Text = Format(VB_ENGINE(93) * 10#, "##0")    ' GL 0-10
     gov_position = VB_ENGINE(93) * 10#
       GL.Caption = Format(VB_ENGINE(93) * 10#, "##0.0")

    omega_label.Caption = "OMEGA = " & Format(VB_ENGINE(128), "##0.000")
 
Else

    gov_low = LSS_reals(319)    '  GOV_LOW_LIM
   gov_high = LSS_reals(318)    '  GOV_HIGH_LIM
 
    ' Parameter(13).Text = Int(LSS_reals(327))   '  GOV_PSCAV_LOLIM
    ' Parameter(14).Text = Int(LSS_reals(326))   '  GOV_PSCAV_HILIM

    'enginerpm.Text = Format(LSS_reals(50), "##0")
    'GL.Text = Format(LSS_reals(56), "##0")
    
          eng_rpm = LSS_reals(50)
      rpm.Caption = Format(LSS_reals(50), "##0.0")
         
     gov_position = LSS_reals(56)
       GL.Caption = Format(LSS_reals(56), "##0.0")

    omega_label.Caption = "OMEGA = " & Format(LSS_reals(7), "##0.000")
    
     MF.Text = Format(LSS_reals(727), "##0.000")
     KDIG.Text = Format(LSS_reals(728), "##0.000")
     AFR.Text = Format(LSS_reals(729), "##0.000")
     
End If
     
If medium_speed Then

  ' Governor Position ( RED BLOB )
  gov.Left = eng_rpm / 1100# * (5940 - 660) + 660
  gov.Top = (100 - gov_position) / 100# * (5160 - 380) + 380
     
     
  ' Governor LOW Limit
  left_line_MED.Y1 = (100 - gov_low) * (5280 - 480) / 100# + 480
  left_line_MED.Y2 = left_line_MED.Y1
  right_line_MED.Y1 = left_line_MED.Y1

  ' Governor HIGH Limit
  right_line_MED.Y2 = (100 - gov_high) * (5280 - 480) / 100# + 480
  
Else

  ' Governor Position ( RED BLOB )
  gov.Left = eng_rpm / 110# * (5940 - 660) + 660
  gov.Top = (100 - gov_position) / 100# * (5160 - 380) + 380
     
     
  ' Governor LOW Limit
  left_line_LSS.Y1 = (100 - gov_low) * (5280 - 480) / 100# + 480
  left_line_LSS.Y2 = left_line_LSS.Y1
  right_line_LSS.Y1 = left_line_LSS.Y1

  ' Governor HIGH Limit
  right_line_LSS.Y2 = (100 - gov_high) * (5280 - 480) / 100# + 480
 
End If

End Sub

Private Sub VScroll1_Change(Index As Integer)

Select Case Index
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

Private Sub VScroll2_Change()

' Adjust
       nCHANNEL = 101
       nVALUE = 255 - VScroll2.Value
       Call setAI(nCHANNEL, nVALUE)

End Sub

Private Sub VScroll3_Change()

' Adjust PACHAI(I)
       nCHANNEL = 100
       nVALUE = 255 - VScroll3.Value
       Call setAI(nCHANNEL, nVALUE)

End Sub

Private Sub VScroll4_Change()

' Adjust
       nCHANNEL = 102
       nVALUE = 255 - VScroll4.Value
       Call setAI(nCHANNEL, nVALUE)

End Sub
