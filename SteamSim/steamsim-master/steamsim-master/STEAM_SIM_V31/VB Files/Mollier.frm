VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "Graphs32.ocx"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Mollier 
   Caption         =   "Mollier Chart"
   ClientHeight    =   10275
   ClientLeft      =   3345
   ClientTop       =   3735
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   12165
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9135
      Left            =   0
      Top             =   60
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   16113
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   0
         Top             =   6720
      End
      Begin VB.CheckBox Check1 
         Caption         =   "SI Units"
         Height          =   375
         Left            =   8760
         TabIndex        =   3
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Grid Lines"
         Height          =   375
         Left            =   8760
         TabIndex        =   2
         Top             =   8040
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7560
         TabIndex        =   1
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1260
         TabIndex        =   0
         Top             =   6720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1605
         Left            =   240
         TabIndex        =   4
         Top             =   7080
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   2831
         _Version        =   393216
         Rows            =   6
         Cols            =   7
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         AllowUserResizing=   3
      End
      Begin GraphsLib.Graph Graph1 
         Height          =   6615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   10575
         _Version        =   327680
         _ExtentX        =   18653
         _ExtentY        =   11668
         _StockProps     =   96
         BorderStyle     =   1
         AutoInc         =   0
         BackgroundStyle =   "0~0~4~0~0~0~0"
         ColorData       =   "13~12~9~6~14~1"
         CurveType       =   1
         FontSize        =   "170~130~100~100"
         GraphTitle      =   "Enthalpy - Entropy Diagram (Water/Steam)"
         GraphType       =   9
         GridLineStyle   =   5
         GridStyle       =   3
         LeftTitleStyle  =   1
         LegendSize      =   74
         LegendText      =   "Constant Entropy~HP Inlet~HP Bleed~IP~LP Bleed~LP Outlet~ "
         NumPoints       =   6
         NumSets         =   6
         OverlayColor    =   "15"
         OverlayGraph    =   1
         OverlayGraphStyle=   4
         OverlayPattern  =   "0"
         OverlaySymbol   =   "4"
         PatternedLines  =   1
         RandomData      =   0
         SymbolData      =   "1~13~5~5~5~13"
         SymbolSize      =   85
         ThickLines      =   0
         XAxisMax        =   2.4
         XAxisMin        =   1.4
         XAxisMinorTicks =   -1
         XAxisPos        =   2
         XAxisStyle      =   2
         XAxisTicks      =   5
         YAxisMax        =   "3600~0"
         YAxisMin        =   "2000~0"
         YAxisMinorTicks =   "-1~0"
         YAxisStyle      =   "2~0"
         YAxisTicks      =   "16~1"
         OverlayTrendSets=   "0"
      End
   End
End
Attribute VB_Name = "Mollier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'first 5 series used to graph HP inlet,HP bleed,IP
'LP Bleed and LP outlet
Option Explicit

Const NUMBER_OF_POINTS = 6
Const NUMBER_OF_SERIES = 6
Const SI_MAX_ENTROPY = 10
Const SI_MIN_ENTROPY = 6
Const SI_MAX_ENTHALPY = 3600
Const SI_MIN_ENTHALPY = 2000
Const BS_MAX_ENTROPY = 2.4
Const BS_MIN_ENTROPY = 1.4
Const BS_MAX_ENTHALPY = 1540
Const BS_MIN_ENTHALPY = 850

Private Sub Check1_Click()
'allow user to change units of graph
'flip flag allowing timer to change graph axis and values
  SIUnits = Not SIUnits
End Sub

Private Sub Check2_Click()
'linear grid lines
 If Graph1.GridLineStyle = 4 Then
   Graph1.GridLineStyle = 5
 Else
   Graph1.GridLineStyle = 4
 End If
 'redraw graph
 Graph1.DrawMode = 3
End Sub

Private Sub Command1_Click()
Dim H As Long
 'Testing only
  Pressure(1) = 62.09
  Pressure(2) = 17.18
  Pressure(3) = 3.81
  Pressure(4) = 0.55
  Pressure(5) = 0.05
  
  Temperature(1) = 503.7
  Temperature(2) = 339.5
  Temperature(3) = 179.6
  Temperature(4) = 83.7
  Temperature(5) = 32.9
  
  
  Enthalpy(1) = 3427
  Enthalpy(2) = 3120
  Enthalpy(3) = 2819
  Enthalpy(4) = 2524
  Enthalpy(5) = 2254
  
  Dryness(1) = 1
  Dryness(2) = 1
  Dryness(3) = 1
  Dryness(4) = 0.946
  Dryness(5) = 0.873
End Sub

Private Sub Form_Load()
Dim i As Integer




 'set si unints   to  false
 'SIUnits = False
 'set up grid
 'resize column zero width
 
 
 
 MSFlexGrid1.ColWidth(0) = 1400 'width of row header
 MSFlexGrid1.ColWidth(1) = 1350 'width of row header
 ' centre text in cells
 For i = 0 To MSFlexGrid1.Rows
    MSFlexGrid1.ColAlignment(i) = 4
 Next
 
 ' === 22/01/02 ===
On Error Resume Next
Call JYAScreenInit(Me, id_Mollier)

 
 
 'column headers
 MSFlexGrid1.TextMatrix(0, 1) = "Constant Entropy"
 MSFlexGrid1.TextMatrix(0, 2) = "HP Inlet"
 MSFlexGrid1.TextMatrix(0, 3) = "HP Bleed"
 MSFlexGrid1.TextMatrix(0, 4) = "IP"
 MSFlexGrid1.TextMatrix(0, 5) = "LP Bleed"
 MSFlexGrid1.TextMatrix(0, 6) = "LP Outlet"
 
 
 ' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

 'Enable timer
 Timer1.Enabled = True
 
 
End Sub

Private Sub GraphUSA()
Dim r As Integer
Dim Label As String

   'range of axis
   MAXENTHALPY = BS_MAX_ENTHALPY
   MINENTHALPY = BS_MIN_ENTHALPY
   MAXENTROPY = BS_MAX_ENTROPY
   MINENTROPY = BS_MIN_ENTROPY
   
   'Grid row headers
   MSFlexGrid1.TextMatrix(GRID_ROW_PRESSURE, 0) = "Pressure (psi)"
   MSFlexGrid1.TextMatrix(GRID_ROW_ENTHALPY, 0) = "Enthalpy (Btu/lb)"
   MSFlexGrid1.TextMatrix(GRID_ROW_TEMPERATURE, 0) = "Temperature(" & Chr(176) & "F)"
   MSFlexGrid1.TextMatrix(GRID_ROW_ENTROPY, 0) = "Entropy (Btu/lb" & Chr(176) & "R)"
   MSFlexGrid1.TextMatrix(GRID_ROW_DRYNESS, 0) = "Dryness "
   
   'set up graph titles
   Graph1.LeftTitle = "Specific Enthalpy (Btu/lb)"
   Graph1.LeftTitleStyle = 1
   Graph1.BottomTitle = "Specific Entropy (Btu/lb " & Chr(176) & "R)"
   
   'set up graph
   Graph1.NUMPOINTS = NUMBER_OF_POINTS
   Graph1.NumSets = NUMBER_OF_SERIES
   
   Graph1.XAxisTicks = 5
   Graph1.XAxisMinorTicks = -3
   Graph1.XAxisMax = MAXENTROPY
   Graph1.XAxisMin = MINENTROPY
   Graph1.XAxisStyle = 2 'user defined
   'X axis labels
   Label = "1.4" + Chr$(9) + "1.6" + Chr$(9) + "1.8" + Chr$(9) + "2.0" + Chr$(9) + "2.2" + Chr$(9) + "2.4"
   r = VBAGLabels(6, Label)
   'Y axis
   Graph1.YAxisMax = MAXENTHALPY
   Graph1.YAxisMin = MINENTHALPY
   Graph1.YAxisTicks = 15
   Graph1.YAxisStyle = 2 'user defined

   'Redraw New graph
   Graph1.DrawMode = 3
End Sub

Private Sub GraphUK()
Dim r As Integer
Dim Label As String
   ' ranges of axis
   MAXENTHALPY = SI_MAX_ENTHALPY
   MINENTHALPY = SI_MIN_ENTHALPY
   MAXENTROPY = SI_MAX_ENTROPY
   MINENTROPY = SI_MIN_ENTROPY
   
   'table headers
   MSFlexGrid1.TextMatrix(GRID_ROW_PRESSURE, 0) = "Pressure (Bar)"
   MSFlexGrid1.TextMatrix(GRID_ROW_ENTHALPY, 0) = "Enthalpy (kJ/kg)"
   MSFlexGrid1.TextMatrix(GRID_ROW_TEMPERATURE, 0) = "Temperature(" & Chr(176) & "C)"
   MSFlexGrid1.TextMatrix(GRID_ROW_ENTROPY, 0) = "Entropy (kJ/kg" & Chr(176) & "K)"
   MSFlexGrid1.TextMatrix(GRID_ROW_DRYNESS, 0) = "Dryness"
   
   'set up graph titles
   Graph1.LeftTitle = "Specific Enthalpy (kJ/kg)"
   Graph1.LeftTitleStyle = 1
   Graph1.BottomTitle = "Specific Entropy (kJ/kg " & Chr(176) & "K)"
   
   'set up graph
   Graph1.NUMPOINTS = NUMBER_OF_POINTS
   Graph1.NumSets = NUMBER_OF_SERIES
   ' x axis
   Graph1.XAxisTicks = 4
   Graph1.XAxisMax = MAXENTROPY
   Graph1.XAxisMin = MINENTROPY
   Graph1.XAxisMinorTicks = -4
   Graph1.XAxisStyle = 2 ' user defined
   'x axis labels
   Label = "6" + Chr$(9) + "7" + Chr$(9) + "8" + Chr$(9) + "9" + Chr$(9) + "10"
   r = VBAGLabels(5, Label)
   'Y axis
   Graph1.YAxisMax = MAXENTHALPY
   Graph1.YAxisMin = MINENTHALPY
   Graph1.YAxisTicks = 16
  
   Graph1.YAxisStyle = 2 'user defined

   'Redraw graph
   Graph1.DrawMode = 3
End Sub

Private Sub Form_Resize()

' === 22/01/02 ===
'On Error Resume Next
'If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
'    HavenStretch1.Visible = False
'    Call JYAScreenResize(Me, id_Mollier)
'    HavenStretch1.Visible = True
'End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
  SIUnits = False
  ResetGrid
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

  If SIUnits Then
   GraphUK
 Else
   GraphUSA
 End If
 
    ' 15/7/99 boil_run_flag is true when RUNNING the VB TEST
    If Not boil_run_flag Then
        For i = 1 To 5
                    Pressure(i) = LSS_reals(59 + i)
                    Enthalpy(i) = LSS_reals(64 + i)
                 Temperature(i) = LSS_reals(69 + i)
                     Dryness(i) = LSS_reals(74 + i)
        Next i
    End If
            
            
 'plot points
 UpdateGraph
 
End Sub

