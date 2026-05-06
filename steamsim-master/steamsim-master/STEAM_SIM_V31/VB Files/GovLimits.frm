VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "GRAPHS32.OCX"
Begin VB.Form GovLimits 
   Caption         =   "Main Engine Governor Limits"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin GraphsLib.Graph Graph1 
      Height          =   6375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8775
      _Version        =   327680
      _ExtentX        =   15478
      _ExtentY        =   11245
      _StockProps     =   96
      BorderStyle     =   1
   End
End
Attribute VB_Name = "GovLimits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer


Graph1.GraphType = 6  'line
Graph1.GraphStyle = 1   'continuous line, default(0) is scatter
'NumOfPlots = 2
NumOfPlots = 1

Graph1.NumSets = 2
Graph1.NumPoints = 100 'Defines length of X axis
Graph1.AutoInc = 0     'Off

'Appearance
Graph1.BackgroundUse = 0 'Entire graphing window
Graph1.Palette = 0

Graph1.ForegroundUse = 4
Graph1.Foreground = 0 'set labels in black

Graph1.ForegroundUse = 6
Graph1.Foreground = 0 'set axes in black

Graph1.ForegroundUse = 7
Graph1.Foreground = 0 'set gridlines in black

Graph1.ForegroundUse = 1
Graph1.Foreground = 0 'set left axis title in black

Graph1.ForegroundUse = 2
Graph1.Foreground = 0 'set right axis title in black

Graph1.GridLineStyle = 2 'dotted lines for grid
Graph1.GridStyle = 3 'GridStyle '0 = no grid, 1= x only, 2= y only, 3 = both x and y lines in grid


Graph1.Background = 127  '0= Black 1=dk blue, 5=purp

'Axes
Graph1.TickEvery = 10  'X axis - show every 10 th tick
Graph1.LabelEvery = 10 'X-axis label every 10th

'Graph1.TickEvery = XLen / 10  'X axis - show every 5th tick

Graph1.YAxisStyle = 2         'Y axis - user defined scale
Graph1.YAxisMax = 400
Graph1.XAxisStyle = 2         'Y axis - user defined scale
Graph1.XAxisMax = 100


'Graph1.YAxisMin = -Amplitude
Graph1.YAxisMin = 0
Graph1.YAxisTicks = 10 'Number of ticks on positive and negative arms

Graph1.LeftTitle = " %"
Graph1.LeftTitleStyle = 1  'up

Graph1.RightTitle = " %"
Graph1.RightTitleStyle = 2 'down



'Graph1.AutoInc = 0        'off


'this it for timer............
    Graph1.ThisSet = 1

    For i = 1 To 100
        Graph1.ThisPoint = i
        Graph1.GraphData = i
    Next
    
    Graph1.ThisSet = 2
    Graph1.GraphStyle = 1 'symbols
    For i = 1 To 100
        Graph1.ThisPoint = i
        Graph1.GraphData = i * 2
    Next

'..............................




'draw it
Graph1.DrawMode = 2


End Sub

