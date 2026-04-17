VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "Graphs32.ocx"
Begin VB.Form LUTableEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "P_V Diagram Look Up Design"
   ClientHeight    =   8985
   ClientLeft      =   1425
   ClientTop       =   1905
   ClientWidth     =   11850
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8985
   ScaleWidth      =   11850
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   9120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   9120
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   9120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   9120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   480
      Width           =   615
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   6
      Top             =   480
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327680
      BuddyControl    =   "Text1(0)"
      BuddyDispid     =   196609
      BuddyIndex      =   0
      OrigLeft        =   9720
      OrigTop         =   480
      OrigRight       =   9960
      OrigBottom      =   855
      Max             =   1000
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load P-Crank"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save P-Crank"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load P-V"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save P_V"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Quit"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   7800
      Width           =   1335
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8775
      _Version        =   327680
      _ExtentX        =   15478
      _ExtentY        =   13150
      _StockProps     =   96
      BottomTitle     =   "Crank Angle (deg) from FPO"
      GraphData       =   "10~10~10~10~10~10~10~10~10~10^20~2~20~20~20~20~20~20~20~3^60~80~100~40~50~0~40~80~40~20^50~20~40~80~40~20~40~100~50~50^30"
      GraphTitle      =   "Drag and Drop Graph"
      GraphType       =   6
      Hot             =   1
      LeftTitle       =   "% Increase on Compression Pressure"
      NumPoints       =   10
      NumSets         =   4
      RandomData      =   0
      SymbolData      =   "6"
      ThickLines      =   0
      YAxisMax        =   "100~0"
      YAxisStyle      =   "1~0"
      YAxisTicks      =   "5~1"
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   13
      Top             =   1200
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327680
      OrigLeft        =   9720
      OrigTop         =   480
      OrigRight       =   9960
      OrigBottom      =   855
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   14
      Top             =   1920
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327680
      OrigLeft        =   9720
      OrigTop         =   480
      OrigRight       =   9960
      OrigBottom      =   855
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   3
      Left            =   9720
      TabIndex        =   15
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327680
      OrigLeft        =   9720
      OrigTop         =   480
      OrigRight       =   9960
      OrigBottom      =   855
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   375
      Index           =   4
      Left            =   9720
      TabIndex        =   16
      Top             =   3360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327680
      OrigLeft        =   9720
      OrigTop         =   480
      OrigRight       =   9960
      OrigBottom      =   855
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Rate of fall just before Exh V/V closes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   10080
      TabIndex        =   20
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Degrees held at Scxav Press after bdc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   10080
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Rate of rise after min when Air V/V open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   10080
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Proportional fall below Scav wrt  1bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   10080
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Exh V/V decay Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   10080
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "LUTableEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Graphics Server 4 Demos - Demo Number = 3
'By The Technical Support Team
'
'This example shows how to set up a interactive graph
'using the hothit function.
'
'Copyright (c) 1991 - 1995 Bits Per Second Ltd
'*********************************************'

Dim GraphData() As Double
Dim gbdraw As Integer

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
'save current Graphdata values to disc
Dim filenum As Integer
Dim CharStr As String

filenum = FreeFile

Open "C:\HAVEN\LU_PV.dat" For Output Access Write As #filenum

For i = LBound(PVLookUp) To UBound(PVLookUp)
    CharStr = CStr(PVLookUp(i))
    Write #filenum, CharStr
Next
Close #filenum

End Sub

Private Sub Command3_Click()
Dim filenum As Integer
Dim CharStr As String
Dim i As Integer

filenum = FreeFile
EditPVLU = True

Open "C:\HAVEN\LU_PV.dat" For Input Access Read As #filenum

For i = LBound(PVLookUp) To UBound(PVLookUp)
    Input #filenum, CharStr
    PVLookUp(i) = CSng(CharStr)
    Graph1.ThisPoint = i
    Graph1.GraphData = PVLookUp(i)
Next
    
    
Close #filenum
Graph1.DrawMode = 3
Graph1.MousePointer = 0

End Sub

Private Sub Command4_Click()
'save current Graphdata values for P-Crank to disc
Dim filenum As Integer
Dim CharStr As String

filenum = FreeFile

Open "C:\HAVEN\LU_PC.dat" For Output Access Write As #filenum

For i = LBound(PCLookUp) To UBound(PCLookUp)
    CharStr = CStr(PCLookUp(i))
    Write #filenum, CharStr
Next
Close #filenum


End Sub

Private Sub Command5_Click()
Dim filenum As Integer
Dim CharStr As String
Dim i As Integer

filenum = FreeFile
EditPVLU = False
Open "C:\HAVEN\LU_PC.dat" For Input Access Read As #filenum

For i = LBound(PCLookUp) To UBound(PCLookUp)
    Input #filenum, CharStr
    PCLookUp(i) = CSng(CharStr)
    Graph1.ThisPoint = i
    Graph1.GraphData = PCLookUp(i)
Next
    
    
Close #filenum
Graph1.DrawMode = 3
Graph1.MousePointer = 0
End Sub

Private Sub form_Load()
'sizeform Me

'Set up a basic graph
Graph1.NumSets = 1
Graph1.NumPoints = 16
Graph1.GraphType = 6
Graph1.GraphStyle = 5 'lines +  symbols
'Graph1.GraphStyle = 0 'lines only (default)

Graph1.SymbolSize = 100
Graph1.AutoInc = 0

'You need to set up the Min and Max values of the Y
'axis.
Graph1.YAxisMin = -10
Graph1.YAxisStyle = 2
Graph1.XAxisStyle = 2
Graph1.YAxisMax = 100
Graph1.YAxisTicks = Graph1.YAxisMax / 10
Graph1.XAxisTicks = 4

Graph1.SDKMouse = 1
Graph1.DrawMode = 3
Graph1.LeftTitleStyle = 1 'up



'initialise graphdata to 50% throughout
For i = LBound(PVLookUp) To UBound(PVLookUp)
        PVLookUp(i) = 50#
Next

'Extra stuff to set up Power Cards
UpDown1(0).Value = ExhVVDecayRate * 100
UpDown1(1).Value = MinPressProportion * 100
UpDown1(2).Value = ScavInletRiseRate * 100
UpDown1(3).Value = ConstScavPAngle
UpDown1(4).Value = ScavPDecayRate * 100



End Sub

Private Sub Graph1_HotHit(hitset As Integer, hitpoint As Integer)

'hitpoint and hitset are the
'current point and set of the graph
Graph1.ThisSet = hitset
Graph1.ThisPoint = hitpoint
Graph1.MousePointer = 3
gbdraw = True
End Sub

Private Sub Graph1_SDKPress(pressStatus As Integer, pressX As Double, pressY As Double, PressDataX As Double, pressdatay As Double)
Dim i As Integer

If pressStatus <> 0 Then
If pressX > (Graph1.SDKInfo(7) + Graph1.SDKInfo(5)) Then
gbdraw = False
End If
If pressY > (Graph1.SDKInfo(6) + Graph1.SDKInfo(8)) Then
gbdraw = False
End If
End If
If pressStatus = 0 And gbdraw Then
    Debug.Print "saving data in SDKPress"
    
    
    Graph1.GraphData = pressdatay
    For i = 1 To Graph1.NumPoints
        Graph1.ThisPoint = i
        PVLookUp(i) = Graph1.GraphData
        PCLookUp(i) = Graph1.GraphData
    Next
    Graph1.DrawMode = 3
    Graph1.MousePointer = 0
    gbdraw = False
End If
End Sub

Private Sub Graph1_SDKTrack(trackx As Double, tracky As Double, TrackDataX As Double, trackdatay As Double)
57 If gbdraw Then
'Debug.Print "saving data in SDKTrack"

58 Graph1.GraphData = trackdatay
59 Graph1.DrawMode = 3
60 End If
End Sub

Private Sub UpDown1_Change(Index As Integer)
Select Case Index
Case 0
    ExhVVDecayRate = UpDown1(Index).Value / 100
    Text1(Index).Text = ExhVVDecayRate
Case 1
    MinPressProportion = UpDown1(Index).Value / 100
    Text1(Index).Text = MinPressProportion
Case 2
    ScavInletRiseRate = UpDown1(Index).Value / 100
    Text1(Index).Text = ScavInletRiseRate
Case 3
    ConstScavPAngle = UpDown1(Index).Value
    Text1(Index).Text = ConstScavPAngle
Case 4
    ScavPDecayRate = UpDown1(Index).Value / 100
    Text1(Index).Text = ScavPDecayRate


End Select

End Sub
