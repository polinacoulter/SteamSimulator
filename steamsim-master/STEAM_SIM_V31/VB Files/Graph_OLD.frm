VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "GRAPHS32.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form GraphFrm_OLD 
   Caption         =   "Multi Channel Recorder"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   16575
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   840
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   105
      Max             =   255
      TabIndex        =   7
      Top             =   8415
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8100
      Left            =   120
      Top             =   120
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   14288
      Begin VB.Frame LegendFrame 
         Caption         =   "Legend"
         Height          =   6270
         Left            =   12480
         TabIndex        =   13
         Top             =   75
         Width           =   3495
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   360
            Left            =   2520
            TabIndex        =   28
            Top             =   4920
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   345
            Left            =   240
            TabIndex        =   27
            Top             =   5550
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            BorderWidth     =   5
            Index           =   0
            X1              =   240
            X2              =   720
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   5
            Index           =   1
            X1              =   225
            X2              =   705
            Y1              =   1365
            Y2              =   1365
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   5
            Index           =   2
            X1              =   240
            X2              =   720
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFF00&
            BorderWidth     =   5
            Index           =   3
            X1              =   240
            X2              =   720
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   4
            X1              =   240
            X2              =   720
            Y1              =   3255
            Y2              =   3255
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF00FF&
            BorderWidth     =   5
            Index           =   5
            X1              =   240
            X2              =   720
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   5
            Index           =   6
            X1              =   240
            X2              =   720
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderWidth     =   5
            Index           =   7
            X1              =   240
            X2              =   720
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   21
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   975
            TabIndex        =   20
            Top             =   1320
            Width           =   2400
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   2
            Left            =   960
            TabIndex        =   19
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Index           =   3
            Left            =   960
            TabIndex        =   18
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   4
            Left            =   930
            TabIndex        =   17
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Index           =   5
            Left            =   960
            TabIndex        =   16
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   6
            Left            =   960
            TabIndex        =   15
            Top             =   4305
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   7
            Left            =   960
            TabIndex        =   14
            Top             =   4920
            Width           =   2415
         End
      End
      Begin GraphsLib.Graph Graph1 
         Height          =   6195
         Left            =   180
         TabIndex        =   9
         Top             =   90
         Width           =   12255
         _Version        =   327680
         _ExtentX        =   21616
         _ExtentY        =   10927
         _StockProps     =   96
         Background      =   "15~-1~-1~-1~-1~-1~-1"
         ExtraData       =   "0~0~0~0~0"
         GraphType       =   22
         Labels          =   3
         NumSets         =   8
         Ticks           =   3
         XAxisPos        =   2
         YAxisMax        =   "100~0"
         YAxisStyle      =   "2~0"
         YAxisTicks      =   "10~1"
         LabelYFormat    =   ""
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Height          =   1605
         Left            =   165
         TabIndex        =   0
         Top             =   6375
         Width           =   15675
         Begin VB.CommandButton CmdPrint2 
            Caption         =   "Print"
            Height          =   375
            Left            =   14400
            TabIndex        =   25
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton CmdLoad 
            Caption         =   "Load"
            Height          =   375
            Left            =   4635
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   4620
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdHalt 
            Caption         =   "Halt"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6705
            TabIndex        =   11
            Top             =   1170
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Stop After First Page"
            Height          =   375
            Left            =   825
            TabIndex        =   8
            Top             =   1005
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   195
            Max             =   9
            TabIndex        =   5
            Top             =   495
            Value           =   9
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.CommandButton CmdSetup 
            Caption         =   "Set Up"
            Height          =   375
            Left            =   8160
            TabIndex        =   4
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   14385
            TabIndex        =   3
            Top             =   300
            Width           =   975
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6705
            TabIndex        =   2
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6705
            TabIndex        =   1
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Sample Interval:"
            Height          =   300
            Left            =   10755
            TabIndex        =   22
            Top             =   1185
            Width           =   1305
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "RECORDER HALTED"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   10245
            TabIndex        =   12
            Top             =   420
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.Label Label3 
            Caption         =   "Speed: "
            Height          =   300
            Left            =   12210
            TabIndex        =   10
            Top             =   1185
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Scroll Graph"
            Height          =   255
            Left            =   105
            TabIndex        =   6
            Top             =   255
            Width           =   2535
         End
      End
      Begin GraphsLib.Graph Graph2 
         Height          =   6210
         Left            =   120
         TabIndex        =   26
         Top             =   90
         Visible         =   0   'False
         Width           =   12255
         _Version        =   327680
         _ExtentX        =   21616
         _ExtentY        =   10954
         _StockProps     =   96
         Background      =   "15~-1~-1~-1~-1~-1~-1"
         ExtraData       =   "0~0~0~0~0"
         GraphType       =   22
         Labels          =   3
         NumSets         =   8
         Ticks           =   3
         XAxisPos        =   2
         YAxisMax        =   "100~0"
         YAxisStyle      =   "2~0"
         YAxisTicks      =   "10~1"
         LabelYFormat    =   ""
      End
   End
End
Attribute VB_Name = "GraphFrm_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'  6/10/99
'Const Filename = "c:\windows\temp\GraphLog.txt"


' === 12/11/02 ===
'Const Filename = "c:\temp\GraphLog.txt"
Const Filename = "c:\Steam_Sim\logs\GraphLog.txt"



'14/09/00 Dim LatestData(1 To NUMSET) As Single
'14/09/00 Dim GraphData(1 To NUMSET, 1 To NUMPOINT) As Single

Dim LatestData(NUMSET) As Single
Dim GraphData(NUMSET, 1 To NUMPOINT) As Single

Dim XPos As Integer
Dim GraphHalt As Boolean
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Sub PW(PrinterDialog As CommonDialog)


Dim pageheight, pageWidth, pagetop, pageleft As Single
Dim colourprint As Integer
Dim HeightRatio, GraphAspectratio As Single
Dim margin As Single
Dim i As Integer
Dim LastY As Integer
Dim MyStr As String

On Error GoTo printerCancel:

'----------------------------------
'Printer.PaperSize = vbPRPSA4
PrinterDialog.Orientation = cdlLandscape
'PrinterDialog. = cdlLandscape

PrinterDialog.Flags = cdlPDNoSelection Or cdlPDDisablePrintToFile
PrinterDialog.CancelError = True
        
'On Error GoTo printerCancel
PrinterDialog.ShowPrinter
      
Printer.ScaleMode = 5 ' Inches
Printer.Copies = PrinterDialog.Copies
Printer.Orientation = PrinterDialog.Orientation
Printer.PaperSize = vbPRPSA4

'Inialise printer
Printer.Print ""
    
'Determine if printer can handle colour
If GetDeviceCaps(Printer.hdc, BITSPIXEL) = 1 Then
  colourprint = graphMonochromeWithBorder
  '18/09/00 don't bother with thick lines, assume colour is OK
  'For i = 1 To 8 'make all lines thick if not in colour
    'Graph2.ThisSet = i
    'Graph2.ThickLines = 1
    'Graph2.PatternData = 2 ' set line thickness to 2
  'Next

Else
  colourprint = graphColorWithBorder
End If
Graph2.PrintStyle = colourprint

'----------------------------------


margin = 1 'Inch


pageheight = Printer.ScaleHeight - (2 * margin)
pageWidth = Printer.ScaleWidth - (2 * margin)
pagetop = Printer.ScaleTop + margin
pageleft = Printer.ScaleLeft + margin



Printer.Print

' Set up permanent printinfo settings
Graph2.PrintInfo(1) = Printer.hdc
Graph2.PrintInfo(6) = Printer.ScaleLeft
Graph2.PrintInfo(7) = Printer.ScaleTop
Graph2.PrintInfo(8) = Printer.ScaleWidth
Graph2.PrintInfo(9) = Printer.ScaleHeight

' Print a title in middle top  (nothing fancy)
Printer.CurrentY = pagetop
Printer.FontSize = 14
MyStr = "CMA Steam Simulator - Plot" & Space(10) & "Date:" & Space(1) & Format(Date, "mm/dd/yyyy") & Space(4) & "Time: " & Format(Time, "hh:mm:ss")
'Printer.CurrentX = pageleft + (pageWidth / 2) - (Len(MyStr) / 2)
Printer.CurrentX = pageleft + pageWidth / 8

Printer.Print MyStr






'Printer.CurrentX = pageWidth / 2
'Printer.Print "pageheight = " & Printer.ScaleHeight
'Printer.CurrentX = pageWidth / 2
'Printer.Print "pageWidth = " & Printer.ScaleWidth
'Printer.CurrentX = pageWidth / 2
'Printer.Print "pagetop = " & Printer.ScaleTop
'Printer.CurrentX = pageWidth / 2
'Printer.Print "pageleft = " & Printer.ScaleLeft
'Printer.CurrentX = pageWidth / 2
'Printer.Print "orientation = " & Printer.Orientation
'Printer.CurrentX = pageWidth / 2

'Printer.Print "Graph2.height " & Graph2.Height
'Printer.CurrentX = pageWidth / 2
'Printer.Print "Graph2.width = " & Graph2.Width

'HeightRatio = Me.ScaleHeight / Printer.ScaleHeight


'Printer.CurrentX = pageWidth / 2
'Printer.Print "Me.ScaleHeight = " & Me.ScaleHeight
'Printer.CurrentX = pageWidth / 2
'Printer.Print "HeightRatio = " & Format(str(HeightRatio), "###0.000")


pagetop = pagetop + 2 * (Printer.TextHeight("CMA Steam"))

Graph2.PrintInfo(2) = pageleft
Graph2.PrintInfo(3) = pagetop
Graph2.PrintInfo(4) = pageWidth

'what was original aspect ratio of Graph, W:H ?
GraphAspectratio = Graph2.Width / Graph2.Height

'So if new width is pagewidth, new height is ...

'Graph2.PrintInfo(5) = pageheight / 4

Graph2.PrintInfo(5) = pageWidth / GraphAspectratio



Graph2.DrawMode = graphPrint

' Now print the legend under the graph
Printer.CurrentX = pageleft
Printer.CurrentY = pagetop + (Printer.TextHeight("CMA Steam")) + Graph2.PrintInfo(5)
Printer.FontSize = 10

For i = 0 To 7
  Printer.CurrentX = pageleft
  'LastY = Printer.CurrentY
  Label6.Caption = str(i)
  Label6.ForeColor = Label2(i).ForeColor
  
  ' === 29/11/01 TESTED and works OK ===
 ' Printer.Print Label6.Caption & Space(2) & Label2(i).Caption
  Printer.ForeColor = Line1(i).BorderColor
  Printer.Print Label6.Caption & Space(2) & Label2(i).Caption
  
  
  'Printer.CurrentY = LastY
  'Label7.Caption = str(i + 4)
  'Label7.ForeColor = Label2(i).ForeColor
  'Printer.CurrentX = pageleft + pageWidth / 2
  'Printer.Print Label7.Caption & Space(2) & Label2(i + 4).Caption

  'Printer.Print Label6.Caption & Space(2) & Label2(i).Caption & Space(70 - Len(Label2(i).Caption)) & Label7.Caption & Space(2) & Label2(i + 4).Caption
  'Printer.CurrentY = Printer.CurrentY + Printer.TextHeight(Label6.Caption)
Next

Printer.Print
Printer.EndDoc

printerCancel:

End Sub
Public Sub PrintWindow(PrinterDialog As CommonDialog)
  Const margin = 1
  Dim colourprint As Long, hdcGrid As Long, PixelsX As Long, PixelsY As Long
  Dim GridWidth As Long, GridHeight As Long
  Dim WidthRatio As Double, HeightRatio As Double
  Dim GraphWidth As Single, GraphHeight As Single
    
    
    GoTo skipjf
    
  Printer.PaperSize = vbPRPSA4

  Printer.Print ""
  Printer.ScaleMode = 5
  'Printer.PaperSize = vbPRPSA4
  
  Graph1.PrintInfo(1) = Printer.hdc
   Graph1.PrintInfo(6) = Printer.ScaleLeft
   Graph1.PrintInfo(7) = Printer.ScaleTop
   Graph1.PrintInfo(8) = Printer.ScaleWidth
   Graph1.PrintInfo(9) = Printer.ScaleHeight
   Graph1.PrintInfo(2) = 1 'inch from left
   Graph1.PrintInfo(3) = 1 'inch from top
   Graph1.PrintInfo(4) = 3 'inch wide
   Graph1.PrintInfo(5) = 4 'inch high
  Graph1.DrawMode = 5
  Printer.EndDoc
  
'GoTo skip
skipjf:
  Printer.PaperSize = vbPRPSA4

  PrinterDialog.Flags = cdlPDNoSelection Or cdlPDDisablePrintToFile
  PrinterDialog.Orientation = cdlLandscape
  PrinterDialog.CancelError = True
        
  On Error GoTo printerCancel
  PrinterDialog.ShowPrinter
      
  Printer.ScaleMode = 5 ' Inches
  Printer.Orientation = PrinterDialog.Orientation 'vbPRORLandscape
  Printer.Copies = PrinterDialog.Copies
  'Inialise printer
  Printer.Print

  'Determine if printer can handle colour
  If GetDeviceCaps(Printer.hdc, BITSPIXEL) = 1 Then
    colourprint = graphMonochromeWithBorder
  Else
    colourprint = graphColorWithBorder
  End If
  
  HeightRatio = Me.ScaleHeight / Printer.ScaleHeight
  'In Inches
  '15/09/00 GraphHeight = Printer.ScaleHeight - (2 * MARGIN) - (dataGridPhasor.Height / HeightRatio)
  'GraphHeight = Printer.ScaleHeight - (2 * MARGIN) - (Graph1.Height / HeightRatio)
  'GraphWidth = (Printer.ScaleWidth - 2.05 * MARGIN) / 2
  
  GraphHeight = Graph1.Height
  GraphWidth = Graph1.Width

  
  'For i = 2 To 3
    With Graph1
      .PrintStyle = colourprint
    
      'Printer Canvas
      .PrintInfo(1) = Printer.hdc
    
      'Size of page
      .PrintInfo(6) = Printer.ScaleLeft
      .PrintInfo(7) = Printer.ScaleTop
      .PrintInfo(8) = Printer.ScaleWidth
      .PrintInfo(9) = Printer.ScaleHeight
    
      'Graph Offset
     ' If i = 2 Then
        .PrintInfo(2) = margin
    '  Else
    '    .PrintInfo(2) = Printer.ScaleWidth - MARGIN - GraphWidth
    '  End If
      
      .PrintInfo(3) = margin 'y
      'Size of graph
      .PrintInfo(4) = GraphWidth
      .PrintInfo(5) = GraphHeight
    
      'Print Graph
      .DrawMode = graphPrint
    End With
  'Next i
  
  'Print Grid
  Printer.ScaleMode = 1 ' Must Be in twips
  'Helper Function
  ' 15/09/00 put this grid bit in later
  'Call ImportFlexGrid(cTP, dataGridPhasor, Printer.ScaleWidth - 1428.84)
  
  'cTP.MarginBottom = 1428.84 '567 equals to 1 cm
  'cTP.MarginLeft = 1428.84
  'cTP.MarginTop = 0
  'Printer.CurrentY = Printer.ScaleHeight - 1428.84 - dataGridPhasor.Height
  
  'cTP.DrawTable Printer
  
  Printer.EndDoc

printerCancel:
Skip:
    Exit Sub
End Sub


Private Sub cmdExit_Click()
  Unload Me

End Sub

Private Sub cmdHalt_Click()
Debug.Print "Timer Interval = " & Timer1.Interval


Label4.Visible = True

GraphHalt = Not GraphHalt
If GraphHalt Then
  cmdHalt.Caption = "Resume"
  Label4.Caption = "RECORDER HALTED"
Else
  cmdHalt.Caption = "Halt"
  Label4.Caption = "RECORDER RUNNING"
End If



End Sub

Private Sub CmdLoad_Click()
Call ReadGraphData
End Sub

Private Sub cmdPrint_Click()

End Sub


Private Sub CmdPrint2_Click()
  'Call PrintWindow(CommonDialog1)
  
  On Error GoTo errh
  
  Dim i, j As Integer
  
  'First Transfer existing Real Time Plot data to a standard (non real-time) graph
  Graph2.Visible = True
  Graph1.ZOrder 'hide with original graph
  Graph2.GraphType = 6 ' 2d line
  
  Graph2.ThickLines = 0 ' thin lines
  
  Graph2.LeftTitle = "%"
  Graph2.BottomTitle = "Time"
  
  '18/09/00
  Graph2.ThickLines = 0     ' OFF is 1 pixel thick
  
  ' === 30/11/01 Graph printed the lines too thin so try this... ===
  ' === 30/11/01 FAR TOO thick a print so remove again....
  'Graph2.ThickLines = 1     ' ON is 3 pixels thick
  
  
  '18/09/00 Graph2.ThickLines = 1
  '18/09/00  Graph2.PatternData = 2 ' set line thickness to 2
  'Graph2.GridStyle = 1 'y-axis grid lines if you want
  
  Graph2.NUMPOINTS = 200
  Graph2.NumSets = 8

  
  Graph2.Color(1) = 0 'black
  Graph2.Color(2) = 1 'blue
  Graph2.Color(3) = 2 'green
  Graph2.Color(4) = 3 'cyan
  Graph2.Color(5) = 4 'red
  Graph2.Color(6) = 5 'magenta
  Graph2.Color(7) = 6 'yellow
  Graph2.Color(8) = 8 'grey

  
  
  'Graph2.RandomData = 0 ' off
  'Graph2.DataReset = 1 ' graph data
  
  'get block of data from history array
  
  
  For i = 1 To UBound(GraphData, 2)
      
      For j = 1 To UBound(GraphData, 1)
        Graph2.ThisSet = j
        '21/09/00 read data from actal graph1, not data array - this should cope with historic data as well
        'Graph2.data(201 - i) = GraphData(j, i)
        
        If HistoricDataSelected Then
          Graph1.ThisSet = j 'added 21/09/00
          Graph2.data(i) = Graph1.data(i)
        Else
          Graph2.data(201 - i) = GraphData(j, i)
        End If
      Next
  Next

  Graph2.DrawMode = 3 ' draw off screen then display
  
  
  Call PW(CommonDialog1)
  Graph2.Visible = False
  
errh:
  
End Sub

Private Sub CmdSave_Click()
  Call StoreData
  
End Sub

Private Sub CmdSetup_Click()
Dim msg As String
Dim i, j As Integer
Dim Rsp As Integer

    
    


  Rsp = 0
  
  
  If NotFirstGraph Then ' Say that graph will be reset (if there was a graph)
    msg = " The Graph Will Now Be Reset." & vbCrLf & vbTab & "Continue ?"
    Rsp = MsgBox(msg, vbYesNo + vbCritical, "Reset Graph")
  End If
  
  If Rsp = vbYes Or Not NotFirstGraph Then
  
     'reset colours. For some reason they change
    Graph1.Color(1) = 0 'black
    Graph1.Color(2) = 1 'blue
    Graph1.Color(3) = 2 'green
    Graph1.Color(4) = 3 'cyan
    Graph1.Color(5) = 4 'red
    Graph1.Color(6) = 5 'magenta
    Graph1.Color(7) = 6 'yellow
    Graph1.Color(8) = 8 'grey
  
  
     
     Label4.Visible = False
     HScroll1.Value = HScroll1.max 'reset scroll bar to point to first "page"

     XPos = 0
     NotFirstGraph = True

    'Reset the Historic Data
    HistoricDataSelected = False
    
    For i = 1 To UBound(HistoricData, 2)
      For j = 1 To UBound(HistoricData, 1)
        HistoricData(j, i) = 0#
      Next
    Next

    GraphRun = False
  
  
    cmdStop_Click
    'show set up form
    
    '11/09/00 clear legend description of each channel immediately Setup selected
    For i = 1 To NUMSET
      LegText(i) = " "
    Next
    
    '15/09/00 LText LegText, Graph1
    Graph1.GraphType = 22
    Graph1.GraphStyle = 1
    Graph1.DrawMode = 3
    '11/09/00
    
    
    ' 12/09/08 SetUp.Show vbModal
    frm_PlotVars.Show
    
    'set up graph legend
    LText LegText, Graph1
    'redraw graph
    Graph1.GraphType = 22
    Graph1.GraphStyle = 1
    Graph1.DrawMode = 3
    'create log file
    CreateLog Filename
    'set timer interval
    'Timer1.Interval = SetUp.TimerInterval
    GraphFrm.Label3.Caption = str(Timer1.Interval / 1000) & Space(1) & "sec"

    'enable start button
    cmdStart.Enabled = True
  End If
End Sub

Private Sub cmdStart_Click()
  'disable start button
  cmdStart.Enabled = False
  'enable stop button
  cmdStop.Enabled = True
  cmdHalt.Enabled = True
  HScroll1.Enabled = False
  HScroll1.Visible = False
  'enable timer
  Timer1.Enabled = True
  GraphRun = True
  Label4.Caption = "RECORDER RUNNING"

End Sub

Private Sub cmdStop_Click()
Dim temp As Integer
  'disable stop button
  cmdStop.Enabled = False
  cmdHalt.Enabled = False
  Timer1.Enabled = False
  'close log file
  CloseLog
  'set maxium num of point in log file
  temp = TotalPoints
  
  '14/09/00 If temp > NUMPOINT * NUMSET Then
    'HScroll1.max = temp - 1600
  'Else
    'HScroll1.max = 0
  'End If
  
  '11/09/00 this wipes graph clean Do we need it ? HScroll1_Change
  'HScroll1.Enabled = True
  GraphRun = False
  
  Label4.Caption = "RECORDER STOPPED"
  Label4.Visible = True

   If Check1.Enabled = False Then 'there is historic data available so show scroll bar
   
      HScroll1.Enabled = True
      HScroll1.Visible = True
   End If
   
End Sub



Private Sub Form_Load()
Dim i As Long
'set up graph
Me.Width = 16275
HavenStretch1.Width = 15960
  

  
 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

  
'Graph1.ThickLines = 1
'Graph1.PatternData = 5

  
Graph1.NUMPOINTS = NUMPOINT
Graph1.NumSets = NUMSET
Graph1.LeftTitle = "%"
  
Graph1.Color(1) = 0 'black
Graph1.Color(2) = 1 'blue
Graph1.Color(3) = 2 'green
Graph1.Color(4) = 3 'cyan
Graph1.Color(5) = 4 'red
Graph1.Color(6) = 5 'magenta
Graph1.Color(7) = 6 'yellow
Graph1.Color(8) = 8 'grey


' === 28/11/01 Did nothing so remove...===
'Graph1.AutoInc = 1

  
  'Graph1.TickEvery = 200
  'Graph1.LabelEvery = 200
  'set legend
'15/09/00 For i = 1 To NUMSET
  'Graph1.Legend(i) = " "
'Next i
  
'08/09/00
Label3.Caption = Format(str(Timer1.Interval / 1000), "##0.0") & Space(1) & "sec"


' === 14/05/03 TESTING (Copied from Diesel Sim ) ===

' === 16/05/03 ===
GoTo skip_this_test:

Graph1.Ticks = 1    ' graphOn
Graph1.Labels = 1   ' graphOn
Graph1.ForegroundUse = 4
Graph1.Foreground = 0 'set labels in black

Graph1.ForegroundUse = 6
Graph1.Foreground = 0 'set axes in black

Graph1.AutoInc = 0   ' OFF
Graph1.GraphStyle = 1   ' Continuous Line

Graph1.TickEvery = NUMPOINT / 10 'X axis - show every 10 th tick

'Graph1.LabelXFormat = "mm:ss"

Graph1.LabelEvery = NUMPOINT / 10 'X-axis label every 10th

'For i = 1 To 200
'  Graph1.Label(i) = ""
'Next i
'Graph1.Label(1) = "0:00"
'Graph1.Label(20) = "0:20"

'With Graph1
'    .LabelXType = 2
'    .LabelXDateStart = "0:0"
'    .LabelXDateInc = "0:15"
'    .LabelXFormat = "mm:ss"
'End With

skip_this_test:

' ==================================================
  
  
  'redraw graph
Graph1.DataReset = 1

'11/09/00
'Graph1.DrawMode = 3
NotFirstGraph = False
  
  
End Sub

Private Sub HScroll1_Change()
  Dim i As Integer
  Dim j As Integer
  Dim indx As Integer
  
  
  HistoricDataSelected = True
  
  'change graph type to 2d line
  Graph1.GraphType = 6 ' 2d line
  'Graph1.MissingData = graphZeroData
  'Graph1.MissingLineMode = graphOmit
  Graph1.ThickLines = 0 ' thin lines
  'Graph1.ThickLines = 1 ' thick lines
  Graph1.RandomData = 0 ' off
  Graph1.DataReset = 1 ' graph data
  
  '21/09/00 set up colours again
  Graph1.NUMPOINTS = NUMPOINT
  Graph1.NumSets = NUMSET
  Graph1.LeftTitle = "%"
  
  Graph1.Color(1) = 0 'black
  Graph1.Color(2) = 1 'blue
  Graph1.Color(3) = 2 'green
  Graph1.Color(4) = 3 'cyan
  Graph1.Color(5) = 4 'red
  Graph1.Color(6) = 5 'magenta
  Graph1.Color(7) = 6 'yellow
  Graph1.Color(8) = 8 'grey
  
  
  
  
  
  'get block of data from history array
  
  indx = HScroll1.max - HScroll1.Value
  
  For i = (indx * 200) + 1 To (indx + 1) * 200
      
      For j = 1 To NUMSET
        Graph1.ThisSet = j
        Graph1.data(201 - (i - (indx * 200))) = HistoricData(j, i)
      Next
  Next
  
  Graph1.DrawMode = 3 ' draw off screen then display
  
End Sub

Private Sub HScroll2_Change()
'testing only
  'A_OUTPUT(0) = HScroll2.Value
End Sub

Private Sub Timer1_Timer()
  'store latest data in point 1 and call draw mode 8
Dim i As Integer
Dim j As Integer

'08/09/00 do not plot during halt

If d_run And GraphRun And Not GraphHalt Then


  'shift all entries in historic array down one
  For i = UBound(HistoricData, 2) To LBound(HistoricData, 2) + 1 Step -1
    For j = LBound(HistoricData, 1) To UBound(HistoricData, 1)
      HistoricData(j, i) = HistoricData(j, i - 1)
    Next
  Next

  For i = UBound(GraphData, 2) To LBound(GraphData, 2) + 1 Step -1
    For j = LBound(GraphData, 1) To UBound(GraphData, 1)
      GraphData(j, i) = GraphData(j, i - 1)
    Next
  Next


  'read data from ao array
  GetData LatestData
  'loop through sets
  
  ' === 28/11/01 ===
  ' Plotting on top of a previous plot screws up the colours so do this....
  For i = 1 To NUMSET
    If LatestData(i) > 0.25 Then
        If i < NUMSET Then
            For j = i + 1 To NUMSET
                If Abs(LatestData(j) - LatestData(i)) < 0.25 Then
                    LatestData(j) = LatestData(j) + 0.5
                End If
            Next j
        End If
    Else
        LatestData(i) = 0.25
    End If
  Next i
  ' === 28/11/01 end of edit ===
  
  
  For i = 1 To NUMSET
  
    Graph1.ThisSet = i
    Graph1.data(1) = LatestData(i)
    HistoricData(i, 1) = LatestData(i)
    GraphData(i, 1) = LatestData(i)
    
    ' === 28/11/01 =====
    ' Can't seem to get this to work so comment out...
    '
    '' If we plot two graphs on TOP of each other then force the colour to BLACK
   '
   '     Graph1.ThisPoint = 1
   '    For j = 1 To 8
   '
   '         If Abs(plot_Reals(i) - plot_Reals(j)) < 5# Then
   '
   '                 Graph1.ThickLines = 1
   '
   '                 Graph1.ColorData = 0   ' Black
   '
   '               ' Using the next line plots NOTHING...
   '                 Graph1.Color(j) = 0   ' Black
   '
   '         Else
   '             ' === 28/11/01 ==
   '             If i = 1 Then
   '                   Graph1.ColorData = 14 ' Yellow
   '                '  Graph1.Color(i) = 14   ' Yellow
   '             ElseIf i = 8 Then
   '                   Graph1.ColorData = 8  'avoid white, use grey instead
   '                '  Graph1.Color(i) = 8   ' Grey
   '             Else
   '                   Graph1.ColorData = i - 1 '0,1,2, - 7)
   '                '  Graph1.Color(i) = i - 1
   '             End If
   '
   '         End If
   '
   '     Next j
   ' ==================

    
    
  Next
  '14/09/00 RotateArray GraphData, LatestData
  'RotateArray HistoricData, LatestData 'maintain a history 10 times longer
  
  
  
  'redraw graph
  Graph1.DrawMode = 8
  XPos = XPos + 1
  If Check1.Value And XPos = NUMPOINT Then
    cmdStop_Click
  End If
  
  'If XPos > NUMPOINT Then Check1.Enabled = False
  If XPos > NUMPOINT Then
      Check1.Enabled = False
  Else
      Check1.Enabled = True
  End If
  
  

 
    
End If

  'Debug.Print "graph1.thicklines = " & Graph1.ThickLines
  'Debug.Print "graph1.patternedlines" & Graph1.PatternedLines
  'Debug.Print "graph1.timer.interval = " & Timer1.Interval

  '08/09/00
  Label3.Caption = Format(str(Timer1.Interval / 1000), "##0.0") & Space(1) & "sec"
  Graph1.XAxisTicks = 20
  Call Label_X_Axis(SetUp.TimerInterval / 1000, Graph1)
End Sub

