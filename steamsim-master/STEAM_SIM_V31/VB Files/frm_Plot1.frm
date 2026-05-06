VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "Graphs32.ocx"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form frm_Plot1 
   Caption         =   "Time series graph"
   ClientHeight    =   8385
   ClientLeft      =   2520
   ClientTop       =   1665
   ClientWidth     =   18510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8385
   ScaleWidth      =   18510
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8295
      Left            =   120
      Top             =   120
      Width           =   17655
      _ExtentX        =   31141
      _ExtentY        =   14631
      Begin VB.Frame LegendFrame 
         Caption         =   "Legend"
         Height          =   6375
         Left            =   14400
         TabIndex        =   7
         Top             =   0
         Width           =   3255
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   7
            Left            =   960
            TabIndex        =   15
            Top             =   4920
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   6
            Left            =   960
            TabIndex        =   14
            Top             =   4320
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   5
            Left            =   960
            TabIndex        =   13
            Top             =   3720
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   4
            Left            =   960
            TabIndex        =   12
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   3
            Left            =   960
            TabIndex        =   11
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   2
            Left            =   960
            TabIndex        =   10
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   9
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   8
            Top             =   720
            Width           =   1935
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
         Begin VB.Line Line1 
            BorderColor     =   &H00000080&
            BorderWidth     =   5
            Index           =   6
            X1              =   240
            X2              =   720
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFF00&
            BorderWidth     =   5
            Index           =   5
            X1              =   240
            X2              =   720
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   4
            X1              =   240
            X2              =   720
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0FFFF&
            BorderWidth     =   5
            Index           =   3
            X1              =   240
            X2              =   720
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000C000&
            BorderWidth     =   5
            Index           =   2
            X1              =   240
            X2              =   720
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   5
            Index           =   1
            X1              =   240
            X2              =   720
            Y1              =   1440
            Y2              =   1440
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
      End
      Begin GraphsLib.Graph Graph1 
         Height          =   2175
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   2415
         _Version        =   327680
         _ExtentX        =   4260
         _ExtentY        =   3836
         _StockProps     =   96
         BorderStyle     =   1
         PrintStyle      =   3
      End
      Begin VB.Frame ClearText 
         Caption         =   "Plot Controls"
         Height          =   1575
         Left            =   0
         TabIndex        =   0
         Top             =   6480
         Width           =   17655
         Begin VB.Timer Timer1 
            Left            =   15960
            Top             =   480
         End
         Begin VB.CommandButton PrintGraph 
            Caption         =   "Print Graph"
            Height          =   495
            Left            =   11880
            TabIndex        =   6
            Top             =   495
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2400
            TabIndex        =   3
            Text            =   "1 div/s"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Exit 
            Caption         =   "Exit"
            Height          =   495
            Left            =   13800
            TabIndex        =   2
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Configure 
            Caption         =   "Configure Channels"
            Height          =   495
            Left            =   375
            TabIndex        =   1
            Top             =   360
            Width           =   1695
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   16440
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Caption         =   "Chart Speed"
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Shape Shape1 
         Height          =   6375
         Left            =   0
         Top             =   0
         Width           =   14295
      End
   End
End
Attribute VB_Name = "frm_Plot1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const Interval = 4   'Sets sine Interval, was 4
Const Amplitude = 100 'Sets amplitude
Const XLen = 120       'Sets number of points visible

Dim Seconds As Double

Private Sub Combo1_Change()
    If Val(Combo1.text) > 0.01 And Val(Combo1.text) <= 10 Then
        Timer1.Interval = 100 / Val(Combo1.text)
        Debug.Print "Val(Combo1.Text) = " & Val(Combo1.text)
    End If
End Sub

Private Sub Combo1_Click()
    If Val(Combo1.text) > 0.01 And Val(Combo1.text) <= 10 Then
        Timer1.Interval = 100 / Val(Combo1.text)
        Debug.Print "Val(Combo1.Text) = " & Val(Combo1.text)
    End If
End Sub


Private Sub Configure_Click()
    Load frm_PlotVars
End Sub

Private Sub Exit_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim NumOfPlots As Integer



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)



'show legend text for any existing, allocated channels
For i = 1 To 8
    frm_Plot1.Label2(i - 1) = LegendText(i)
Next




'Fill Combo Box with timing options
Combo1.AddItem "10 mm/s"
Combo1.AddItem "5 mm/s"
Combo1.AddItem "2 mm/s"
Combo1.AddItem "1 mm/s"
Combo1.AddItem "0.5 mm/s"
Combo1.AddItem "0.4 mm/s"
Combo1.AddItem "0.3 mm/s"
Combo1.AddItem "0.2 mm/s"
Combo1.AddItem "0.1 mm/s"
Combo1.AddItem "0.05 mm/s"
Combo1.AddItem "0.01 mm/s"
'Combo1.



Graph1.GraphType = 22  'Time series
Graph1.GraphStyle = 1   'continuous line, default(0) is scatter
'NumOfPlots = 2
NumOfPlots = 8

Graph1.NumSets = NumOfPlots
Graph1.NUMPOINTS = XLen 'Defines length of X axis
Graph1.AutoInc = 0     'Off

'Appearance
Graph1.BackgroundUse = 0 'Entire graphing window
Graph1.Palette = 0

'Graph1.ThickLines = 1 'enable various thicknesses
'Graph1.PatternData = 5 'actual thickness in pixels (1-5)

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
Graph1.GridStyle = GridStyle 'GridStyle '0 = no grid, 1= x only, 2= y only, 3 = both x and y lines in grid


Graph1.Background = 127  '0= Black 1=dk blue, 5=purp

' === 28/11/01 ==
'Graph1.Color(1) = 0 'black
Graph1.Color(1) = 14 'yellow

Graph1.Color(2) = 1 'blue
Graph1.Color(3) = 2 'green
Graph1.Color(4) = 3 'cyan
Graph1.Color(5) = 4 'red
Graph1.Color(6) = 5 'magenta
Graph1.Color(7) = 6 'brown
Graph1.Color(8) = 8 'grey
'Graph1.SymbolSize = 550  '25 v small, 75 thickish

'Graph1.SymbolSize = 75  '25 v small, 75 thickish

'Axes
Graph1.TickEvery = XLen / 10  'X axis - show every 10 th tick
Graph1.LabelEvery = XLen / 10 'X-axis label every 10th

'Graph1.TickEvery = XLen / 10  'X axis - show every 5th tick

Graph1.YAxisStyle = 2         'Y axis - user defined scale
Graph1.YAxisMax = Amplitude


'Graph1.YAxisMin = -Amplitude
Graph1.YAxisMin = 0
Graph1.YAxisTicks = 10 'Number of ticks on positive and negative arms

Graph1.LeftTitle = "Value %"
Graph1.LeftTitleStyle = 1  'up
Graph1.RightTitle = "Value %"
Graph1.RightTitleStyle = 2 'down

'Graph fills entire form
'Graph1.Top = Me.ScaleTop
'Graph1.Left = Me.ScaleLeft
'Graph1.Width = Me.ScaleWidth
'Graph1.Height = Me.ScaleHeight

'Graph fills entire shape1
Graph1.Top = Shape1.Top
Graph1.Left = Shape1.Left
Graph1.Width = Shape1.Width
Graph1.Height = Shape1.Height
Seconds = 0         'Initialize

'Graph1.DrawMode = 2 'Redraw
'Timer1.Interval = 1 'orig setting
'Timer1.Interval = 200
Timer1.Interval = 100


Graph1.AutoInc = 0        'off


'Graph1.DrawMode = 2


    
On Error Resume Next
Call JYAScreenInit(Me, id_plot1)
 

End Sub

Private Sub Form_Resize()


On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_plot1)
    HavenStretch1.Visible = True
End If


End Sub

Private Sub PrintGraph_Click()
Dim BeginPage, EndPage, NumCopies, i
Dim msg



' Set Cancel to True
CommonDialog1.CancelError = True

'only 1 page for graphs
CommonDialog1.FromPage = 1
CommonDialog1.ToPage = 1



On Error GoTo ErrHandler1
    
' Display the Print dialog box
CommonDialog1.ShowPrinter

' Get user-selected values from the dialog box
'BeginPage = CommonDialog1.FromPage
'EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
For i = 1 To NumCopies
    On Error GoTo errorhandler2  ' Set up error handler.
    PrintForm   ' Print form.
Next i


ErrHandler1:
    ' User pressed the Cancel button
Exit Sub

errorhandler2:
    msg = "The form can't be printed."
    MsgBox msg  ' Display message.
    Resume Next
Exit Sub

 
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim j As Integer

'If UnloadPlot Then
'    UnloadPlot = False
'    Unload Me
'End If

'Always load newest value into ThisPoint = 1, for each
'data set being modified, and then redisplay
Seconds = (Seconds + Interval) Mod 360

'For i = 1 To 8
   ' plot_reals(i) = (GSGetSin(Seconds + (30 * i)) * 100) + 200
'Next


For i = 1 To 8

    'Debug.Print "Plot_Reals(" & i & ") = " & plot_Reals(i)
    Graph1.ThisSet = i
    Graph1.ThisPoint = 1
    
         ' === 28/11/01 ==
         '   If i = 1 Then
         '       Graph1.ColorData = 14 ' Yellow
         '   ' If i <> 7 Then
         '   ElseIf i <> 8 Then
         '       Graph1.ColorData = i - 1 '0,1,2, - 7)
         '   Else
         '       Graph1.ColorData = 8 'avoid white, use grey instead
         '   End If
    
    ' === 28/11/01 =====
    ' If we plot two graphs on TOP of each other then force the colour to BLACK
    
    For j = 1 To 8
    
        If Abs(plot_Reals(i) - plot_Reals(j)) < 0.5 Then
                Graph1.ColorData = 0   ' Black
            
        Else
            ' === 28/11/01 ==
            If i = 1 Then
                Graph1.ColorData = 14 ' Yellow
            ElseIf i = 8 Then
                Graph1.ColorData = 8 'avoid white, use grey instead
            Else
                Graph1.ColorData = i - 1 '0,1,2, - 7)
            End If

            
        End If
        
    Next j
    ' ==================
    
    If PlotVar(i) <> 0 Then
        Graph1.GraphData = plot_Reals(i) * 100 / 250 'normalised
    Else
        'not allocated so don't plot
        Graph1.GraphData = 0 'invisible
    End If
Next



'Special DrawMode only for time-series
Graph1.DrawMode = 8
End Sub

