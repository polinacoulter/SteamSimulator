VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "GRAPHS32.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form Time_Graph 
   ClientHeight    =   12285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16875
   LinkTopic       =   "Form1"
   ScaleHeight     =   12285
   ScaleWidth      =   16875
   StartUpPosition =   3  'Windows Default
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   11415
      Left            =   60
      Top             =   60
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   20135
      Begin VB.CommandButton CmdSetup 
         Caption         =   "Set Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         TabIndex        =   43
         Top             =   6600
         Width           =   1275
      End
      Begin VB.Frame LegendFrame 
         Caption         =   "Legend"
         Height          =   6270
         Left            =   12780
         TabIndex        =   32
         Top             =   120
         Width           =   3495
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   8
            Left            =   960
            TabIndex        =   42
            Top             =   4920
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   7
            Left            =   960
            TabIndex        =   41
            Top             =   4305
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Index           =   6
            Left            =   960
            TabIndex        =   40
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   5
            Left            =   930
            TabIndex        =   39
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Index           =   4
            Left            =   960
            TabIndex        =   38
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   3
            Left            =   960
            TabIndex        =   37
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   2
            Left            =   975
            TabIndex        =   36
            Top             =   1320
            Width           =   2400
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   35
            Top             =   720
            Width           =   2415
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
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   5
            Index           =   6
            X1              =   240
            X2              =   720
            Y1              =   4440
            Y2              =   4440
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
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   4
            X1              =   240
            X2              =   720
            Y1              =   3255
            Y2              =   3255
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
            BorderColor     =   &H0000FF00&
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
            X1              =   225
            X2              =   705
            Y1              =   1365
            Y2              =   1365
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
         Begin VB.Label Label2 
            Caption         =   "Label6"
            Height          =   345
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   5550
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            Height          =   360
            Left            =   2520
            TabIndex        =   33
            Top             =   4920
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.CommandButton Start_command 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7500
         Width           =   1275
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   10800
      End
      Begin VB.CommandButton Stop_command 
         BackColor       =   &H008080FF&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   8220
         Width           =   1275
      End
      Begin VB.TextBox Title_Text 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   1620
         TabIndex        =   25
         Text            =   "Title"
         Top             =   6960
         Width           =   2835
      End
      Begin VB.Frame Frame1 
         Height          =   1395
         Left            =   1620
         TabIndex        =   17
         Top             =   7560
         Width           =   2715
         Begin VB.OptionButton Stop_plot_option 
            Caption         =   "10"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton Stop_plot_option 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   540
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Stop_plot_option 
            Caption         =   "100"
            Height          =   255
            Index           =   2
            Left            =   1380
            TabIndex        =   19
            Top             =   540
            Width           =   855
         End
         Begin VB.OptionButton Stop_plot_option 
            Caption         =   "1000"
            Height          =   255
            Index           =   3
            Left            =   1380
            TabIndex        =   18
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Number of Plots"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2475
         Left            =   4620
         TabIndex        =   6
         Top             =   6480
         Width           =   4875
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "1 minute"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   20
            Left            =   180
            TabIndex        =   15
            Top             =   600
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "256 minutes ( 4 hours )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   2400
            TabIndex        =   14
            Top             =   2040
            Width           =   2355
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "128 minutes ( 2 hours )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2400
            TabIndex        =   13
            Top             =   1680
            Width           =   2415
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "64 minutes   ( 1 hour )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2400
            TabIndex        =   12
            Top             =   1320
            Width           =   2355
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "32 minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2400
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "16 minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   10
            Top             =   2040
            Width           =   1995
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "8 minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   9
            Top             =   1680
            Width           =   1995
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "4minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   8
            Top             =   1320
            Width           =   1995
         End
         Begin VB.OptionButton Plot_Time_option 
            Caption         =   "2 minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Plot Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   900
            TabIndex        =   16
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.Frame History_Frame 
         Height          =   2415
         Left            =   9600
         TabIndex        =   1
         Top             =   6540
         Width           =   6435
         Begin VB.FileListBox File_List 
            BackColor       =   &H00C0C0C0&
            Height          =   1455
            Left            =   3420
            TabIndex        =   3
            Top             =   600
            Width           =   2775
         End
         Begin VB.ListBox Date_List 
            BackColor       =   &H00C0C0C0&
            Height          =   1425
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   5
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Date_label 
            Alignment       =   2  'Center
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4200
            TabIndex        =   4
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.DirListBox dirList 
         Height          =   1665
         Left            =   10020
         TabIndex        =   0
         Top             =   7560
         Visible         =   0   'False
         Width           =   3135
      End
      Begin GraphsLib.Graph Graph1 
         Height          =   6315
         Left            =   60
         TabIndex        =   29
         Top             =   120
         Width           =   12675
         _Version        =   327680
         _ExtentX        =   22357
         _ExtentY        =   11139
         _StockProps     =   96
         BorderStyle     =   1
         Background      =   "15~-1~-1~-1~-1~-1~-1"
         BottomTitle     =   "Time"
         LeftTitle       =   "%"
         ThickLines      =   0
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3795
         Left            =   13260
         Max             =   100
         TabIndex        =   27
         Top             =   5520
         Value           =   100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   3795
         Left            =   13860
         Max             =   100
         TabIndex        =   24
         Top             =   5700
         Value           =   100
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   3795
         Left            =   14340
         Max             =   100
         TabIndex        =   23
         Top             =   5580
         Value           =   100
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Plot Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Top             =   6660
         Width           =   2715
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   75
         Left            =   6660
         TabIndex        =   30
         Top             =   11760
         Width           =   15
      End
   End
End
Attribute VB_Name = "Time_Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim File_name_A As String
Dim File_name_B As String
Dim File_name_C As String

Dim Date_String As String

Dim Directory_Name As String

' === 22/05/03 ===
Dim N_pens_used As Integer
Dim N_plot_current As Integer
Dim N_plots_MAX As Integer

Dim Plotter_is_running As Boolean

Dim Data_actual(480, 8) As Integer


Dim N_Data_points As Integer

Dim i As Integer
Dim j As Integer

Dim New_data(8) As Integer
'
Dim Time_Now As Date
Dim Time_Graph_Starts As Date

'
Dim One_minute As Date
Dim Two_minutes As Date
Dim Three_minutes As Date
Dim Four_minutes As Date
Dim Five_minutes As Date
Dim Six_minutes As Date
Dim Seven_minutes As Date
Dim Eight_minutes As Date
'
Dim Timer_Count As Integer
'
'
'
' ===== RWJ Found useful code from: ====
' WEb site graphicsserver.com
' then Knowledge Base then
' look for article:  GSKB154
'  ( Important info about Daylight Saving time )
' ====================================
''
'



Private Sub Command1_Click()

End Sub

Private Sub CmdSetup_Click()

' === 27/05/03 ===

     'reset colours. For some reason they change
    Graph1.Color(1) = 0 'black
    Graph1.Color(2) = 1 'blue
    Graph1.Color(3) = 2 'green
    Graph1.Color(4) = 3 'cyan
    Graph1.Color(5) = 4 'red
    Graph1.Color(6) = 5 'magenta
    Graph1.Color(7) = 6 'yellow
    Graph1.Color(8) = 8 'grey

    frm_PlotVars.Show

End Sub

Private Sub Date_List_Click()

Date_label.Caption = Date_List.text

Directory_Name = Date_List.text

File_List.Path = "c:\Haven\Plots" & "\" & Directory_Name

File_List.Refresh

End Sub

Private Sub Form_Load()

' Default values
    N_Data_points = 120     ' 2 Minutes
    Timer1.Interval = 1000

' ===22/05/03 ===
N_pens_used = 3

N_plots_MAX = 1
N_plot_current = 0

Plotter_is_running = False
Call Init_Plotter(Plotter_is_running)

' ====================================
Graph1.NUMPOINTS = N_Data_points
Graph1.NumSets = 8
Graph1.LeftTitle = "%"
  
Graph1.Color(1) = 0 'black
Graph1.Color(2) = 1 'blue
Graph1.Color(3) = 2 'green
Graph1.Color(4) = 3 'cyan
Graph1.Color(5) = 4 'red
Graph1.Color(6) = 5 'magenta
Graph1.Color(7) = 6 'yellow
Graph1.Color(8) = 8 'grey

' ====================================
    Graph1.NUMPOINTS = N_Data_points
    
    Graph1.GraphType = 6    ' GraphLine
    Graph1.LabelXType = 0
    Graph1.LabelStyle = 1
    
    If N_Data_points = 480 Then
        Graph1.TickEvery = N_Data_points / 16
    Else
        Graph1.TickEvery = N_Data_points / 12
    End If
        
' === 22/05/03 ===
For j = 1 To N_pens_used
    For i = 1 To N_Data_points
        Data_actual(i, j) = 0
        Graph1.ThisSet = j
        Graph1.data(i) = Data_actual(i, j)
    Next i
Next j

Graph1.YAxisStyle = 2   ' User defined
Graph1.YAxisMax = 100   ' 100 %

Graph1.BottomTitle = Date
           
Graph1.DrawMode = 2
'======================
Call Update_the_graph
'======================

' ====================================

One_minute = CDate("0:1:0")
Two_minutes = CDate("0:2:0")
Three_minutes = CDate("0:3:0")
Four_minutes = CDate("0:4:0")
Five_minutes = CDate("0:5:0")
Six_minutes = CDate("0:6:0")
Seven_minutes = CDate("0:7:0")
Eight_minutes = CDate("0:8:0")




' === 23/05/03===
Directory_Name = Format(Date, "Medium Date")
Call Make_Directory(Directory_Name)

File_List.Path = "c:\Haven\Plots" & "\" & Directory_Name



dirList.Path = "c:\Haven\Plots"

Date_List.Clear
i = 0
Do While dirList.List(i) <> ""
    Date_String = dirList.List(i)
    Date_String = Mid$(Date_String, 16, Len(Date_String))
    Date_List.AddItem Date_String
    i = i + 1
Loop


End Sub



Private Sub File_List_Click()

Dim Plot_Title_Text As String
Dim Plot_Date As String

If Plotter_is_running Then
    Exit Sub
End If


'File_name_A = "Plot_" & Date & "_" & Time_Now
'File_name_B = Replace(File_name_A, "/", "_")
'File_name_C = Replace(File_name_B, ":", "_")

File_name_C = File_List.Filename


Open "c:\Haven\Plots\" & Directory_Name & "\" & File_name_C For Input As #77



Input #77, Plot_Title_Text
Input #77, Plot_Date

' === 20/05/03 ===
Input #77, N_Data_points
Input #77, N_pens_used

Input #77, i
Plot_Time_option(i).Value = True

Call Initialise_the_Graph

           
' === 22/05/03 ===
For j = 1 To N_pens_used
    For i = 1 To N_Data_points
            Input #77, Data_actual(i, j)
    Next i
Next j
' === 22/05/03 Clear down un-used plots ===
For j = N_pens_used + 1 To 8
    For i = 1 To N_Data_points
            Data_actual(i, j) = 0
            Graph1.ThisSet = j
            Graph1.data(i) = Data_actual(i, j)
    Next i
Next j
        

Input #77, Time_Graph_Starts
               
Close #77

Graph1.BottomTitle = Plot_Title_Text & "  " & Plot_Date


'======================
Call Update_the_graph
'======================



End Sub


Private Sub Option1_Click()

End Sub

Private Sub Start_command_Click()

Start_command.BackColor = vbGreen
Stop_command.BackColor = &HC0C0C0   ' Grey

Plotter_is_running = True
Call Init_Plotter(Plotter_is_running)
    
    
If Plot_Time_option(20).Value Then

    N_Data_points = 120         ' 1 Minute
    Timer1.Interval = 500
    
ElseIf Plot_Time_option(0).Value Then
    N_Data_points = 120         ' 2 Minutes
    Timer1.Interval = 1000
ElseIf Plot_Time_option(1).Value Then
    N_Data_points = 240         ' 4 Minutes
    Timer1.Interval = 1000
ElseIf Plot_Time_option(2).Value Then
    N_Data_points = 480         ' 8 Minutes
    Timer1.Interval = 1000
ElseIf Plot_Time_option(3).Value Then
    N_Data_points = 480         ' 16 Minutes
    Timer1.Interval = 2000
ElseIf Plot_Time_option(4).Value Then
    N_Data_points = 480         ' 32 Minutes
    Timer1.Interval = 4000
ElseIf Plot_Time_option(5).Value Then
    N_Data_points = 480         ' 64 Minutes
    Timer1.Interval = 8000
ElseIf Plot_Time_option(6).Value Then
    N_Data_points = 480         ' 128 Minutes
    Timer1.Interval = 16000
ElseIf Plot_Time_option(7).Value Then
    N_Data_points = 480         ' 128 Minutes
    Timer1.Interval = 32000
End If
   
    Graph1.NUMPOINTS = N_Data_points
    
    Graph1.GraphType = 6    ' GraphLine
    Graph1.LabelXType = 0
    Graph1.LabelStyle = 1
    
    If N_Data_points = 480 Then
        Graph1.TickEvery = N_Data_points / 16
    Else
        Graph1.TickEvery = N_Data_points / 12
    End If
        
For j = 1 To 8
    For i = 1 To N_Data_points
        Data_actual(i, j) = 0
        Graph1.ThisSet = j
        Graph1.data(i) = Data_actual(i, j)
    Next i
Next j

Graph1.YAxisStyle = 2   ' User defined
Graph1.YAxisMax = 100   ' 100 %


Graph1.BottomTitle = Format(Date, "Medium Date")
           
Graph1.DrawMode = 2
    
    
File_name_A = Title_Text.text & " " & Format(Date, "Medium Date") & " " & Time_Now

File_name_B = Replace(File_name_A, "/", " ")
File_name_C = Replace(File_name_B, ":", " ")


'======================
Call Update_the_graph
'======================

Timer1.Enabled = True

End Sub

Private Sub Stop_command_Click()


Plotter_is_running = False
Call Init_Plotter(Plotter_is_running)


Start_command.BackColor = &HC0C0C0       ' Grey
Stop_command.BackColor = &H8080FF       'Redish

Call Record_History

Timer1.Enabled = False

End Sub

Private Sub Stop_plot_option_Click(Index As Integer)


Select Case Index

    Case 0
        N_plots_MAX = 1
    Case 1
        N_plots_MAX = 10
    Case 2
        N_plots_MAX = 100
    Case 3
        N_plots_MAX = 1000
        
End Select

End Sub

Private Sub Timer1_Timer()


' === Time Series Plotting ===
' Shift the data in the Data_actual array
For j = 1 To N_pens_used

    For i = 1 To N_Data_points - 1
        Data_actual(i, j) = Data_actual(i + 1, j)
    Next i

    Data_actual(N_Data_points, j) = New_data(j)

Next j

Time_Now = Time
Time_Graph_Starts = Time_Now

'======================
Call Update_the_graph
'======================

' We need an historical record of the current plot INCLUDING the times...

Timer_Count = Timer_Count + 1

If Timer_Count = N_Data_points Then

   Timer_Count = 0
   
    Call Record_History
            
   N_plot_current = N_plot_current + 1
            
   If N_plot_current = N_plots_MAX Then
   
            Timer1.Enabled = False
            
            N_plot_current = 0
            
            Plotter_is_running = False
            Call Init_Plotter(Plotter_is_running)
          
            Start_command.BackColor = &HC0C0C0       ' Grey
            Stop_command.BackColor = &H8080FF        'Redish
    End If
    
       
End If


End Sub

Private Sub VScroll1_Change()

New_data(1) = VScroll1.max - VScroll1.Value

End Sub
Sub Update_the_graph()

Dim Time_multiplier As Integer


' Copy the data to the graph for plotting

For j = 1 To N_pens_used

    For i = 1 To N_Data_points
        Graph1.ThisSet = j
        Graph1.data(i) = Data_actual(i, j)
    Next i
    
Next j

If N_Data_points = 120 Then                 ' 2 Minutes
    
        
    If Plot_Time_option(20).Value Then       '  1 Minute
        Graph1.Label(120) = Format(Time_Graph_Starts, "hh:mm:ss")
        Graph1.Label(60) = ""
        Graph1.Label(1) = Format(Time_Graph_Starts - One_minute, "hh:mm:ss")
    Else
        Graph1.Label(120) = Format(Time_Graph_Starts, "hh:mm:ss")
        Graph1.Label(60) = Format(Time_Graph_Starts - One_minute, "hh:mm:ss")
        Graph1.Label(1) = Format(Time_Graph_Starts - Two_minutes, "hh:mm:ss")
    End If
    
ElseIf N_Data_points = 240 Then             ' 4 Minutes
    
    Graph1.Label(240) = Format(Time_Graph_Starts, "hh:mm:ss")
    Graph1.Label(180) = Format(Time_Graph_Starts - One_minute, "hh:mm:ss")
    Graph1.Label(120) = Format(Time_Graph_Starts - Two_minutes, "hh:mm:ss")
    Graph1.Label(60) = Format(Time_Graph_Starts - Three_minutes, "hh:mm:ss")
    Graph1.Label(1) = Format(Time_Graph_Starts - Four_minutes, "hh:mm:ss")

ElseIf N_Data_points = 480 Then

            Time_multiplier = 1
            
    If Plot_Time_option(2).Value Then       '  8 Minutes
            Time_multiplier = 1
    ElseIf Plot_Time_option(3).Value Then   ' 16 Minutes
            Time_multiplier = 2
    ElseIf Plot_Time_option(4).Value Then   ' 32 Minutes
            Time_multiplier = 4
    ElseIf Plot_Time_option(5).Value Then   ' 64 Minutes
            Time_multiplier = 8
    ElseIf Plot_Time_option(6).Value Then   ' 128 Minutes
            Time_multiplier = 16
    ElseIf Plot_Time_option(7).Value Then   ' 128 Minutes
            Time_multiplier = 32
    End If
                
        Graph1.Label(480) = Format(Time_Graph_Starts, "hh:mm:ss")
        Graph1.Label(420) = Format(Time_Graph_Starts - One_minute * Time_multiplier, "hh:mm:ss")
        Graph1.Label(360) = Format(Time_Graph_Starts - Two_minutes * Time_multiplier, "hh:mm:ss")
        Graph1.Label(300) = Format(Time_Graph_Starts - Three_minutes * Time_multiplier, "hh:mm:ss")
        Graph1.Label(240) = Format(Time_Graph_Starts - Four_minutes * Time_multiplier, "hh:mm:ss")
        Graph1.Label(180) = Format(Time_Graph_Starts - Five_minutes * Time_multiplier, "hh:mm:ss")
        Graph1.Label(120) = Format(Time_Graph_Starts - Six_minutes * Time_multiplier, "hh:mm:ss")
        Graph1.Label(60) = Format(Time_Graph_Starts - Seven_minutes * Time_multiplier, "hh:mm:ss")
        Graph1.Label(1) = Format(Time_Graph_Starts - Eight_minutes * Time_multiplier, "hh:mm:ss")
End If

Graph1.DrawMode = 2


End Sub
Sub Record_History()
Dim Date_str As String

File_name_A = Title_Text.text & " " & Time_Now
File_name_B = Replace(File_name_A, "/", " ")
File_name_C = Replace(File_name_B, ":", " ")
        
On Error GoTo errh


Directory_Name = Format(Date, "Medium Date")
Call Make_Directory(Directory_Name)


Open "c:\Haven\Plots\" & Directory_Name & "\" & File_name_C For Output As #77


Write #77, Title_Text.text

Date_str = Format(Date, "Medium Date")

Write #77, Date_str

Write #77, N_Data_points
Write #77, N_pens_used

For i = 0 To 7
    If Plot_Time_option(i).Value = True Then
        Write #77, i
    End If
Next i
i = 20
If Plot_Time_option(20).Value = True Then
        Write #77, i
End If

            
For j = 1 To N_pens_used

    For i = 1 To N_Data_points
        Write #77, Data_actual(i, j)
    Next i
       
Next j

Write #77, Time_Now
               
Close #77

File_List.Refresh


Exit Sub

errh:

MsgBox "Bad characters in the Title," & vbCrLf & "please retry with a better Title.", vbExclamation, "File Error"


End Sub
Sub Initialise_the_Graph()

    Graph1.NUMPOINTS = N_Data_points
    
    Graph1.GraphType = 6    ' GraphLine
    Graph1.LabelXType = 0
    Graph1.LabelStyle = 1
    
    If N_Data_points = 480 Then
        Graph1.TickEvery = N_Data_points / 16
    Else
        Graph1.TickEvery = N_Data_points / 12
    End If

End Sub
Sub Init_Plotter(Plotter_Run_flag As Boolean)

Plotter_is_running = Plotter_Run_flag

For i = 0 To 7
    Plot_Time_option(i).Enabled = Not Plotter_Run_flag
Next i
For i = 0 To 3
    Stop_plot_option(i).Enabled = Not Plotter_Run_flag
Next i

    Plot_Time_option(20).Enabled = Not Plotter_Run_flag
    Title_Text.Enabled = Not Plotter_Run_flag
    History_Frame.Visible = Not Plotter_Run_flag
    
End Sub
Sub Make_Directory(Dir_name As String)

On Error GoTo dir_Err

MkDir "c:\Haven\Plots\" & Dir_name

Exit Sub

dir_Err:



End Sub


Private Sub VScroll2_Change()

New_data(2) = VScroll2.max - VScroll2.Value

End Sub

Private Sub VScroll3_Change()

New_data(3) = VScroll3.max - VScroll3.Value

End Sub
