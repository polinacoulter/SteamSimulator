VERSION 5.00
Begin VB.Form SetUp 
   Caption         =   "Graph Set up"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Setup.frx":0000
      Left            =   345
      List            =   "Setup.frx":0002
      TabIndex        =   7
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Series"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Sample Interval"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Variables"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Series"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "SetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SetNum As Integer
Dim num As Integer
Dim IO As Long
Dim Speed As Long



Private Sub Combo1_Click()
'selects series and enables ao list
'also enable clear button
  List2.Enabled = True
  If Sets(Combo1.ItemData(Combo1.ListIndex)) <> -1 Then
    Command4.Enabled = True
  Else
    Command4.Enabled = False
  End If
End Sub

Private Sub Combo2_Click()
'sets speed of timer
  Select Case Combo2.ListIndex
    Case 0
      'fast
      Speed = 500 'milli seconds
    Case 1
      'medium
      '08/09/00 Speed = 6000 'milli seconds
      Speed = 1000
    Case 2
      'slow
      '08/09/00 Speed = 20000 'milli seconds
      Speed = 2000
    Case 3
      'very slow
      '08/09/00 Speed = 50000 'milli seconds
      Speed = 5000
      
    '08/09/00
    Case 4
      Speed = 10000
    
    Case 5
       Speed = 20000
   
    Case 6
       Speed = 30000
    
    Case 6
       Speed = 60000
   
  End Select
  'StrSpeed = Combo2.List(Combo2.ListIndex) 'Put the text into string to display
     
  
End Sub

Private Sub Command1_Click()
'ok button
  Unload Me
End Sub

Private Sub Command4_Click()
' clear button
  Sets(Combo1.ItemData(Combo1.ListIndex)) = -1
  LegText(Combo1.ItemData(Combo1.ListIndex)) = " "
  OKTest
End Sub

Private Sub form_Load()
Dim i As Byte
Dim SQL As String

'initialize  sets array
  For i = 1 To NUMSET
    Sets(i) = -1
    LegText(i) = " "
  Next
  
'set up series
  With Combo1
    For i = 1 To 8
      .AddItem "Pen " & i
      .ItemData(i - 1) = CLng(i)
    Next
    .ListIndex = 0
  End With
  
'set up time base combo
'08/09/00 specify actual time base
  'With Combo2
  '  .AddItem "Fast"
  '  .AddItem "Medium"
  '  .AddItem "Slow"
  '  .AddItem "Very Slow"
  '  .ListIndex = 0
  'End With
  
  With Combo2
    .AddItem "0.5 sec"  'added 11/09/00
    .AddItem "1 sec"
    .AddItem "2 sec"
    .AddItem "5 sec"
    .AddItem "10 sec"
    .AddItem "20 sec"
    .AddItem "30 sec"
    .AddItem "1 min"

    'added 11/09/00.ListIndex = 0
    .ListIndex = 1 'default to 1 sec
    
  End With
  
  
  
  'connect to database
  SQL = "Select [i/o ref],description From AO"
  'fill list box with analogue outputs
  FillList DB, List2, SQL
  'set up filename
End Sub

Private Sub List2_Click()
' selection of analogue output
  Sets(Combo1.ItemData(Combo1.ListIndex)) = List2.ItemData(List2.ListIndex)
  LegText(Combo1.ItemData(Combo1.ListIndex)) = List2.List(List2.ListIndex)
  
  '11/09/00 Fill in text on main Form as we go along ....
    LText LegText, GraphFrm.Graph1
    GraphFrm.Graph1.GraphType = 22
    GraphFrm.Graph1.GraphStyle = 1
    GraphFrm.Graph1.DrawMode = 3
  
  
  
  'enable of clear button
  Command4.Enabled = True
  num = num + 1
  '30/11/99
  If Combo1.ListIndex <= 6 Then
  
  Combo1.ListIndex = Combo1.ListIndex + 1
  Else
  End If
  
  OKTest
End Sub

Private Sub OKTest()
  Command1.Enabled = num > 0
End Sub

Public Function TimerInterval() As Long
'returns the contents of the speed variable
  TimerInterval = Speed
End Function
