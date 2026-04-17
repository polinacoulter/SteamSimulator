Attribute VB_Name = "CustomTGraph"
Option Explicit
Private FILENUMBER As Integer
Private NUMPOINTS As Long
Global Xlabels(0 To NUMPOINT) As String
Global Labels(0 To NUMPOINT) As String
Global Timebase As Byte

Global NotFirstGraph As Boolean
Global GraphRun As Boolean
Global testdata As Integer
Global testcounter As Integer

Global HistoricData(8, 1 To 2000) As Single

Global HistoricDataSelected As Boolean


Type Record
  Series As Integer
  Value As Single
End Type

Public Sub Label_X_Axis(Timebase As Integer, chart As Graph)


  chart.XAxisStyle = 2
  
  chart.ThisPoint = NUMPOINT
  chart.LabelText = str(200)
  chart.ThisPoint = NUMPOINT / 2
  chart.LabelText = str(100)
  chart.ThisPoint = 1
  chart.LabelText = str(0)
 

 End Sub
 
Public Sub LabelxAxis(Index As Integer, chart As Graph, Label As Integer)

'11/09/00 JBF - This routine by AM is not called

'Dim label(NumPoints) As Variant
Dim i As Integer
  ' default values
  chart.TickEvery = 20
  chart.LabelEvery = 20
  
  'rotate labels
  For i = 0 To NUMPOINTS - 1
        Xlabels(i) = Xlabels(i + 1)
  Next
  Select Case Index
    Case 0  '0 to 200 sec
      Xlabels(200) = str(Label)
      'correct x axis label
      chart.BottomTitle = "Seconds"
    Case 1 '0 to 20 min
       Xlabels(NUMPOINTS) = Format(str(Label / 10), "#0.0")
      'correct x axis label
      chart.BottomTitle = "Minutes"
    Case 2 '0 to 2 hrs
      chart.TickEvery = 25
      chart.LabelEvery = 25
      Xlabels(NUMPOINTS) = Format((Label / 25) / 4, "#.00")
      'correct x axis label
      chart.BottomTitle = "Hours"
    Case 3 '0 to 20 hrs
      'assign labels to chart
      Xlabels(NUMPOINTS) = Format(str(Label / 10), "#0.0")
      'correct x axis label
      chart.BottomTitle = "Hours"
  End Select
  'assign labels to chart
  chart.LabelText = Xlabels
End Sub

Public Sub LText(text() As String, chart As Graph)
' use text as tet for legend
Dim i As Integer, Lower As Integer, Upper As Integer
  ' array bounds
  Lower = LBound(text)
  Upper = UBound(text)
  'place text on chart
  For i = Lower To Upper
    chart.Legend(i) = text(i)
  Next
End Sub

Public Sub RotateArray(data() As Single, Value() As Single)
' takes the data array for the graph and moves values
'left, dicarding value in at start of array and placing the new value
'at the end of array
'data array must be a two dimensional array
'with series in dimension 1 and points in dimension 2
Dim Lower As Integer, Upper As Integer, Index As Integer
Dim LSets As Integer, USets As Integer
Dim i As Integer
  'calculated dimension of array
  LSets = LBound(data, 1)
  USets = UBound(data, 1)
  Lower = LBound(data, 2) 'data points
  Upper = UBound(data, 2) ' data points

  'loop through array, shifting entries left one
  For i = LSets To USets
    Index = Lower
    '14/09/00 While Index < Upper
    While Index < Upper - 1
      data(i, Index) = data(i, Index + 1)
      Index = Index + 1
    Wend
  'add new value to data array
  data(i, Upper) = Value(i)
  'StoreData i, Value(i)
  Next
End Sub

Public Sub UpdateGraphT(data() As Single, chart As Graph)
'transfers data to graph and redraws graph.
'data array must be a two dimensional array
'with series in dimension 1 and points in dimension 2
Dim i As Integer
Dim Lower As Integer, Upper As Integer, Index As Integer
Dim LSets As Integer, USets As Integer

   'calculated dimension of array
  LSets = LBound(data, 1)
  USets = UBound(data, 1)
  Lower = LBound(data, 2) 'data points
  Upper = UBound(data, 2) ' data points

  'transfer data to graph
  For i = LSets To USets
    chart.ThisSet = i 'set series number
    For Index = Lower To Upper
      chart.ThisPoint = Index
      chart.GraphData = data(i, Index)
    Next
  Next
  'redraw Graph
  chart.DrawMode = 3
End Sub



Public Sub GetData(data() As Single)
'read data form ao array
Dim Upper As Integer
Dim i As Integer



    testcounter = testcounter + 1
    
    If testcounter <= 20 Then
         testdata = 20
    Else
    
      If testcounter <= 50 Then
        testdata = testdata + 1
      Else
        testdata = testdata - 1
      End If
      If testcounter = 80 Then
        testcounter = 20
        testdata = 20
      End If
    End If
    
    
  Upper = UBound(data)
  For i = 1 To Upper
  If Sets(i) > 0 Then
    'works for AM's local routine with A_OUTPUT but cannot be networked
    'data(i) = (A_OUTPUT(Sets(i)) / 255) * 100
    'collect data from model across network instead
    
    
        
    
    'TEST ONLY !!!!
    'plot_Reals(i) = 255 - (i * 255 / 8) 'to test colours at equal intevals of 12.5%
    'data(i) = testdata + (7 * i) 'makes sawtooth waves 'The Test Equatio
    
    
    data(i) = (plot_Reals(i) * 100) / 255 'normalised 'The Real Equation
  Else
    data(i) = 0
  End If
  Next
End Sub

Public Sub CreateLog(Filename As String)
Dim temp As Record
' open log file for writing
FILENUMBER = FreeFile
   Open Filename For Output As #FILENUMBER 'Len = Len(temp)
   'Open Filename For Random Access Write As #FILENUMBER Len = Len(temp)
  NUMPOINTS = 0
End Sub

Public Sub CloseLog()
'closes log file
  Close #FILENUMBER
End Sub

Public Sub OpenLog(Filename As String)
Dim Value As Record
'opens log file for writing
FILENUMBER = FreeFile
  '01/09/00 Open Filename For Input As #FILENUMBER 'Len = Len(Value)
   Open Filename For Output As #FILENUMBER 'Len = Len(Value)
 'Open Filename For Random Access Read As #FILENUMBER Len = Len(Value)
End Sub
'Public Sub StoreData(SetNum As Integer, data As Single)
Public Sub StoreData()
Dim Filename As String
Dim temp As Record
'stores data of graph in file, so graph data ia not lost
  'increase total
  'NUMPOINTS = NUMPOINTS + 1
  'temp.Series = CInt(SetNum)
  'temp.Value = data
  'write to file
  'Put #FILENUMBER, NUMPOINTS, temp
  'Write #FILENUMBER, SetNum; data
  'Debug.Print "Writing " & data & " into " & SetNum
  Dim i, j As Integer
  
   FILENUMBER = FreeFile
   Filename = "GraphTest.dat"
   Open Filename For Output As #FILENUMBER

  
  
   For i = 1 To UBound(HistoricData, 2)
      For j = 1 To UBound(HistoricData, 1)
        Write #FILENUMBER, HistoricData(j, i)
      Next
    Next
  
  Close #FILENUMBER
  
End Sub
'Public Sub ReadDataT(Position As Long, data() As Single)
'read in data from log file
Public Sub ReadGraphData()
Dim i As Long, j As Integer, num As Integer
Dim base As Long
Dim temp As Record
Dim Filename As String

  'i = Position ' start at first series
  'While i <= Position + 1599 And EOF(FILENUMBER) = False
   ' j = 1
    '  base = ((i \ 8) * 8) '+1
    'While j <= 8 And EOF(FILENUMBER) = False
      'read data
     ' If base <= NUMPOINT * 8 - 8 Then
      '  Get #FILENUMBER, (base + j), temp
      'Else
       ' temp.Value = 0
     ' End If
      'store data
      'data(j, (((i - Position) \ 8) + 1)) = temp.Value
      'j = j + 1
    'Wend
    'i = i + 8
  'Wend
  'num = -1
  'base = (Position \ 8) '+1
  'While num <> base And EOF(FILENUMBER) = False
  '  Input #FILENUMBER, temp.Series, temp.Value
  '  If temp.Series = 1 Then num = num + 1
  'Wend
  ''store data allready read
  'data(1, 1) = temp.Value
  'i = 1
  'num = NUMSET * NUMPOINT
  'While i < num And EOF(FILENUMBER) = False
  '  Input #FILENUMBER, temp.Series, temp.Value
  '  data(temp.Series, (i \ 8) + 1) = temp.Value
  '  i = i + 1
  'Wend
  Filename = "GraphTest.dat"

  FILENUMBER = FreeFile
   Open Filename For Input As #FILENUMBER

    'While num <> base And EOF(FILENUMBER) = False
  '  Input #FILENUMBER, temp.Series, temp.Value
  '  If temp.Series = 1 Then num = num + 1
  'Wend

  
   For i = 1 To UBound(HistoricData, 2)
      For j = 1 To UBound(HistoricData, 1)
        If EOF(FILENUMBER) = False Then
          Input #FILENUMBER, HistoricData(j, i)
        End If
      Next
    Next
  
  Close #FILENUMBER
  
End Sub

Public Function TotalPoints()
  TotalPoints = NUMPOINTS
End Function
