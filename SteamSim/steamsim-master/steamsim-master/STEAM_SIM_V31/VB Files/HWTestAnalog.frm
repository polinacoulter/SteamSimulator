VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form HWTestAnalog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analog Hardware Test"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   8205
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9495
      Left            =   60
      Top             =   60
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   16748
      Begin VB.Frame Frame2 
         Caption         =   "Analog Outputs"
         Height          =   2955
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   7575
         Begin VB.OptionButton Option1 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   255
         End
         Begin VB.ListBox AOTextList 
            ForeColor       =   &H00800000&
            Height          =   450
            Left            =   3960
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   1800
            Width           =   255
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   0
            Left            =   4680
            TabIndex        =   13
            Top             =   2400
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Enabled         =   -1  'True
         End
         Begin ComctlLib.Slider Slider1 
            Height          =   255
            Left            =   1800
            TabIndex        =   20
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327682
            LargeChange     =   10
            Max             =   255
            TickFrequency   =   16
         End
         Begin ComctlLib.Slider Slider2 
            Height          =   255
            Left            =   1800
            TabIndex        =   21
            Top             =   1920
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327682
            LargeChange     =   10
            Max             =   255
            TickFrequency   =   16
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   22
            Top             =   2400
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   2
            Left            =   4200
            TabIndex        =   23
            Top             =   2400
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   7
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "100 %"
            Height          =   255
            Index           =   6
            Left            =   3120
            TabIndex        =   33
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "100 %"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   32
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "0 %"
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   31
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "All Variable"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   30
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "All Zero"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   29
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "All HalfScale"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   28
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "All Full Scale"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Single Variable"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   26
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "0 %"
            Height          =   255
            Index           =   8
            Left            =   1680
            TabIndex        =   25
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   24
            Top             =   1920
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   310
         Left            =   960
         Top             =   2160
      End
      Begin VB.Frame Frame3 
         Caption         =   "Analog Inputs"
         Height          =   6375
         Left            =   0
         TabIndex        =   0
         Top             =   3000
         Width           =   7575
         Begin VB.ListBox AIlist 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   5580
            Left            =   240
            TabIndex        =   1
            Top             =   600
            Width           =   7095
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   5
            Left            =   5700
            TabIndex        =   11
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   4
            Left            =   5700
            TabIndex        =   10
            Top             =   900
            Width           =   615
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   3
            Left            =   2580
            TabIndex        =   9
            Top             =   1140
            Width           =   615
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   2
            Left            =   5700
            TabIndex        =   8
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   2580
            TabIndex        =   7
            Top             =   900
            Width           =   615
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   2580
            TabIndex        =   6
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Value"
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Channel"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Location"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Description"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "HWTestAnalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AnyVal_Click()
Dim i As Integer

HWTestOn = True
For i = 0 To 767
    AO(i) = Slider1.Value
Next
End Sub

Private Sub FSD_Click()
Dim i As Integer

'28/9/97 compile error - variable undeclared
'FSD.Value = 1
'HSD.Value = 0
'ZSD.Value = 0
HWTestOn = True
For i = 0 To 767
    AO(i) = 255
Next
'HWTestOn = False

End Sub

Private Sub HSD_Click()
Dim i As Integer

'28/9/97 compile error - variable undeclared

'FSD.Value = 0
'HSD.Value = 1
'ZSD.Value = 0
HWTestOn = True
For i = 0 To 767
    AO(i) = 127
Next
'HWTestOn = False
End Sub





Private Sub AOTextList_DblClick()
Dim i As Integer


' 31/10/97 MOVED CODE to slider2
'set all channels to zero except the one being tested
'If ResetAOChannels Then
'    For i = 0 To 767
'        If i <> AOTextList.ItemData(AOTextList.ListIndex) Then
'            AO(i) = 0
'        End If
'    Next
'End If

'set AO channel
'AO(AOTextList.ItemData(AOTextList.ListIndex)) = Slider2.Value
'Debug.Print "AO channel = " & AOTextList.ItemData(AOTextList.ListIndex)

End Sub

Private Sub Check1_Click()
'set flag
' 4/11/97 no check box any more ResetAOChannels = Check1.Value

End Sub

Private Sub Command1_Click()


End Sub

Private Sub Form_Load()
'if the list is not loaded, then load now
'add text and AO number in data field


' 17/11/97 TEST to see where the file is saved
' 17/11/97 TEST edited c:\windows\recent\HWtestanalog.frm

' ===== 14/11/97 =========================
Dim ii As Integer


 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


GoTo skipfornow

'6/2/98 need this later
'If useEnglish Then
'   Open App.Path & "\EAItest.txt" For Input As #1
'Else
'   Open App.Path & "\IAItest.txt" For Input As #1
'End If
'For ii = 1 To 31
'   Input #1, AItest_strings(ii)
'Next ii
'Close #1

'HWTestAnalog.Caption = AItest_strings(1)
'Frame2.Caption = AItest_strings(2)
'Label2(0).Caption = AItest_strings(3)
'Label2(1).Caption = AItest_strings(4)
'Label2(2).Caption = AItest_strings(5)
'Label2(3).Caption = AItest_strings(6)
'Label2(7).Caption = AItest_strings(7)
'Frame3.Caption = AItest_strings(8)
'Label3(6).Caption = AItest_strings(9)
'Label3(7).Caption = AItest_strings(9)
'Label3(0).Caption = AItest_strings(10)
'Label3(1).Caption = AItest_strings(11)
'Label3(3).Caption = AItest_strings(12)
'Label3(2).Caption = AItest_strings(13)
'Label3(4).Caption = AItest_strings(14)
'Label3(5).Caption = AItest_strings(15)

'===========================================

'31/10/97
ResetAOChannels = True
Slider1.Enabled = False
Slider2.Enabled = False
Label2(4).Enabled = False
Label2(5).Enabled = False
Label2(6).Enabled = False
Label2(8).Enabled = False
AOTextList.Visible = False


'6/2/98 through to end - will be needed later
'
'If AOTextList.NewIndex < 0 Then
'
'AOTextList.AddItem AItest_strings(16)
'AOTextList.ItemData(AOTextList.NewIndex) = 271 'incorrect
'
' AOTextList.AddItem AItest_strings(17)
'AOTextList.ItemData(AOTextList.NewIndex) = 261

' AOTextList.AddItem AItest_strings(18)
'AOTextList.ItemData(AOTextList.NewIndex) = 269

 'AOTextList.AddItem AItest_strings(19)
'AOTextList.ItemData(AOTextList.NewIndex) = 270

' AOTextList.AddItem AItest_strings(20)
'AOTextList.ItemData(AOTextList.NewIndex) = 267

 'AOTextList.AddItem AItest_strings(21)
'AOTextList.ItemData(AOTextList.NewIndex) = 268

 'AOTextList.AddItem AItest_strings(22)
'AOTextList.ItemData(AOTextList.NewIndex) = 258

 'AOTextList.AddItem AItest_strings(23)
'AOTextList.ItemData(AOTextList.NewIndex) = 257

 'AOTextList.AddItem AItest_strings(24)
'AOTextList.ItemData(AOTextList.NewIndex) = 260

 'AOTextList.AddItem AItest_strings(25)
'AOTextList.ItemData(AOTextList.NewIndex) = 259

 'AOTextList.AddItem AItest_strings(26)
'AOTextList.ItemData(AOTextList.NewIndex) = 284

 'AOTextList.AddItem AItest_strings(27)
'AOTextList.ItemData(AOTextList.NewIndex) = 262

 'AOTextList.AddItem AItest_strings(28)
'AOTextList.ItemData(AOTextList.NewIndex) = 264

 'AOTextList.AddItem AItest_strings(29)
'AOTextList.ItemData(AOTextList.NewIndex) = 263

 'AOTextList.AddItem AItest_strings(30)
'AOTextList.ItemData(AOTextList.NewIndex) = 266

 'AOTextList.AddItem AItest_strings(31)
'AOTextList.ItemData(AOTextList.NewIndex) = 265

'set List Index to first entry (default selection)
'AOTextList.ListIndex = 0
'End If
skipfornow:
'set flag for AO test
AOTestOn = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
AOTestOn = False
End Sub

Private Sub Slider1_Change()


'Label2(6).Caption = Format(Slider1.Value * 100 / 255, "##0") & " %"

End Sub

Private Sub Option1_Click(Index As Integer)
Dim i As Integer
Dim AOValue As Integer

'disable slider & text unless this option selected
Slider1.Enabled = False
Slider2.Enabled = False
Label2(4).Enabled = False
Label2(5).Enabled = False
Label2(6).Enabled = False
Label2(8).Enabled = False
AOTextList.Visible = False

'9/2/98
Label1.Visible = False
UpDown1(0).Visible = False
UpDown1(1).Visible = False
UpDown1(2).Visible = False


Select Case Index
Case 0
' 30/10/97 RWJ
' AO meters are 2.50 Volts fsd
' AO  chips are 2.55 Volts fsd
' We want a meter fsd and NOT a chip fsd so:
'   AOValue = 255
    AOValue = 250
   For i = 0 To 767
      AO(i) = AOValue
   Next
   
  
   
Case 1
' 30/10/97
'   AOValue = 127
    AOValue = 125
   For i = 0 To 767
      AO(i) = AOValue
   Next
Case 2
    AOValue = 0
   For i = 0 To 767
      AO(i) = AOValue
   Next
Case 3
    Slider1.Enabled = True
    Label2(4).Enabled = True
    Label2(5).Enabled = True
'    Label2(6).Enabled = True
    AOValue = Slider1.Value
    
' 31/10/97
  For i = 0 To 767
    AO(i) = AOValue
  Next

Case 4

' 19/2/98  Commented Out
'    AOTextList.Visible = True
    
    Slider2.Enabled = True
    Label2(6).Enabled = True
    Label2(8).Enabled = True
     AOValue = Slider2.Value
     
     Label1.Visible = True
     UpDown1(0).Visible = True
     UpDown1(1).Visible = True
     UpDown1(2).Visible = True
  
End Select

HWTestOn = True
AOTestOn = True
Call set_AO(AO(0))  'pass first array element
                    'set_AO then sets whole array




End Sub



Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
HWTestOn = True
For i = 0 To 767
    AO(i) = Slider1.Value
Next

' 6/2/98
Call set_AO(AO(0))  'pass first array element
                    'set_AO then sets whole array

End Sub

Private Sub Slider2_Change()
Dim i As Integer
Dim AOChIndx As Integer
Dim resp As String


'Label1(4).Caption = Format(Slider2.Value * 100 / 255, "##0") & " %"

'====31/10/97
HWTestOn = True
For i = 0 To 767
' 6/8/98
'        If i <> AOTextList.ItemData(AOTextList.ListIndex) Then
            AO(i) = 0
'        End If
Next

' 6/2/98
Call set_AO(AO(0))  'pass first array element
                    'set_AO then sets whole array

'set AO channel
' 6/2/98 Removed:
'AO(AOTextList.ItemData(AOTextList.ListIndex)) = Slider2.Value


End Sub

Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
HWTestOn = True
For i = 0 To 767
    AO(i) = 0
Next
AO(AO_Number) = Slider2.Value
' 6/2/98
Call set_AO(AO(0))  'pass first array element
                    'set_AO then sets whole array

End Sub

Private Sub Text1_Change()
'copy through any manual changes to value
' Race Problem can't do Slider2.Value = Val(Text1.Text)

End Sub

Private Sub Timer1_Timer()
Dim i As Integer

' 11/2/98
'Dim AIstrings(100) As String
'22/4/98
Dim AIstrings(0 To 100) As String
Dim AIStrLen As Integer

AIstrings(0) = "E.R. LOCAL MELO Temmperature Control"
AIstrings(1) = "E.R. LOCAL Charge Air Temperature Control"
AIstrings(2) = "E.R. LOCAL Pitch Control"
AIstrings(3) = "E.R. LOCAL RPM Control"
AIstrings(4) = "E.R. LOCAL Fuel Oil Pressure Control"
AIstrings(5) = "E.R. LOCAL HTCW Temperature Control"
AIstrings(6) = "E.R. LOCAL S.W. Temperature Control"
AIstrings(7) = "E.R. LOCAL L.T.C.W. Temperature Control"

AIstrings(24) = "MECD#5    REAL PID Output"
AIstrings(25) = "MECD#2    Pitch Control"
AIstrings(26) = "MECD#3    FO Viscosity Control"
AIstrings(27) = "MECD#2    Rev Limiter"
AIstrings(29) = "MECD#2    RPM Control "
AIstrings(48) = "BRIDGE    Rudder Angle Control"
AIstrings(49) = "BRIDGE    RPM Control"
AIstrings(50) = "BRIDGE    Winch Load Control"
AIstrings(51) = "BRIDGE    Pitch Control"
AIstrings(56) = "ELEC#3    D/A 1     AVR Control"
AIstrings(57) = "ELEC#3    Shaft Alt AVR Control"
AIstrings(58) = "ELEC#3    D/A 2     AVR Control"
AIstrings(59) = "ELEC#3    Synchroscope Phase Difference"


Call get_pots(AI(0))

' show latest AI for each as %
AIlist.clear
'For i = 24 To 31
'   AIlist.AddItem " AI(" & i & ")  =  " & Format(AI(i), "##0")
For i = 0 To 100
  If AIstrings(i) <> "" Then
  
'   AIlist.AddItem " AI(" & i & ")  =  " & Format(AI(i), "##0")

    AIlist.AddItem " AI(" & i & ")" & "  =  " & Format(AI(i), "000") & "  " & Mid$(AIstrings(i), 1, 40)

'    AIStrLen = Len(AIstrings(i))
' Crashed:
'    AIlist.AddItem " AI(" & Format(i), "00" & ")" & "  =  " & Format(AI(i) * 100 / 255, "000") & "% " & Mid$(AIstrings(i), 1, AIStrLen)
  End If
Next i
       
End Sub

Private Sub UpDown1_Change(Index As Integer)


    AO_Number = UpDown1(0).Value + UpDown1(1).Value * 10 + UpDown1(2).Value * 100
    Label1.Caption = Format(UpDown1(2).Value, "0") & " " & Format(UpDown1(1).Value, "0") & " " & Format(UpDown1(0).Value, "0")
    Dim i As Integer
    HWTestOn = True
    For i = 0 To 767
        AO(i) = 0
    Next
    AO(AO_Number) = Slider2.Value
' 6/2/98
    Call set_AO(AO(0))  'pass first array element
                    'set_AO then sets whole array
End Sub
