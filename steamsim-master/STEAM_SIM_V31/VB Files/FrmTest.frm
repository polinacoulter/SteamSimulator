VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form FrmHardwareTest 
   Caption         =   "Hardware Testing"
   ClientHeight    =   12885
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12885
   ScaleWidth      =   14925
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   12735
      Left            =   60
      Top             =   60
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   22463
      Begin VB.Frame TabPage 
         Height          =   5640
         Index           =   3
         Left            =   7440
         TabIndex        =   0
         Top             =   9300
         Visible         =   0   'False
         Width           =   7935
         Begin VB.OptionButton optAOOption 
            Caption         =   "All GaugesFull Scale"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   8
            Top             =   840
            Width           =   3735
         End
         Begin VB.OptionButton optAOOption 
            Caption         =   "All Gauges Half Scale"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   7
            Top             =   1290
            Width           =   3855
         End
         Begin VB.OptionButton optAOOption 
            Caption         =   "All Gauges Zero Scale"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   6
            Top             =   1740
            Value           =   -1  'True
            Width           =   4455
         End
         Begin VB.OptionButton optAOOption 
            Caption         =   "All Gauges"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   5
            Top             =   2190
            Width           =   1695
         End
         Begin VB.OptionButton optAOOption 
            Caption         =   "Single Gauge"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   3
            Top             =   3090
            Width           =   1695
         End
         Begin VB.TextBox TBGaugeNum 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2640
            TabIndex        =   2
            Top             =   3540
            Width           =   615
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   3240
            TabIndex        =   1
            Top             =   3540
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            OrigLeft        =   3360
            OrigTop         =   3480
            OrigRight       =   3600
            OrigBottom      =   3855
            Enabled         =   -1  'True
         End
         Begin ComctlLib.Slider SliderAll 
            Height          =   255
            Left            =   1080
            TabIndex        =   4
            Top             =   2640
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            _Version        =   327682
            Max             =   100
            TickFrequency   =   5
         End
         Begin ComctlLib.Slider SliderOne 
            Height          =   255
            Left            =   1080
            TabIndex        =   9
            Top             =   4560
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            _Version        =   327682
            Max             =   100
            TickFrequency   =   5
         End
         Begin VB.Label Label9 
            Caption         =   "IO Reference"
            Height          =   255
            Left            =   1080
            TabIndex        =   10
            Top             =   3600
            Width           =   1455
         End
      End
      Begin VB.Frame TabPage 
         Height          =   7560
         Index           =   2
         Left            =   6600
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   7935
         Begin VB.CommandButton Reset_cmd 
            Caption         =   "Reset New       Value"
            Height          =   495
            Left            =   6720
            TabIndex        =   13
            Top             =   6960
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Display in %"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   7080
            Width           =   1695
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6135
            ItemData        =   "FrmTest.frx":0000
            Left            =   120
            List            =   "FrmTest.frx":0002
            TabIndex        =   14
            Top             =   720
            Width           =   7695
         End
         Begin VB.Label label5 
            Caption         =   "Address"
            Height          =   255
            Left            =   4680
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "    Value    ( 0 - 255 )"
            Height          =   375
            Left            =   5520
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "New Value"
            Height          =   375
            Left            =   6360
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   420
            Width           =   1455
         End
      End
      Begin VB.Frame TabPage 
         Height          =   4560
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TBBlockNumDI 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   375
            Left            =   2520
            TabIndex        =   42
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox TBChannelNumDI 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   375
            Left            =   2520
            TabIndex        =   41
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox TBSwitchDes 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   480
            TabIndex        =   40
            Top             =   3360
            Width           =   4935
         End
         Begin VB.Label Label6 
            Caption         =   "Block Address"
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   1035
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "IO  Reference"
            Height          =   255
            Left            =   480
            TabIndex        =   45
            Top             =   1980
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Switch Description"
            Height          =   255
            Left            =   480
            TabIndex        =   44
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Warning: Model Halted"
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
            Height          =   300
            Left            =   1800
            TabIndex        =   43
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame TabPage 
         Height          =   5040
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   7440
         Visible         =   0   'False
         Width           =   7935
         Begin VB.OptionButton DOOptAllOne 
            Caption         =   "All Channels"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   33
            Top             =   720
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton DOOptAllOne 
            Caption         =   "Individual Channel"
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   32
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   4455
            Begin VB.OptionButton Option1 
               Caption         =   "On - Single Module "
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   31
               Top             =   240
               Width           =   2175
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Off "
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   30
               Top             =   1320
               Width           =   2175
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Flashing - Single Module"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   29
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton Option1 
               Caption         =   "On - All modules "
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   28
               Top             =   960
               Width           =   2175
            End
         End
         Begin VB.TextBox TBBlockNumDO 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   1140
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   3360
            TabIndex        =   24
            Text            =   "Combo2"
            Top             =   4080
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3360
            TabIndex        =   23
            Text            =   "Combo1"
            Top             =   4440
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   21
            Text            =   "0 0 0 1"
            Top             =   3840
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "ER Control Air On"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   3960
            Width           =   1575
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   495
            Index           =   0
            Left            =   3000
            TabIndex        =   22
            Top             =   4200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDownblock 
            Height          =   375
            Left            =   3120
            TabIndex        =   26
            Top             =   1140
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            AutoBuddy       =   -1  'True
            OrigLeft        =   2880
            OrigTop         =   1080
            OrigRight       =   3120
            OrigBottom      =   1455
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   34
            Top             =   4200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   495
            Index           =   2
            Left            =   2520
            TabIndex        =   35
            Top             =   4200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   495
            Index           =   3
            Left            =   2280
            TabIndex        =   36
            Top             =   4200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Block Address"
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Light"
            Height          =   255
            Left            =   3360
            TabIndex        =   37
            Top             =   3840
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin ComctlLib.TabStrip TabStrip1 
         Height          =   8445
         Left            =   0
         TabIndex        =   47
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   14896
         MultiRow        =   -1  'True
         TabFixedWidth   =   8916
         TabFixedHeight  =   11642
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   4
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Digital Input Tests"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Digital Output Tests"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Analogue Input Tests"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Analogue Output Test"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   10800
         Top             =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   2280
         TabIndex        =   49
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   135
         Left            =   1680
         TabIndex        =   48
         Top             =   480
         Width           =   15
      End
   End
   Begin VB.Menu mnuCard 
      Caption         =   "&Card"
      Visible         =   0   'False
      Begin VB.Menu mnuCard1 
         Caption         =   "MECD"
      End
      Begin VB.Menu mnuCard2 
         Caption         =   "Electrical"
      End
      Begin VB.Menu mnuBoth 
         Caption         =   "Both"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmHardwareTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentTabPage As Byte
Dim OldChannelNO As Integer
Dim OldblockNO As Integer
Dim OldCard As Integer
Dim CurrentWinCard As Integer
Dim CBox As New clsBoxSearch




Private Sub Reset()
Dim i As Integer
'set up digital output tab page
  
  'Label3.Enabled = False
  'Label4.Enabled = False
  Label1.Enabled = False
  'set up block numbers
  If mnuCard1.Checked Then
    UpDownblock.max = D_OUTPUT_END_ADDR
    UpDownblock.Min = D_OUTPUT_START_ADDR
    UpDownblock.Enabled = True
    
    UpDown1.Min = 0
    UpDown1.max = AO_OFFSET - 1
  Else
    If mnuCard2.Checked Then
      UpDownblock.max = D_OUTPUT_END_ADDR_B
      UpDownblock.Min = D_OUTPUT_START_ADDR_B
      UpDownblock.Enabled = True
    
      UpDown1.Min = AO_OFFSET
      UpDown1.max = REALADDR_AO 'NO_A_OUTPUT
    Else
      UpDown1.Min = 0
      UpDown1.max = REALADDR_AO
    End If
  End If
  
  TBBlockNumDO.text = Format(UpDownblock.Value, "000")
  DOOptAllOne(0).Value = True
  DOOptAllOne(0).Value = True
  DOOptAllOne_Click 0
  Option1(1).Value = True ' placed here to delay message and avoid fault
'set up digital input
  TBBlockNumDI.text = ""
  TBChannelNumDI = ""
  TBSwitchDes.text = ""
  TBSwitchDes.Enabled = False
'set up analogue output
  optAOOption(2).Value = True
  SliderAll.Value = 0
  SliderAll.Enabled = False
  SliderOne.Value = 0
  SliderOne.Enabled = False
  TBGaugeNum.text = "0"
  TBGaugeNum.Enabled = False
'set up analogue input
  List1.clear
 
'set up arrays
  For i = 0 To REALADDR_DO 'NO_D_OUTPUT
        D_OUTPUT(i) = False
  Next
  For i = 0 To REALADDR_DI 'NO_D_INPUT
    D_INPUT(i) = False
  Next
  For i = 0 To REALADDR_AO 'NO_A_OUTPUT
    A_OUTPUT(i) = 0
  Next
  For i = 0 To REALADDR_AI 'NO_A_INPUT
    A_INPUT(i) = 0
  Next
  Combo1_Click
End Sub

Private Sub Combo1_Click()
  Dim SQL As String
       
  SQL = "SELECT [i/o ref]"
  Combo2.clear
  
  If Combo1.ListIndex <> -1 Then
  Select Case Combo1.ListIndex
    Case 0  ' desciption
    If mnuCard1.Checked Then
      SQL = SQL & ", [Description] From [DO card 1]"
    Else
      If mnuCard2.Checked Then
        SQL = SQL & ", [Description] From [DO card 2]"
      Else
        SQL = SQL & ", [Description] From [DO]"
      End If
    End If
      'Label4.Caption = "Description"
    Case 1  ' io reference
      If mnuCard1.Checked Then
        SQL = SQL & ",[i/o ref] from [DO card 1]"
      Else
        If mnuCard2.Checked Then
          SQL = SQL & ",[i/o ref] from [DO card 2]"
        Else
          SQL = SQL & ", [i/o ref] From [DO]"
        End If
      End If
      Label4.Caption = "IO Reference"
    Case 2  ' Drawing reference
      If mnuCard1.Checked Then
        SQL = SQL & ", [Drawing Ref] From [DO card 1]"
      Else
        If mnuCard2.Checked Then
          SQL = SQL & ", [Drawing Ref] From [DO card 2]"
        Else
          SQL = SQL & ", [Drawing Ref] From [DO]"
        End If
      End If
      Label4.Caption = "Drawing Reference"
    Case 3
    Dim i As Integer
    For i = 0 To REALADDR_DO
      Combo2.AddItem str(i)
      Combo2.ItemData(i) = i
    Next
  End Select
    If Combo1.ListIndex <> 3 Then
      FillList DB, Combo2, SQL
    End If
  End If
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
'preform a search for match using text in combo box text
  CBox.FindIndexStr Combo1, Combo1.text, KeyAscii
End Sub

Private Sub Combo2_Click()
  'use IO ref held in item data array of combo2
  'to set do array
  If DOOptAllOne(1).Value And Combo2.ListIndex <> -1 Then 'testing individual lights
   
    D_OUTPUT(OldChannelNO) = False  'switch off last light
    D_OUTPUT(Combo2.ItemData(Combo2.ListIndex)) = True 'switch on current light
    OldChannelNO = Combo2.ItemData(Combo2.ListIndex) ' store old light index
 End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'searches for a match in drop down list of combo2
 If Combo1.ListIndex <> 3 Then
  CBox.FindIndexStr Combo2, Combo2.text, KeyAscii
  'Else
   ' If Val(Combo2.Text & Chr(KeyAscii)) >= 0 And Val(Combo2.Text & Chr(KeyAscii)) <= REALADDR_DO Then
    '  D_OUTPUT(OldChannelNO) = False  'switch off last light
     ' D_OUTPUT(Val(Combo2.Text & Chr(KeyAscii))) = True
     ' OldChannelNO = Val(Combo2.Text & Chr(KeyAscii))
    'Else
     ' Combo2.Text = ""
      'KeyAscii = 0
    'End If
 End If
 
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Activate()
Dim i As Long, start As Long, finish As Long

' === 19/11/02 ADDED for HA TEST then REMOVED again ...===
'Exit Sub


'get di states. Allowing push switch changes to be seen on a panel
'where a perminant switch has a higher io refernece number

If mnuCard1.Checked Then
        Set_D_Output Timer1, CurrentWinCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR
        start = 0
        finish = DI_OFFSET - 1
Else
         Set_D_Output Timer1, CurrentWinCard, D_OUTPUT_START_ADDR_B, D_OUTPUT_END_ADDR_B
         start = DI_OFFSET
         finish = REALADDR_DI 'NO_D_INPUT
End If
        
For i = start To finish
        Changes(i) = D_INPUT(i)
Next
      
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)
      
End Sub

Private Sub Form_Load()
Dim Status As Integer, status2 As Integer
Dim SQL As String
Dim i As Integer
Dim Ret_value As Boolean, Ret_value2 As Boolean
FrmHardwareTest.Width = TabStrip1.Width + 300
FrmHardwareTest.Height = TabStrip1.Height + 700


' === 19/11/02 ADDED for HA TEST then REMOVED again ...===
'Exit Sub



' === 12/12/01 ===
AI_update_required = True

  'set up profibus
  'Ret_value = IO_Init(WINCARD_A, Status)
  'Ret_value2 = IO_Init(WINCARD_B, status2)
  
  'If Ret_value = False And Ret_value2 = False Then
  '  ErrorMessage Status
   ' Unload Me
  'Else
    'set current win cardS
    mnuCard1_Click
    'set up tab book pages
    'Reset
    'set up tab book
    CurrentTabPage = 1
    TabStrip1_Click
    'set up card menu
    'mnuCard1.Checked = True
    'mnuCard2.Checked = False
    'mnuBoth.Checked = False
  
    'set up combo box on do test tab page
    With Combo1
      .AddItem "Description"
      .ItemData(0) = 0
      .AddItem "I/O Reference"
      .ItemData(1) = 2
      .AddItem "Drawing Reference"
      .ItemData(2) = 4
      .AddItem "User Digital Output IO Ref"
      .text = ""
      .Enabled = False
    End With
    Combo2.text = ""
    Combo2.Enabled = False
  'End If

  'set up di database
  'SQL = "Select [i/o ref], [Description] From DI"
  'DIDataBase.Open SQL, DB, adOpenDynamic, adLockReadOnly, adCmdText
  
  'OpenDI_AI_DataBase ' move this and closedi_ai_database to form host load event
End Sub

Private Sub DOOptAllOne_Click(Index As Integer)
'allows the choice of testing all lights at once
'or test lights individualy
Dim i As Integer

  Select Case Index
    Case 0  ' test all lights
      'enable option choices
     If mnuBoth.Checked = False Then
      Option1(0).Enabled = True 'all on
      TBBlockNumDO.Enabled = True
      UpDownblock.Enabled = True
      Option1(2).Enabled = True 'all flashing
      Label2.Enabled = True
     Else
      Option1(0).Enabled = False 'all on
      TBBlockNumDO.Enabled = False
      UpDownblock.Enabled = False
      Option1(2).Enabled = False 'all flashing
      Label2.Enabled = False
     End If
      Option1(1).Enabled = True 'all off
      Option1(3).Enabled = True
      Option1(1).Value = True 'select option
      'disable option two choices i.e. testing individualy
      '17/11/99
      Combo1.Enabled = False 'selection of light
      Combo2.Enabled = False
      Label1.Enabled = False
      'Label3.Enabled = False
      'Label4.Enabled = False
      D_OUTPUT(OldChannelNO) = False  'switch off any single light
      
      UpDown2(0).Enabled = False
      UpDown2(1).Enabled = False
      UpDown2(2).Enabled = False
      UpDown2(3).Enabled = False
      Text1.Enabled = False
      
      
      
    Case 1  'individual test
      'switch off all lights
      For i = 0 To REALADDR_DO 'NO_D_OUTPUT
        D_OUTPUT(i) = False
      Next
      '17/11/99 just use updowns
      'enable option choices
      'Combo1.Enabled = True
      'Combo2.Enabled = True
      'Label2.Enabled = False
      'TBBlockNumDO.Enabled = False
      'UpDownblock.Enabled = False
      'Label3.Enabled = True
      'Label4.Enabled = True
       UpDown2(0).Enabled = True
       UpDown2(1).Enabled = True
       UpDown2(2).Enabled = True
       UpDown2(3).Enabled = True
       Text1.Enabled = True
       
      
      'disable option one choices
      Option1(0).Enabled = False  'all on
      Option1(1).Enabled = False  'all off
      Option1(2).Enabled = False  'all flashimg
      Option1(3).Enabled = False
      Option1(1).Value = True ' select
      'set light accordind to text box value
      OldChannelNO = 0  'copy channel number
  End Select
End Sub



'Private Sub Form_Unload(Cancel As Integer)
'Dim Ret_value As Boolean
'Dim Status As Integer
  'stop io on bus. Must be called same number of time as io_init
  'Ret_value = IO_Exit(WINCARD_A, Status)
  'Ret_value = IO_Exit(WINCARD_B, Status)
  'close database's
  'DIDataBase.Close
  'CloseDI_AI_Database
'End Sub

Private Sub mnuBoth_Click()
'test all cards at once
  mnuBoth.Checked = True
  mnuCard1.Checked = False
  mnuCard2.Checked = False
  Reset
End Sub

Private Sub mnuCard1_Click()
Dim SQL As String
'set current card
  OldCard = CurrentWinCard
  CurrentWinCard = WINCARD_A
  ' set tick
  mnuCard1.Checked = True
  mnuCard2.Checked = False
  mnuBoth.Checked = False
  'set updown channel min and max
  'UpDownChannel.Min = 0
  'UpDownChannel.max = DO_OFFSET - 1
  'UpDownChannel.Value = 0
  'set up down block numbers
  'UpDownblock.Min = D_OUTPUT_START_ADDR
  'UpDownblock.max = D_OUTPUT_END_ADDR
  'UpDownblock.Value = D_OUTPUT_START_ADDR
  
  Reset
  'UpDown1.Min = 0
  'UpDown1.max = AO_OFFSET - 1
  
  'Combo1_Click
  
  'set up di database
  'SQL = "Select [i/o ref], [Description] from [DI card 1]"
  'DIDataBase.Open SQL, DB, adOpenDynamic, adLockReadOnly, adCmdText
End Sub

Private Sub mnuCard2_Click()
Dim SQL As String
'set current card
  OldCard = CurrentWinCard
  CurrentWinCard = WINCARD_B
  'move tick on menu
  mnuCard1.Checked = False
  mnuCard2.Checked = True
  mnuBoth.Checked = False
  ' set updown channel min and max
  'UpDownChannel.Min = DO_OFFSET
  'UpDownChannel.max = NO_D_OUTPUT - 1
  'UpDownChannel.Value = UpDownChannel.Min
  'set up down block numbers
  'UpDownblock.Min = D_OUTPUT_START_ADDR_B
  'UpDownblock.max = D_OUTPUT_END_ADDR_B
  'UpDownblock.Value = D_OUTPUT_START_ADDR_B
  
  
  'UpDown1.Min = AO_OFFSET
  'UpDown1.max = NO_A_OUTPUT
  
  Reset
  
  'Combo1_Click
  'set up di database
  'SQL = "Select [i/o ref], [Description] from [DI card 2]"
  'DIDataBase.Open SQL, DB, adOpenDynamic, adLockReadOnly, adCmdText
End Sub

Private Sub optAOOption_Click(Index As Integer)
'test analogue output
Dim i As Integer, Value As Integer
  Select Case Index
    Case 0  'all full scale
      'diable sliders and textbox
      SliderAll.Enabled = False
      SliderAll.Value = 0
      SliderOne.Value = 0
      SliderOne.Enabled = False
      TBGaugeNum.Enabled = False
      For i = 0 To REALADDR_AO 'NO_A_OUTPUT
        A_OUTPUT(i) = 255   'scaled value
      Next
    Case 1  'all half scale
     'diable sliders and textbox
      SliderAll.Enabled = False
      SliderAll.Value = 0
      SliderOne.Value = 0
      SliderOne.Enabled = False
      TBGaugeNum.Enabled = False
      For i = 0 To REALADDR_AO 'NO_A_OUTPUT
        A_OUTPUT(i) = 128 ' scaled value
      Next
    Case 2  'all zero
       'diable sliders and textbox
      SliderOne.Value = 0
      SliderAll.Value = 0
      SliderAll.Enabled = False
      SliderOne.Enabled = False
      TBGaugeNum.Enabled = False
      For i = 0 To REALADDR_AO 'NO_A_OUTPUT
        A_OUTPUT(i) = 0
      Next
    Case 3  'all variable
     'disable sliders and textbox. Test value for sliders change event
      SliderOne.Value = 0
      For i = 0 To REALADDR_AO 'NO_A_OUTPUT
        A_OUTPUT(i) = 0
      Next
      SliderAll.Enabled = True
      SliderOne.Enabled = False
      TBGaugeNum.Enabled = False
    Case 4  ' single variable
      For i = 0 To REALADDR_AO 'NO_A_OUTPUT
        A_OUTPUT(i) = 0
      Next
     'disable sliders and textbox. Test value from sliders change event
      SliderAll.Enabled = False
      SliderAll.Value = 0
      SliderOne.Enabled = True
      TBGaugeNum.Enabled = True
  End Select
End Sub

Private Sub Option1_Click(Index As Integer)
'selects a method of testing all lights
'in the digital output array that correspond to the block number in textbox
Dim i As Integer
Dim temp As Integer

  Select Case Index
    Case 0  'single module on
    temp = GetChannelNum(Val(TBBlockNumDO), CurrentWinCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR, D_OUTPUT_START_ADDR_B) + 1  ' start of block
      For i = temp To temp + NO_D_OUTPUT_CHANNELS - 1 Step 2
        D_OUTPUT(i - 1) = False
        D_OUTPUT(i) = True
      Next
    Case 1  'all off
      'temp = GetChannelNum(Val(TBBlockNumDO), CurrentWinCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR, D_OUTPUT_START_ADDR_B) '+ 1 ' start of block
      'For i = temp To temp + NO_D_OUTPUT_CHANNELS - 1
      '25/8/25
      For i = 0 To REALADDR_DO
        D_OUTPUT(i) = False
      Next
    Case 2  'single module flashing
    temp = GetChannelNum(Val(TBBlockNumDO), CurrentWinCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR, D_OUTPUT_START_ADDR_B) + 1 ' start of block
      For i = temp To temp + NO_D_OUTPUT_CHANNELS - 1 Step 2
        D_OUTPUT(i) = False
        D_OUTPUT(i - 1) = True
      Next
    Case 3 ' all on steady
      For i = 1 To REALADDR_DO Step 2
        D_OUTPUT(i) = True
      Next i
  End Select
End Sub

Private Sub Reset_cmd_Click()
Dim i As Integer


' 15/11/99
For i = 0 To 64
    OLD_A_input(i) = A_INPUT(i)
Next i

End Sub

Private Sub SliderAll_Change()
'provides meter value for all gauges
Dim i As Integer, Value As Integer
  Value = ((AO_MAX * (SliderAll.Value / 100) / AO_MAX) * 255)
  For i = 0 To NO_A_OUTPUT
    A_OUTPUT(i) = Value
  Next
End Sub


Private Sub SliderOne_Click()
  'provides value for a single gauge
  A_OUTPUT(Val(TBGaugeNum.text)) = (SliderOne.Value / 100) * 255
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
'allows data on tab to be reset as tab page is to be changed
  Reset
  If (TabStrip1.SelectedItem.Index = 2 Or TabStrip1.SelectedItem.Index = 4) And Timer1.Enabled Then
  'zero gauges and switch off lights when changing tab page
   Timer1_Timer
  End If
End Sub

Private Sub TabStrip1_Click()
'displays the correct page in tab strip
Dim NewPage As Byte

  NewPage = TabStrip1.SelectedItem.Index - 1  'page to display
  If NewPage <> CurrentTabPage Then 'test if page is currently shown
    'position page in client area of tab strip
    TabPage(NewPage).Left = TabStrip1.ClientLeft + 50
    TabPage(NewPage).Top = TabStrip1.ClientTop
    TabPage(NewPage).Visible = True
    TabPage(CurrentTabPage).Visible = False
    CurrentTabPage = NewPage  ' copy current page number
    If NewPage = 1 Then ' do page. show combo boxes
      '17/11/99 hide combo boxes now we have updowns
      'Combo1.ListIndex = -1
      Combo1.Enabled = False
      Combo1.Visible = False
      'Combo2.clear
      Combo2.Enabled = False
      Combo2.Visible = False
    Else
      Combo1.Enabled = False
      Combo1.Visible = False
      Combo2.Enabled = False
      Combo2.Visible = False
    End If
  End If
End Sub

'Private Sub TBBlockNumDO_Change()
'updates updown value.Keeping text box and updown value consistant
'on d_output frame
  'TBBlockNumDO.Text = Format(TBBlockNumDO.Text, "000")
  'If Val(TBBlockNumDO.Text) >= UpDownblock.Min And Val(TBBlockNumDO.Text) <= UpDownblock.max Then
   ' UpDownblock.Value = Val(TBBlockNumDO.Text)
  'Else
   ' UpDownblock.Value = UpDownblock.Min
  '  TBBlockNumDO.Text = Format(UpDownblock.Value, "000")
 ' End If
'End Sub

Private Sub TBBlockNumDO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Val(TBBlockNumDO.text) >= UpDownblock.Min And Val(TBBlockNumDO.text) <= UpDownblock.max Then
    UpDownblock.Value = Val(TBBlockNumDO.text)
  Else
    TBBlockNumDO.text = str(UpDownblock.Value)
  End If
  TBBlockNumDO.text = Format(TBBlockNumDO.text, "000")
End If
End Sub

Private Sub TBBlockNumDO_LostFocus()
'Keep text and up down value coherent
  TBBlockNumDO_KeyPress 13
End Sub

'Private Sub TBGaugeNum_Change()
    'If Val(TBGaugeNum.Text) >= UpDown1.Min And Val(TBGaugeNum) <= UpDown1.max Then
     ' UpDown1.Value = Val(TBGaugeNum)
    'Else
     ' TBGaugeNum.Text = str(UpDown1.Value)
    'End If
'End Sub

Private Sub TBGaugeNum_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Val(TBGaugeNum.text) >= UpDown1.Min And Val(TBGaugeNum) <= UpDown1.max Then
      UpDown1.Value = Val(TBGaugeNum)
    Else
      TBGaugeNum.text = str(UpDown1.Value)
    End If
    SliderOne_Click
  End If
End Sub

Private Sub TBGaugeNum_LostFocus()
' Keep text box and updown value coherent
  TBGaugeNum_KeyPress 13
End Sub

Private Sub Timer1_Timer()

' === 19/11/02 ADDED for HA TEST then REMOVED again ...===
'Exit Sub


'Update Profibus and blocks
'If mnuBoth.Checked = False Then
'  Select Case TabStrip1.SelectedItem.index
'    Case 1
'      'test digital input
'      If mnuCard1.Checked Then
'        Get_D_input Timer1, CurrentWinCard, D_INPUT_START_ADDR, D_INPUT_END_ADDR
'      Else
'        Get_D_input Timer1, CurrentWinCard, D_INPUT_START_ADDR_B, D_INPUT_END_ADDR_B
'     End If
'      DITest CurrentWinCard
'    Case 2
 '     ' test digital output
'      If mnuCard1.Checked Then
'        Set_D_Output Timer1, CurrentWinCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR
'      Else
'         Set_D_Output Timer1, CurrentWinCard, D_OUTPUT_START_ADDR_B, D_OUTPUT_END_ADDR_B
'     End If
'    Case 3
'      'test analogue input
'      If mnuCard1.Checked Then
'        Get_A_Input Timer1, CurrentWinCard, A_INPUT_START_ADDR, A_INPUT_END_ADDR
'      Else
 '       Get_A_Input Timer1, CurrentWinCard, A_INPUT_START_ADDR_B, A_INPUT_END_ADDR_B
  '    End If
  '    AITest CurrentWinCard
'    Case 4
'      'test analogue output
'      If mnuCard1.Checked Then
'        Set_A_Output Timer1, CurrentWinCard, A_OUTPUT_START_ADDR, A_OUTPUT_END_ADDR
'      Else
'        Set_A_Output Timer1, CurrentWinCard, A_OUTPUT_START_ADDR_B, A_OUTPUT_END_ADDR_B
'      End If
'  End Select
'Else
  ReadAllCards Timer1
  If TabStrip1.SelectedItem.Index = 1 Then
    'di
    DITest 0
  Else
    If TabStrip1.SelectedItem.Index = 3 Then
      AITest 0
    End If
  End If
'End If
End Sub

Private Sub UpDown1_Change()
  TBGaugeNum.text = str(UpDown1.Value)
  Dim i As Integer
  For i = 0 To REALADDR_AO
    A_OUTPUT(i) = 0
  Next
End Sub

Private Sub UpDown2_Change(Index As Integer)
Dim i As Integer

DO_number = UpDown2(0).Value + UpDown2(1).Value * 10 + UpDown2(2).Value * 100 + UpDown2(3).Value * 1000

Text1.text = Format(UpDown2(3).Value, "0") & " " & Format(UpDown2(2).Value, "0") & " " & Format(UpDown2(1).Value, "0") & " " & Format(UpDown2(0).Value, "0")
 For i = 1 To UBound(D_OUTPUT)
    'D_OUTPUT(i) = False
    If i = 1137 Or i = 1155 And Check2.Value Then 'Fix to allow boiler valves to have air
        D_OUTPUT(i) = True
    Else
        D_OUTPUT(i) = False
    End If
    
 Next
 If DO_number > 0 And DO_number < UBound(D_OUTPUT) Then
    D_OUTPUT(DO_number) = True
 End If
 

End Sub

Private Sub UpDownBlock_Change()
'allows the selection of individual blocks
Dim i As Integer, Block As Integer
  'allows switching off lights of as  block num changes
  If UpDownblock.Value <> OldblockNO Then  'block has changed
    If DOOptAllOne(0) Then
      'testing all lights
      Block = GetChannelNum(OldblockNO, OldCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR, D_OUTPUT_START_ADDR_B) 'array index of first channel on block
      OldCard = CurrentWinCard 'stop errors after switch lights off on last block on different card
      'switch all lights on old block off
      For i = Block To Block + NO_D_OUTPUT_CHANNELS - 1
        D_OUTPUT(i) = False
      Next
      Option1(1).Value = True
    End If
    OldblockNO = UpDownblock.Value  'must be before changing updownchannel value
    'UpDownChannel.Value = GetChannelNum(UpDownblock.Value, CurrentWinCard, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR, D_OUTPUT_START_ADDR_B) 'array index of first channel on block)
    TBBlockNumDO.text = Format(UpDownblock.Value, "000")
  End If
End Sub
