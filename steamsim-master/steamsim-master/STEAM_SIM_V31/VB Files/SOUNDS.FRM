VERSION 5.00
Object = "{044E4D83-136A-101C-99A2-86B9AD896A58}#4.0#0"; "FXSND432.OCX"
Begin VB.Form sounds 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   3120
   ClientTop       =   1620
   ClientWidth     =   7755
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   7755
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Sounds"
      Height          =   495
      Left            =   3720
      TabIndex        =   39
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4200
      TabIndex        =   38
      Text            =   "Text20"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   37
      Text            =   "Compressor 1"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "On"
      Height          =   495
      Left            =   720
      TabIndex        =   36
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Off"
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Text            =   "Stop fan"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   33
      Text            =   "Play fan"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1215
      Index           =   0
      Left            =   2760
      Max             =   100
      TabIndex        =   31
      Top             =   3000
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Sound Mixer "
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         Height          =   495
         Left            =   2640
         TabIndex        =   32
         Top             =   2280
         Width           =   615
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Index           =   6
         Left            =   5160
         Max             =   100
         TabIndex        =   30
         Top             =   840
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Index           =   5
         Left            =   4440
         Max             =   100
         TabIndex        =   29
         Top             =   840
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Index           =   4
         Left            =   3600
         Max             =   100
         TabIndex        =   28
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Text            =   "Compressors"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Text            =   "Pumps"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Text            =   "Fans"
         Top             =   480
         Width           =   735
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Index           =   3
         Left            =   1920
         Max             =   100
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Index           =   2
         Left            =   1200
         Max             =   100
         TabIndex        =   23
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Text            =   "Prop-shaft"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Text            =   "Turbo"
         Top             =   480
         Width           =   615
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1215
         Index           =   1
         Left            =   480
         Max             =   100
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Text            =   "Engine"
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Text            =   "Text14"
      Top             =   3840
      Width           =   3495
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Text            =   "Stop prop"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Text            =   "Stop turbo"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Text            =   "Stop engine"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Text            =   "Play prop"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Text            =   "Play turbo"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "Play engine"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "On"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Off"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "On"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Off"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "On"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Off"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Text            =   "SW suc V/V"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "SW Pump 1"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   5040
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "LO Pump 1"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7080
      Top             =   600
   End
   Begin Fxsnd32LibCtl.FXSnd snd 
      Left            =   7080
      Top             =   120
      _Version        =   262144
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      SoundType       =   0
      AutoPlay        =   0   'False
      FileName        =   ""
      Command         =   0
      Notify          =   -1  'True
      From            =   0
      To              =   0
      Repeat          =   0   'False
      DeviceType      =   ""
      Alias           =   ""
      NotifyDelay     =   0
      DSCommand       =   0
      DSSpeaker       =   1
   End
End
Attribute VB_Name = "sounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
' LUB oil pump
DIGIN(EOP1) = False
End Sub

Private Sub Command2_Click()
' LUB oil pump
DIGIN(EOP1) = True
End Sub



Private Sub Command3_Click()
' SW oil pump
DIGIN(SWP1) = False
End Sub

Private Sub Command3D3_Click()

' SW oil pump
DIGIN(SWP1) = False

End Sub

Private Sub Command3D4_Click()

' SW oil pump
DIGIN(SWP1) = True

End Sub

Private Sub Command3D5_Click()

' SW suction V/V
DIGIN(SWVV) = True


End Sub

Private Sub Command3D6_Click()

' SW suction V/V
DIGIN(SWVV) = False


End Sub

Private Sub Command4_Click()
' SW oil pump
DIGIN(SWP1) = True
End Sub

Private Sub Command5_Click()
' SW suction V/V
DIGIN(SWVV) = False
End Sub

Private Sub Command6_Click()
' SW suction V/V
DIGIN(SWVV) = True
End Sub

Private Sub Command7_Click()
Form1.Visible = False
End Sub

Private Sub Command8_Click()

' Air compressor 1
DIGIN(SAC1) = False

End Sub

Private Sub Command9_Click()
' Air compressor 1
DIGIN(SAC1) = True
End Sub

Private Sub Form_Load()
Dim i As Integer

Form1.Width = 6915
Form1.Height = 6180

' == 6/7/97 added =============

snd.SoundType = 2 ' DirectSound
snd.DSCommand = 1 ' Initialize DirectSound

' 7/7/97 changed to snd.FXError
'If (FXError > 0) Then
If (snd.FXError > 0) Then
MsgBox ("Snd Demo failed to initialize DirectSound!" + Chr$(13) + "This example requires the MS-Game SDK for Windows 95")
End If

snd.DSChannel = 1 ' Channel 1
snd.DSRepeat = True
snd.FileName = "c:\Audrack.w32\eng4.wav"
snd.DSCommand = 2 ' Load wave data into buffer

snd.DSChannel = 2 ' Channel 2
snd.DSRepeat = True
snd.FileName = "c:\Audrack.w32\turbo2.wav"
snd.DSCommand = 2 ' Load wave data into buffer

snd.DSChannel = 3 ' Channel 3
snd.DSRepeat = True
snd.FileName = "c:\Audrack.w32\prop2.wav"
snd.DSCommand = 2 ' Load wave data into buffer

snd.DSChannel = 4 ' Channel 4
snd.DSRepeat = True
snd.FileName = "c:\Audrack.w32\fan2.wav"
snd.DSCommand = 2 ' Load wave data into buffer

snd.DSChannel = 5 ' Channel 5
snd.DSRepeat = True
snd.FileName = "c:\Audrack.w32\pump2.wav"
snd.DSCommand = 2 ' Load wave data into buffer

snd.DSChannel = 6 ' Channel 6
snd.DSRepeat = True
snd.FileName = "c:\Audrack.w32\comp_on2.wav"
snd.DSCommand = 2 ' Load wave data into buffer

snd.DSChannel = 7 ' Channel 7
snd.DSRepeat = False
snd.FileName = "c:\Audrack.w32\compoff2.wav"
snd.DSCommand = 2 ' Load wave data into buffer'=====================


'
' What was here is now in the main form's load event
'

'-----------------------------------------------------------------------------
end_this:

End Sub

Private Sub Text10_Click()

     snd.DSChannel = 3      ' Prop shaft sound
  snd.DSCommand = 3      ' Play the sound

End Sub

Private Sub Text11_Click()

     snd.DSChannel = 1      ' Engine sound
  snd.DSCommand = 4      ' Stop the sound

End Sub

Private Sub Text12_Click()

     snd.DSChannel = 2      ' Turbo charger sound
  snd.DSCommand = 4      ' Stop the sound

End Sub

Private Sub Text13_Click()

     snd.DSChannel = 3      ' Prop shaft sound
  snd.DSCommand = 4      ' Stop the sound

End Sub

Private Sub Text17_Click()

     snd.DSChannel = 4      ' Prop fan sound
  snd.DSCommand = 3      ' Play the sound

End Sub

Private Sub Text18_Click()

     snd.DSChannel = 4      ' Fan sound
  snd.DSCommand = 4      ' Stop the sound


End Sub

Private Sub Text8_Click()

     snd.DSChannel = 1        ' Main Engine sound
  snd.DSCommand = 3        ' Play the sound
  
End Sub

Private Sub Text9_Click()

     snd.DSChannel = 2      ' Turbo charger sound
  snd.DSCommand = 3      ' Play the sound
  
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

' Test
eng_set_point = VScroll2(0).Value
accn = (eng_set_point - eng_speed) * 0.5
eng_speed = eng_speed + accn * (Timer1.Interval * 0.001)
Text14.Text = eng_set_point & "  " & eng_speed & "  " & accn

      snd.DSChannel = 1        ' Main Engine sound
      snd.DSFrequency = 4000 + Int(eng_speed * 100)      ' SHOULD be engine speed from the model (ENGSPEED)
      snd.DSVolume = VScroll2(1).Value * -50
 
      snd.DSChannel = 2        ' Turbo sound
      snd.DSFrequency = 4000 + Int(eng_speed * 100)      ' SHOULD be turbo speed  (TURBSPEED)
      snd.DSVolume = VScroll2(2).Value * -50
      
      snd.DSChannel = 3     ' Prop shaft sound
      snd.DSFrequency = 4000 + Int(eng_speed * 100)       ' SHOULD be prop speed (PROPSPEED)
      snd.DSVolume = VScroll2(3).Value * -50
      
      snd.DSChannel = 4     ' Fan sound
      snd.DSVolume = VScroll2(4).Value * -50
      
      snd.DSChannel = 5     'Pump sound
      snd.DSVolume = VScroll2(5).Value * -50
      
      snd.DSChannel = 6      'Compressor ON sound
      snd.DSVolume = VScroll2(6).Value * -50
      
     snd.DSChannel = 7      'Compressor OFF loading sound
      snd.DSVolume = VScroll2(6).Value * -50
      
If Check1.Value = 1 Then     ' Sounds are Enabled

' Air Compressors
  If DIGOUT(SAC1LMP) And Not old_comp1 Then
      Text20.Text = "Compr. ON"
       snd.DSChannel = 6        ' Compressor ON sound
    snd.DSCommand = 3        ' Play the sound
  End If

  If Not DIGOUT(SAC1LMP) And old_comp1 Then
      Text20.Text = "Compr. OFF"
        snd.DSChannel = 6        ' Compressor ON sound
    snd.DSCommand = 4        ' Stop the sound
      snd.DSChannel = 7       ' Compressor OFF sound
    snd.DSCommand = 3        ' Play the sound
  End If

  If pumpson And Not old_pumpson Then
       Text20.Text = "pumps ON"
      snd.DSChannel = 5        ' Pump sound
    snd.DSCommand = 3        ' Play the sound
  End If
  If Not pumpson And old_pumpson Then
       Text20.Text = "pumps OFF"
       snd.DSChannel = 5        ' Pump sound
    snd.DSCommand = 4        ' Stop the sound
  End If
  
    old_comp1 = DIGOUT(SAC1LMP)
    old_pumpson = pumpson
  
Else

    old_comp1 = False           ' To ensure ON sound up-date when enabled
    old_pumpson = False       ' To ensure ON sound up-date when enabled

      For i = 1 To 7
           snd.DSChannel = i
        snd.DSCommand = 4        ' Stop the sound
      Next i
 
End If






List1.Clear
'text2.Text = "timer1 start"

' Put into tmrMain 27/6/97 IMO
    'Call LOGIC
    'Call ENGINE
    'Call INOUT
'''''''''
'sump level
'LOCATE 22, 6: PRINT USING "###"; RDA(11)
List1.AddItem "Sump level = " & Format$(RDA(11), "###")
'pump disch press
'LOCATE 37, 62: PRINT USING "#.#"; RDA(8)
List1.AddItem "LO pump disch press = " & Format$(RDA(8), "##.00")

'DP across filters
'        LOCATE 38, 27: PRINT USING "#.#"; RDA(27)
List1.AddItem "DP across filters = " & Format$(RDA(27), "##.00")


'lo inlet to clr temp
'LOCATE 35, 76: PRINT USING "###"; RDA(36)
List1.AddItem "lo inlet to clr temp  = " & Format$(RDA(36), "###")

'lo outlet clr temp
'LOCATE 19, 69: PRINT USING "###"; RDA(37)
List1.AddItem "lo outlet from clr temp  = " & Format$(RDA(37), "###")

'lo inlet to ENGINE RDA(6)
'        LOCATE 15, 47: PRINT USING "###"; RDA(6)
List1.AddItem "lo inlet to ENGINE  = " & Format$(RDA(6), "###")
'lo outlet from engine rda(7)
'        LOCATE 31, 17: PRINT USING "###"; RDA(7)
List1.AddItem "lo outlet from ENGINE  = " & Format$(RDA(7), "###")



List2.Clear
List2.AddItem "SW pump disch press = " & Format$(SWP, "##.00")


End Sub

Private Sub VScroll1_Change()

End Sub
