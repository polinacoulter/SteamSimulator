VERSION 5.00
Begin VB.Form CmdLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Mixer"
   ClientHeight    =   7755
   ClientLeft      =   3105
   ClientTop       =   1605
   ClientWidth     =   8490
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   8490
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   52
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   51
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   47
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   46
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   " Sound Mixer 1"
      Height          =   1275
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   7635
      Begin VB.TextBox FreqVal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox FreqVal1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   44
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HASP 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.HScrollBar scrVolume1 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   100
         TabIndex        =   33
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed1 
         Height          =   255
         Index           =   2
         Left            =   4920
         Max             =   1000
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "# 1 Turbo Feed Pump (R)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2835
      End
      Begin VB.HScrollBar scrVolume1 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   100
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed1 
         Height          =   255
         Index           =   1
         Left            =   4920
         Max             =   1000
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "# 1 Turbo Feed Pump (L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2835
      End
      Begin VB.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Sound Mixer 0"
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7635
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   6600
         TabIndex        =   49
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6600
         TabIndex        =   48
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   6600
         TabIndex        =   43
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   6600
         TabIndex        =   42
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   6600
         TabIndex        =   41
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6600
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6600
         TabIndex        =   39
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox FreqVal0 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6600
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HASP 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   8
         Left            =   3120
         Max             =   100
         TabIndex        =   31
         Top             =   3120
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   7
         Left            =   3120
         Max             =   100
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   6
         Left            =   3120
         Max             =   100
         TabIndex        =   29
         Top             =   2400
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   5
         Left            =   3120
         Max             =   100
         TabIndex        =   28
         Top             =   2040
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   4
         Left            =   3120
         Max             =   100
         TabIndex        =   27
         Top             =   1680
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   3
         Left            =   3120
         Max             =   100
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   100
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   7
         Left            =   4920
         Max             =   1000
         TabIndex        =   24
         Top             =   2760
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   6
         Left            =   4920
         Max             =   1000
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   5
         Left            =   4920
         Max             =   1000
         TabIndex        =   22
         Top             =   2040
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   4
         Left            =   4920
         Max             =   1000
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   3
         Left            =   4920
         Max             =   1000
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   2
         Left            =   4920
         Max             =   1000
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   8
         Left            =   4920
         Max             =   1000
         TabIndex        =   18
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CheckBox Check0 
         Caption         =   "ElecCircuit Breaker (R)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Hiss"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Propshaft"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "FD Fan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Steam Vent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Fan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "Pump"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2835
      End
      Begin VB.CheckBox Check0 
         Caption         =   "SSDG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2835
      End
      Begin VB.HScrollBar scrSpeed0 
         Height          =   255
         Index           =   1
         Left            =   4920
         Max             =   1000
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrVolume0 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   100
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Frequency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Volume"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8040
      Top             =   360
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   50
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "'Sound Test' Is Only Enabled When The Model Has Been HALTED And Sounds Are ENABLED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   3615
   End
End
Attribute VB_Name = "CmdLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' 6/4/98
Dim ice_strings(100) As String
Sub InitializeSliders()
Dim i As Integer

For i = 1 To 8
    SoundMix.scrSpeed0(i).Value = sndFreq(0, i) / 100
Next

For i = 1 To 2
    SoundMix.scrSpeed1(i).Value = sndFreq(1, i) / 100
Next


End Sub



Private Sub Check0_Click(Index As Integer)
    
    

End Sub

Private Sub Check1_Click(Index As Integer)
    

End Sub

Private Sub CmdExit_Click()
Dim i As Integer

'copy new state on enable flags then exit

If This_PC_number = 0 Then
    For i = 1 To 8
        SndEnabled(0, i) = Check0(i)
    Next
ElseIf This_PC_number = 0 Then

    For i = 1 To 2
        SndEnabled(1, i) = Check1(i)
    Next
End If

Unload Me


End Sub

Private Sub CmdLoad_Click()
Dim FileNo As Integer
Dim i As Integer
Dim mystr As String


On Error GoTo ErrH

FileNo = FreeFile(0)

Open "c:\STEAM_SIM\Sounds\SoundConfig.dat" For Input As #FileNo
        
For i = 1 To 8
    Input #FileNo, mystr
    sndFreq(0, i) = Val(mystr)
    Input #FileNo, mystr
    sndFreq(1, i) = Val(mystr)
Next

Close #FileNo
Call InitializeSliders

Exit Sub

ErrH:
resp = MsgBox("There was an error reading C:\Steam_Sim\SoundConfig.dat" & Chr(10) & _
       "Ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
Close #FileNo
End Sub

Private Sub CmdQuit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
Dim FileNo As Integer
Dim i As Integer

On Error GoTo ErrH

FileNo = FreeFile(0)

Open "c:\STEAM_SIM\Sounds\SoundConfig.dat" For Output As #FileNo
        
For i = 1 To 8
    Print #FileNo, sndFreq(0, i)
    Print #FileNo, sndFreq(1, i)
Next

'Input #FileNo, PC_name

'This_PC_number = Val(Mid$(PC_name, Len(PC_name) - 1, 2))


Close #FileNo

Exit Sub

ErrH:
resp = MsgBox("There was an error reading C:\Steam_Sim\SoundConfig.dat" & Chr(10) & _
       "Ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
Close #FileNo
      

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim file_num As Integer

Sounds_Enabled = True

' ===== 19/11/97 =========================
Dim ii As Integer

'26/11/99 CMA no need for non-english
GoTo Skip




If useEnglish Then
'  Open App.Path & "\Esounds.txt" For Input As #1
   file_num = FreeFile
   Debug.Print "App.path = " & App.Path
   
' 10/7/98
'  Open App.Path & "\Esounds.txt" For Input As #file_num
   Open "c:\HAVEN\Text\Esounds.txt" For Input As #file_num
   
Else

' 10/7/98
'  Open App.Path & "\Isounds.txt" For Input As #file_num
   Open "c:\HAVEN\Text\Isounds.txt" For Input As #file_num
   
End If

For ii = 1 To 15
   Input #file_num, ice_strings(ii)
Next ii
Close #file_num

SoundMix.Caption = ice_strings(1)
Frame1.Caption = ice_strings(2)
For ii = 0 To 5
    Label2(ii).Caption = ice_strings(ii + 3)
Next ii
Label1.Caption = ice_strings(9)
'frmSoundTest.Caption = ice_strings(10)
'Check2.Caption = ice_strings(11)
'Check4.Caption = ice_strings(12)
'Check6.Caption = ice_strings(13)
'Check5.Caption = ice_strings(14)
'Check7.Caption = ice_strings(15)

'=========================================
Skip:

    ' Setup the sounds
    
' 6/4/98 ADDED:
' 7/4/98 Moved to HOST.FRM (load)
'      frmHiddenSounds.InitialiseSounds
    
    ' Load in the remembered scroll bar levels and set the volumes
    'For i = 0 To 5
    '    scrVolume(i).Value = sndVolume(i)
    'Next i
    'frmHiddenSounds.SetVolumes
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    If DummyPCNumber = 0 Then
        For i = 1 To 8
            sndVolume(0, i) = scrVolume0(i).Value
        Next
    ElseIf DummyPCNumber = 1 Then
        For i = 1 To 2
            sndVolume(1, i) = scrVolume1(i).Value
        Next
    End If
    ' If not running, then as we close we also turn off the test sounds
    If Not d_run Then
        For i = 1 To 8
            frmHiddenSounds.snd.DSChannel = i
            frmHiddenSounds.snd.DSCommand = 4
        Next i
    End If
End Sub


Private Sub scrVolume_Change(Index As Integer)
    sndVolume(Index) = scrVolume(Index)
    frmHiddenSounds.SetVolumes
End Sub

Private Sub Option1_Click(Index As Integer)
Dim i As Integer

If Index = 0 Then
    Option1(1).Value = False
Else
    Option1(0).Value = False
End If

Debug.Print "option1(0).value = " & Option1(0).Value
Debug.Print "option1(1).value = " & Option1(1).Value


'turn off sounds & reset
If DummyPCNumber = 0 Then
    For i = 1 To 8
        frmHiddenSounds.snd.DSChannel = i
        frmHiddenSounds.snd.DSCommand = 4  'Stop Ch
        frmHiddenSounds.snd.DSCommand = 5  'Close Ch & Remove from Memory ready for different wav
    Next i
End If
If DummyPCNumber = 1 Then
    For i = 1 To 2
        frmHiddenSounds.snd.DSChannel = i
        frmHiddenSounds.snd.DSCommand = 4  'Stop Ch
        frmHiddenSounds.snd.DSCommand = 5  'Close Ch & Remove from Memory ready for different wav
    Next i
End If
frmHiddenSounds.snd.DSCommand = 6 'close direct sounds
DummyPCNumber = Index
Call frmHiddenSounds.InitialiseSounds

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim eng_set_point As Single
Dim accn As Single
Dim eng_speed As Single

On Error GoTo ErrH


If Sounds_Enabled Then
'    If d_run Then
'        SoundWhenModelRunning
'    Else
        'For Test, use the local settings
        'If frmSoundTest.Enabled = False Then frmSoundTest.Enabled = True
' 30/10/97
         Sound_test_on = True
        
        'eng_set_point = hscrEng(0).Value
        'accn = (eng_set_point - eng_speed) * 0.5
        'eng_speed = eng_speed + accn * (Timer1.Interval * 0.001)
        
        
        Label5.Caption = DummyPCNumber
        
        
        If DummyPCNumber = 0 Then 'do for PC 0
            For i = 1 To 8
                If SndEnabled(0, i) Then
                    Check0(i).Value = 1
                Else
                    Check0(i).Value = 0
                End If
                
                sndFreq(0, i) = scrSpeed0(i).Value * 100
                FreqVal0(i) = sndFreq(0, i)
                frmHiddenSounds.snd.DSChannel = i        '
                frmHiddenSounds.snd.DSFrequency = sndFreq(0, i)
                If Check0(i) Then
                    If frmHiddenSounds.snd.DSStatus <> 1 Then
                        frmHiddenSounds.snd.DSCommand = 3       ' Play the sound
                    End If
                Else
                    frmHiddenSounds.snd.DSCommand = 4           ' Stop the sound
                End If
            Next
        ElseIf DummyPCNumber = 1 Then
            For i = 1 To 2
                If SndEnabled(1, i) Then
                    Check1(i).Value = 1
                Else
                    Check1(i).Value = 0
                End If
                sndFreq(1, i) = scrSpeed1(i).Value * 100
                FreqVal1(i) = sndFreq(1, i)
                frmHiddenSounds.snd.DSChannel = i        '
                frmHiddenSounds.snd.DSFrequency = sndFreq(1, i)
                If Check1(i) Then
                    If frmHiddenSounds.snd.DSStatus <> 1 Then
                        frmHiddenSounds.snd.DSCommand = 3       ' Play the sound
                    End If
                Else
                    frmHiddenSounds.snd.DSCommand = 4           ' Stop the sound
                End If
            Next
        
        
        End If
        
        'frmHiddenSounds.snd.DSChannel = 1        ' #1 Turbo Feed Pump
        'frmHiddenSounds.snd.DSFrequency = scrSpeed(0).Value
         
        'frmHiddenSounds.snd.DSChannel = 2        ' #2 Turbo Feed Pump
        'frmHiddenSounds.snd.DSFrequency = scrSpeed(1).Value
         
        'frmHiddenSounds.snd.DSChannel = 3     ' Prop shaft sound
        'frmHiddenSounds.snd.DSFrequency = scrSpeed(6).Value
         
         
         
        'If Check1.Value = 1 Then
        '    frmHiddenSounds.snd.DSChannel = 1        '#1 Turbo Feed Pump
        '    frmHiddenSounds.snd.DSPan = 10000          'Left ch only
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3           ' Play the sound
        'Else
        '    frmHiddenSounds.snd.DSChannel = 1      ' Engine sound
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
        'If Check3.Value = 1 Then
        '    frmHiddenSounds.snd.DSChannel = 2        '#2 Turbo Feed Pump
        '    frmHiddenSounds.snd.DSPan = -10000          'Right ch only
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3 ' Play the sound
        'Else
        '    frmHiddenSounds.snd.DSChannel = 2      ' Turbo sound
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
        'Else
        'If Check2.Value = 1 Then
        '    frmHiddenSounds.snd.DSChannel = 3        'Steam Vent
        '    frmHiddenSounds.snd.DSPan = -10000          'Right ch only
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3 ' Play the sound
       ' Else
        '    frmHiddenSounds.snd.DSChannel = 3      'Steam Vent
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
  
            
        'End If
        
        'If Check4.Value = 1 Then
        '    frmHiddenSounds.snd.DSChannel = 3        ' Prop sound
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3      ' Play the sound
        'Else
        '    frmHiddenSounds.snd.DSChannel = 3      ' Prop sound
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
        
        'If Check5.Value = 1 Then
        '    frmHiddenSounds.snd.DSChannel = 4        ' Fan sound
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3         ' Play the sound
        'Else
        '    frmHiddenSounds.snd.DSChannel = 4      ' Fan sound
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
        
        'If Check6.Value = 1 Then
        
            '''   Debug.Print " Pump Sound ON"
            
        '    frmHiddenSounds.snd.DSChannel = 5          ' Pump sound
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3         ' Play the sound
        'Else
        '    frmHiddenSounds.snd.DSChannel = 5     ' Pump sound
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
                
        'If Check7.Value = 1 Then
         '   frmHiddenSounds.snd.DSChannel = 6        ' Compressor ON sound
        '    If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3         ' Play the sound
        'Else
        '    frmHiddenSounds.snd.DSChannel = 6      ' Compressor ON sound
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If
        
        'If Check8.Value = 1 Then
        '    frmHiddenSounds.snd.DSChannel = 8        ' Electrical Circuit Breaker ON
        '    frmHiddenSounds.snd.DSPan = 10000          'Left ch only
         '   If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3         ' Play the sound
            'Check8.Value = 0
        'Else
        '    frmHiddenSounds.snd.DSChannel = 8      ' Electrical Circuit Breaker OFF
        '    frmHiddenSounds.snd.DSCommand = 4      ' Stop the sound
        'End If

        
End If
' 27/11/99 just to test
    Exit Sub
' 27/11/99 just to test
    'Else
    ' If Sounds are disabled, then shade out the sound controls
   ' frmSoundTest.Enabled = False
    
' 30/10/97
      Sound_test_on = False

    For i = 1 To 6
        frmHiddenSounds.snd.DSChannel = i
        frmHiddenSounds.snd.DSCommand = 4
    Next i
    'Check2.Value = 0
    'Check4.Value = 0
    'Check5.Value = 0
    'Check6.Value = 0
    'Check7.Value = 0
' 27/11/99 just to test
'End If

Exit Sub


ErrH:



End Sub


