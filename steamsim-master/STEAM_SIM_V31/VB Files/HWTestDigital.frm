VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form HWTestDigital 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Hardware Test"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9390
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   3375
      Left            =   0
      Top             =   60
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
      Begin VB.Timer Timer1 
         Interval        =   900
         Left            =   2580
         Top             =   3000
      End
      Begin VB.Frame Frame1 
         Caption         =   "Digital Input Changes"
         Height          =   3255
         Left            =   2760
         TabIndex        =   8
         Top             =   0
         Width           =   5535
         Begin VB.Label Label1 
            BackColor       =   &H8000000A&
            Caption         =   "Digital Input Channel"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "ON"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "OFF"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   0
            Left            =   2640
            TabIndex        =   12
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   1
            Left            =   3360
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Index           =   2
            Left            =   1080
            TabIndex        =   10
            Top             =   1440
            Width           =   4335
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000A&
            Caption         =   "Halt Simulator before running this test"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   2880
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Digital Output"
         Height          =   3255
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   2655
         Begin VB.CommandButton Command1 
            Caption         =   "Lamp Test"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   1560
            Width           =   1095
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   0
            Left            =   2160
            TabIndex        =   1
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   195
            Left            =   1440
            TabIndex        =   2
            Top             =   1680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   344
            _Version        =   327682
            Appearance      =   1
            Min             =   1e-4
            Max             =   10
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   1
            Left            =   1920
            TabIndex        =   4
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   2
            Left            =   1680
            TabIndex        =   5
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   9
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   495
            Index           =   3
            Left            =   1440
            TabIndex        =   6
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   873
            _Version        =   327681
            Max             =   5
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1 "
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
            Left            =   1320
            TabIndex        =   7
            Top             =   360
            Width           =   1155
         End
      End
   End
End
Attribute VB_Name = "HWTestDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Lamp Test - put all lamps ON for a few seconds

Call LampsOn
LampTestOn = True

End Sub
Private Sub LampsOn()
Dim i As Integer

For i = 0 To 5119
    'DO_Temp(i) = VB_DO(i)
    VB_DO(i) = True
Next

' PUT MECD 24V RELAY OFF
' This ensures that 24V is fed to MECD#1 and MECD#2
      VB_DO(530) = False
      VB_DO(531) = False
' Same for Instructor's (BRIDGE) 24V RELAY
     VB_DO(1836) = False
     VB_DO(1837) = False
     
' SKIP the MECD Bells
     VB_DO(608) = False
     VB_DO(609) = False
     VB_DO(620) = False
     VB_DO(621) = False
     VB_DO(534) = False
     VB_DO(535) = False
' SKIP the BRIDGE Bell
     VB_DO(1812) = False
     VB_DO(1813) = False


'VB_DO(21) = False 'dont sound alarm in lamp test
Call set_DO(VB_DO(0))

End Sub
Private Sub LampsOff()
Dim i As Integer
'restores lamps to previous state

For i = 0 To 5119
    'DO_Temp(i) = VB_DO(i)
    VB_DO(i) = False
Next
Call set_DO(VB_DO(0))

End Sub



Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Dim i As Integer

' ===== 13/11/97 =========================

'6/2/98 will need this back in
    'Call read_DItest_strings


'HWTestDigital.Caption = DItest_strings(1)
'Frame2.Caption = DItest_strings(2)
'SetDO.Caption = DItest_strings(3)
'Command1.Caption = DItest_strings(4)
'Command2.Caption = DItest_strings(5) & " " & DItest_strings(11)   ' Alarm On
'Frame3.Caption = DItest_strings(6)
'Frame1.Caption = DItest_strings(7)
'Label1.Caption = DItest_strings(8)
'Label2(2).Caption = DItest_strings(9)
'Label5.Caption = DItest_strings(10)
'Label2(1).Caption = DItest_strings(11)  ' ON
'GPIOOnOff.Caption = DItest_strings(11)  ' ON
'Label2(0).Caption = DItest_strings(12)  ' OFF

'==========================================

 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

LampTestCounter = 0
'remember the state of all DIs for comparison again next pass
For i = 0 To 3071
    DIold(i) = DI(i)
Next


'10/2/98 set defaults for DO Board scan


DIBaseDefault = &H50
DITopDefault = &H69



End Sub

Private Sub Form_Unload(Cancel As Integer)

'stop lamp test if one in progress
If LampTestOn Then
    LampTestOn = False
    Call LampsOff
    LampTestCounter = 0
End If

'stop DO test if one in progress

DOTestON = False

End Sub


Private Sub GPIOOnOff_Click()


End Sub



Private Sub SetDO_Click()


End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim StrLen As Integer

'Scan DIs to see if any changes, if so display Channel + Text

Call get_DI_Changes(DI_Changes(0), DI_States(0))
'Debug.Print "HW Digital Test DI_CHanges = " & DI_Changes(0)
'Show last Change
For i = 0 To 127
    If DI_Changes(i) <> 0 Then
        'Label3(2).Caption = "DI(" & DI_Changes(i) & ") is " & DI_States(i)
        Call reset_DI_Changes(DI_Changes(i))    'reset DI_Changes array to remove this entry
        If DI_States(i) <> 0 Then ' DI has gone true (ON)
            Label3(0).Caption = DI_Changes(i)
            'Label3(1).BackColor = &HE0E0E0 ' Grey
            
            Label3(0).BackColor = &H8080FF  'Red

        Else
           'Label3(1).Caption = DI_Changes(i)
            Label3(0).BackColor = &HE0E0E0 ' Grey
            'Label3(1).BackColor = &H80FF80 'Green


        End If
        'Debug.Print "i = " & i
        'Debug.Print "DI_Changes(i) = " & DI_Changes(i)
        'debug.Print "DITEXT(DI_Changes(i)) = " & DITEXT(DI_Changes(i))

        StrLen = Len(DITEXT(DI_Changes(i)))
        
        
        Label3(2).Caption = Mid(DITEXT(DI_Changes(i)), 2, StrLen)

    End If
Next


'Lamp Test
If LampTestOn Then
    ProgressBar1.Visible = True
    LampTestCounter = LampTestCounter + 1
    'Debug.Print "LampTestCounter = " & LampTestCounter
    If LampTestCounter > 20 Then
        LampTestOn = False
        LampTestCounter = 0
        'Debug.Print "calling lights out "
        Call LampsOff
    End If
    ProgressBar1.Value = (LampTestCounter / 21) * (ProgressBar1.max) + ProgressBar1.Min

Else
  ProgressBar1.Visible = False
End If

End Sub

Private Sub UpDown1_Change(Index As Integer)
    
DO_number = UpDown1(0).Value + UpDown1(1).Value * 10 + UpDown1(2).Value * 100 + UpDown1(3).Value * 1000

Label4.Caption = Format(UpDown1(3).Value, "0") & " " & Format(UpDown1(2).Value, "0") & " " & Format(UpDown1(1).Value, "0") & " " & Format(UpDown1(0).Value, "0") & " "
    
Call setDO
    
End Sub

Private Sub UpDown2_Change(Index As Integer)



End Sub
Sub setDO()
' 19/2/98
Dim i As Integer
Dim DONum As Integer


For i = 0 To 5119
    VB_DO(i) = False
Next
DOTestON = True


'If UpDown1.Value > 0 And UpDown1.Value <= 5119 Then
If DO_number > 0 And DO_number <= 5119 Then

'    VB_DO(UpDown1.Value) = True
     VB_DO(DO_number) = True
    
' Make sure the 24v MECD relay is ON
     
'  20/2/98  the MECD 24v relay HAS NOW DECIDED to be ENERGISED
'  when you first switch the mains ON.
'  The contacts have been re-wired so that this relay should
'  be DE_ENERGISED when you are running normally.
'  The design is not good because we do not really
'  Know the state of this relay when power is first applied
'  So we now have to remove the following:
'     VB_DO(531) = True
     
    
    
End If
'Debug.Print " VB_DO number = " & Label4.Caption
Call set_DO(VB_DO(0))

End Sub
