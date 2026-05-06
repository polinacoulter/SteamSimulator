VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form Throttle_Trips 
   Caption         =   "Throttle Trips and Cut-Backs"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   8655
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4335
      Left            =   0
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   7646
      Begin VB.CommandButton AUTO 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7140
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton OFF 
         BackColor       =   &H008080FF&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6600
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton STBD_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STBD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6780
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton BOTH_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BOTH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5700
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton PORT_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4620
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton MANOUVER_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MANOEUVER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6180
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton NORMAL_cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NORMAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4140
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton AUTO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2220
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton AUTO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2760
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton OFF 
         BackColor       =   &H008080FF&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton AUTO 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   180
         Top             =   960
      End
      Begin VB.Label Speed_set_rpm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set Point     100 rpm Astern"
         Height          =   315
         Left            =   6060
         TabIndex        =   31
         Top             =   3300
         Width           =   2175
      End
      Begin VB.Label GOV_pos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Output        0 %"
         Height          =   315
         Left            =   6060
         TabIndex        =   30
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Feedback"
         Height          =   255
         Left            =   6300
         TabIndex        =   29
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Boiler Operating Mode"
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   1740
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Throttle Operating Mode"
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VIAX Cut-Back"
         Height          =   255
         Left            =   3900
         TabIndex        =   18
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Boiler Pressure Cut-Back"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Throttle Trips"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drum Level  Cut-Back"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turbine Overspeed"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   0
         Left            =   1500
         TabIndex        =   8
         Top             =   600
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    VAC   Low"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Condenser Level High"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   2
         Left            =   3300
         TabIndex        =   6
         Top             =   600
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lube Oil Pressure Low"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   3
         Left            =   1500
         TabIndex        =   5
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Boiler Pressure Low"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   4
         Left            =   2400
         TabIndex        =   4
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valve Over-Travel"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   5
         Left            =   3300
         TabIndex        =   3
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Main Throttle Open"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   6
         Left            =   1500
         TabIndex        =   2
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Boiler Drum Level"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   7
         Left            =   2400
         TabIndex        =   1
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label TRIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VIAX High"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   8
         Left            =   3300
         TabIndex        =   0
         Top             =   1920
         Width           =   915
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1020
      TabIndex        =   19
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "Throttle_Trips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private set_dummy_focus As Integer

Private Sub AUTO_Click(Index As Integer)

' 13/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

If Index = 0 Then
  Call setPCDI(313)
ElseIf Index = 1 Then
  Call setPCDI(315)
ElseIf Index = 2 Then
  Call setPCDI(317)
  
ElseIf Index = 3 Then   ' === 05/03/02 Speed Feedback AUTO ===
  Call setPCDI(295)

End If



End Sub

Private Sub Command2_Click()

End Sub

Private Sub BOTH_cmd_Click()

' 14/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

  Call setPCDI(323)



End Sub

Private Sub Form_Load()

' 13/9/00
On Error Resume Next
Call JYAScreenInit(Me, idThrottle_trips)

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)



End Sub

Private Sub Form_Resize()


' 13/9/00
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idThrottle_trips)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub MANOUVER_cmd_Click()

' 14/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

  Call setPCDI(320)



End Sub

Private Sub NORMAL_cmd_Click()

' 14/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

  Call setPCDI(321)


End Sub

Private Sub OFF_Click(Index As Integer)


' 13/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

If Index = 0 Then
  Call setPCDI(314)
ElseIf Index = 1 Then
  Call setPCDI(316)
ElseIf Index = 2 Then
  Call setPCDI(318)
  
ElseIf Index = 3 Then   ' === 05/03/02 Speed Feedback ===
  Call setPCDI(294)
  
End If



End Sub

Private Sub PORT_cmd_Click()

' 14/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


' 6/12/00
' Call setPCDI(324)
  Call setPCDI(322)



End Sub

Private Sub STBD_cmd_Click()


' 14/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

' 6/12/00
' Call setPCDI(322)
  Call setPCDI(324)


End Sub

Private Sub Timer1_Timer()
Dim i As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' === 22/11/02 17.049 ===
Flashing_Flag = Not Flashing_Flag

' === 05/03/02 ===
GOV_pos.Caption = "Output        " & Format(LSS_reals(841), "##0.0") & " %"

' PC_REALS(842) = SPEED_set_rpm
If LSS_reals(842) > 0.5 Then
    Speed_set_rpm.Caption = "Set Point     " & Format(Abs(LSS_reals(842)), "##0") & " rpm Ahead"
ElseIf LSS_reals(842) < -0.5 Then
    Speed_set_rpm.Caption = "Set Point     " & Format(Abs(LSS_reals(842)), "##0") & " rpm Astern"
Else
    Speed_set_rpm.Caption = "Set Point     " & Format(LSS_reals(842), "##0") & "           "
End If


' 13/9/00
For i = 0 To 8
  If VB_steam_logicals(139 + i) Then
    TRIP(i).BackColor = &HFF            ' Red
  Else
    TRIP(i).BackColor = &H808080     ' grey
  End If
Next i

'c   === 13/9/00 ===
'    equivalence(STEAM_logicals(148),Throttle_CUT_BACK_AUTO(1))
'    logical*1           Throttle_CUT_BACK_AUTO(3)
' 13/9/00
For i = 0 To 2
  If VB_steam_logicals(148 + i) Then
            OFF(i).BackColor = &HC0C0C0          ' grey
            AUTO(i).BackColor = &HFFFF&          ' Yellow
  Else
       ' === 22/11/02 ===
            'OFF(i).BackColor = &H8080FF           ' Pink
            'AUTO(i).BackColor = &HC0C0C0          ' grey
     If (LSS_LOGICALS(22 + i)) Then    ' FLASHING...
              OFF(i).BackColor = &HC0C0C0          ' grey
        If Flashing_Flag Then
             AUTO(i).BackColor = &HFFFF&           ' Yellow
        Else
             AUTO(i).BackColor = &HC0C0C0          ' grey
        End If
     Else
              OFF(i).BackColor = &H8080FF          ' Pink
             AUTO(i).BackColor = &HC0C0C0          ' grey
     End If
     
  End If
Next i

' === 05/03/02 ======
'c             === 05/03/02 SPEED_Feedback_AUTO_DO ===
'             LSS_LOGICALS(99) = DO(1035)
If LSS_LOGICALS(99) Then
    OFF(3).BackColor = &HC0C0C0          ' grey
    AUTO(3).BackColor = &HFFFF&        ' Yellow
    GOV_pos.Visible = True
    Speed_set_rpm.Visible = True
Else
    OFF(3).BackColor = &H8080FF          ' Pink
    AUTO(3).BackColor = &HC0C0C0          ' grey
    GOV_pos.Visible = False
    Speed_set_rpm.Visible = False
End If
' === 05/03/02 END ===



' 14/9/00
'c equivalence(STEAM_logicals(151), MANOUVER_mode)
'c   logical*1           MANOUVER_mode
'c equivalence(STEAM_logicals(152), NORMAL_mode)
'c   logical*1           NORMAL_mode
'
'c equivalence(STEAM_logicals(153), BOILER_OP_PORT)
'c   logical*1           BOILER_OP_PORT
'c equivalence(STEAM_logicals(154), BOILER_OP_BOTH)
'c   logical*1           BOILER_OP_BOTH
'c equivalence(STEAM_logicals(155), BOILER_OP_STBD)
'c   logical*1           BOILER_OP_STBD
If VB_steam_logicals(151) Then
    MANOUVER_cmd.BackColor = &HFF00&           ' Green
Else
    MANOUVER_cmd.BackColor = &HC0C0C0           ' grey
End If
If VB_steam_logicals(152) Then
    NORMAL_cmd.BackColor = &HFF00&           ' Green
Else
    NORMAL_cmd.BackColor = &HC0C0C0          ' grey
End If

If VB_steam_logicals(153) Then
    PORT_cmd.BackColor = &HFF00&           ' Green
Else
    PORT_cmd.BackColor = &HC0C0C0           ' grey
End If
If VB_steam_logicals(154) Then
    BOTH_cmd.BackColor = &HFF00&           ' Green
Else
    BOTH_cmd.BackColor = &HC0C0C0           ' grey
End If
If VB_steam_logicals(155) Then
    STBD_cmd.BackColor = &HFF00&           ' Green
Else
    STBD_cmd.BackColor = &HC0C0C0          ' grey
End If



End Sub

Private Sub TRIP_Click(Index As Integer)

' === 12/12/01 ===
Select Case Index

   Case 0   ' 14.108 Turbine Overspeed
   
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 148
    Alarm_set_point.Show
    
   Case 1  ' 14.108 Vac Low
   
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 149
    Alarm_set_point.Show

   Case 2   ' 14.108 Condenser Level High
   
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 150
    Alarm_set_point.Show
    
   Case 3   ' 14.108 Lube Oil Pump Discharge pressure LOW  20 psi
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 151
    Alarm_set_point.Show
   
   Case 4   ' 14.108 Boiler Pressure LOW  600 psi
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 152
    Alarm_set_point.Show
   
   Case 7   ' 14.108 Boiler Drum Level HIGH
    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 155
    Alarm_set_point.Show
        
   
End Select

End Sub
