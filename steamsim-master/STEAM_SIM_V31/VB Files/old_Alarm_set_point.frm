VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form Alarm_set_point 
   Caption         =   "Alarm Set Point"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   6990
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   300
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5100
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   420
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   6316
         TabIndex        =   14
         Top             =   420
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327680
         BuddyControl    =   "Text1"
         BuddyDispid     =   196611
         OrigLeft        =   6240
         OrigTop         =   900
         OrigRight       =   6480
         OrigBottom      =   1755
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   180
         Top             =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   315
         Index           =   3
         Left            =   3420
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   315
         Index           =   2
         Left            =   3420
         TabIndex        =   10
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   315
         Index           =   4
         Left            =   3420
         TabIndex        =   7
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   315
         Index           =   1
         Left            =   3420
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx.x"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   2460
         TabIndex        =   13
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LOW LOW Set Point:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx.x"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   2460
         TabIndex        =   9
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LOW Set Point:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   1155
         TabIndex        =   8
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HIGH Set Point:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   1125
         TabIndex        =   6
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx.x"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   2460
         TabIndex        =   5
         Top             =   660
         Width           =   690
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx.x"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   2460
         TabIndex        =   2
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "HIGH HIGH Set Point:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   750
         TabIndex        =   1
         Top             =   360
         Width           =   1365
      End
   End
End
Attribute VB_Name = "Alarm_set_point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub refresh_values()

' 25/3/98
If nAlarm_HIHI = 0 Then
  Label1(0).Visible = False
  Label2(0).Visible = False
  Command1(0).Visible = False
Else
' HIGH Alarm Set-point
'    Label2(0).Caption = Format(ALARM_HILO(1, nAlarm_HI), "###0.00")
            If Mid$(alarmtxt(nAlarm_HIHI), 1, 4) = "%   " Then
                   Label2(0).Caption = Format(ALARM_HILO(1, nAlarm_HIHI) * 100#, "###0.00") & " " & Mid$(alarmtxt(nAlarm_HIHI), 1, 4)     ' MUST BE A TANK
            Else
                   Label2(0).Caption = Format(ALARM_HILO(1, nAlarm_HIHI), "###0.00") & " " & Mid$(alarmtxt(nAlarm_HIHI), 1, 4)                  ' IS NOT A TANK
            End If
End If

If nAlarm_HI = 0 Then
  Label1(1).Visible = False
  Label2(1).Visible = False
  Command1(1).Visible = False
Else
' HIGH Alarm Set-point
'      Label2(1).Caption = Format(ALARM_HILO(1, nAlarm_HI), "###0.00")
            If Mid$(alarmtxt(nAlarm_HI), 1, 4) = "%   " Then
                   Label2(1).Caption = Format(ALARM_HILO(1, nAlarm_HI) * 100#, "###0.00") & " " & Mid$(alarmtxt(nAlarm_HI), 1, 4)      ' MUST BE A TANK
            Else
                   Label2(1).Caption = Format(ALARM_HILO(1, nAlarm_HI), "###0.00") & " " & Mid$(alarmtxt(nAlarm_HI), 1, 4)                  ' IS NOT A TANK
            End If

End If

If nAlarm_LO = 0 Then
  Label1(2).Visible = False
  Label2(2).Visible = False
  Command1(2).Visible = False
Else
' LOW Alarm Set-point
'      Label2(2).Caption = Format(ALARM_HILO(2, nAlarm_LO), "###0.00")
            If Mid$(alarmtxt(nAlarm_LO), 1, 4) = "%   " Then
                   Label2(2).Caption = Format(ALARM_HILO(2, nAlarm_LO) * 100#, "###0.00") & " " & Mid$(alarmtxt(nAlarm_LO), 1, 4)      ' MUST BE A TANK
            Else
                   Label2(2).Caption = Format(ALARM_HILO(2, nAlarm_LO), "###0.00") & " " & Mid$(alarmtxt(nAlarm_LO), 1, 4)                 ' IS NOT A TANK
            End If
    
End If

If nAlarm_LOLO = 0 Then
  Label1(3).Visible = False
  Label2(3).Visible = False
  Command1(3).Visible = False
Else
' LOW LOW Alarm Set-point
'      Label2(3).Caption = Format(ALARM_HILO(2, nAlarm_LOLO), "###0.00")
            If Mid$(alarmtxt(nAlarm_LOLO), 1, 4) = "%   " Then
                   Label2(3).Caption = Format(ALARM_HILO(2, nAlarm_LOLO) * 100#, "###0.00") & " " & Mid$(alarmtxt(nAlarm_LOLO), 1, 4)      ' MUST BE A TANK
            Else
                   Label2(3).Caption = Format(ALARM_HILO(2, nAlarm_LOLO), "###0.00") & " " & Mid$(alarmtxt(nAlarm_LOLO), 1, 4)                  ' IS NOT A TANK
            End If
    
End If

End Sub

Private Sub Command1_Click(Index As Integer)
Dim New_set_point As String
Dim set_point As Single
Dim int1 As Integer
Dim this_alarm_text As String

' 14/1/98
   If Index = 4 Then
      Unload Me
   End If
   Control_PC_check
   If This_Control_status = False Then Exit Sub


' 7/4/98
'''int1 = DoEvents



If Index = 0 Then

' change the HIGH HIGH Set-point
again:
     On Error GoTo again:
     New_set_point = InputBox("Enter the new HIGH HIGH Set-point", "HIGH HIGH Set-point", "")
     On Error GoTo again:
     set_point = Val(New_set_point)
   If set_point < 0# Or set_point > 100000# Then
      response = MsgBox("Incorrect Entry , try again", vbOK, "Error: Value out of range")
      GoTo again:
   Else
   
            If Mid$(alarmtxt(nAlarm_HIHI), 1, 4) = "%   " Then
                    Call setHILO(nAlarm_HIHI, set_point * 0.01)       ' MUST BE A TANK
            Else
                    Call setHILO(nAlarm_HIHI, set_point)              ' IS NOT A TANK
            End If
             
        ' ====  9/6/98   =======
            If medium_speed Then
                    this_alarm_text = Mid$(alarmtxt(nAlarm_HIHI), 6, 31)
            Else
                    this_alarm_text = Mid$(alarmtxt(nAlarm_HIHI), 5, 32)
            End If

            Call Logging(" Student Changed HI HI Setpoint for Alarm No " & Format(nAlarm_HIHI, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00"), LogAlarms)
        '=== 9/6/98  end  =====
   
             
             
    End If

ElseIf Index = 1 Then

' change the HIGH Set-point
again1:
     On Error GoTo again1:
     New_set_point = InputBox("Enter the new HIGH Set-point", "High Set-point", "")
     On Error GoTo again1:
     set_point = Val(New_set_point)
   If set_point < 0# Or set_point > 100000# Then
      response = MsgBox("Incorrect Entry , try again", vbOK, "Error: Value out of range")
      GoTo again1:
   Else
   
            If Mid$(alarmtxt(nAlarm_HI), 1, 4) = "%   " Then
                    Call setHILO(nAlarm_HI, set_point * 0.01)       ' MUST BE A TANK
            Else
                    Call setHILO(nAlarm_HI, set_point)              ' IS NOT A TANK
            End If
            
        ' ====  9/6/98   =======
            If medium_speed Then
                    this_alarm_text = Mid$(alarmtxt(nAlarm_HI), 6, 31)
            Else
                    this_alarm_text = Mid$(alarmtxt(nAlarm_HI), 5, 32)
            End If

            Call Logging(" Student Changed HI Setpoint for Alarm No " & Format(nAlarm_HI, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00"), LogAlarms)
        '=== 9/6/98  end  =====
       
            
        
   End If
   
ElseIf Index = 2 Then

' change the LOW Set-point
again2:
     On Error GoTo again2
     New_set_point = InputBox("Enter the new LOW Set-point", "LOW Set-point", "")
     On Error GoTo again2
     set_point = Val(New_set_point)
     
  ' 7/7/98
  'If set_point < 0# Or set_point > 100000# Then
   If set_point > 100000# Then
   
      response = MsgBox("Incorrect Entry , try again", vbOK, "Error: Value out of range")
      GoTo again2
   Else
   
            If Mid$(alarmtxt(nAlarm_LO), 1, 4) = "%   " Then
                       Call setHILO(nAlarm_LO, set_point * 0.01)            ' MUST BE A TANK
            Else
                        Call setHILO(nAlarm_LO, set_point)                  ' IS NOT A TANK
            End If
            
        ' ====  9/6/98   =======
            'alarm_no
            If medium_speed Then
                    this_alarm_text = Mid$(alarmtxt(nAlarm_LO), 6, 31)
            Else
                    this_alarm_text = Mid$(alarmtxt(nAlarm_LO), 5, 32)
            End If

            Call Logging("Student Changed LOW Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00"), LogAlarms)
        '=== 9/6/98  end  =====
           
            
             
   End If
   
   
ElseIf Index = 3 Then

' change the LOW LOW Set-point
again3:
     On Error GoTo again3:
     New_set_point = InputBox("Enter the new LOW LOW Set-point", "LOW LOW Set-point", "")
     On Error GoTo again3:
     set_point = Val(New_set_point)
     
  ' 7/7/98
  'If set_point < 0# Or set_point > 100000# Then
   If set_point > 100000# Then
   
      response = MsgBox("Incorrect Entry , try again", vbOK, "Error: Value out of range")
      GoTo again3:
   Else
   
            If Mid$(alarmtxt(nAlarm_LOLO), 1, 4) = "%   " Then
                       Call setHILO(nAlarm_LOLO, set_point * 0.01)            ' MUST BE A TANK
            Else
                        Call setHILO(nAlarm_LOLO, set_point)                  ' IS NOT A TANK
            End If
            
        ' ====  9/6/98   =======
            If medium_speed Then
                    this_alarm_text = Mid$(alarmtxt(nAlarm_LOLO), 6, 31)
            Else
                    this_alarm_text = Mid$(alarmtxt(nAlarm_LOLO), 5, 32)
            End If

            Call Logging("Student Changed LOW LOW Setpoint for Alarm No " & Format(nAlarm_LOLO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00"), LogAlarms)
        '=== 9/6/98  end  =====
            
            
             
   End If
   


ElseIf Index = 4 Then

  Unload Me
  
End If


End Sub

Private Sub Form_Load()

' 8/4/98
Me.Height = 2940
Me.Width = 7110

Call refresh_values

End Sub

Private Sub Form_Unload(Cancel As Integer)

' 30/12/97
nAlarm_HIHI = 0
nAlarm_HI = 0
nAlarm_LO = 0
nAlarm_LOLO = 0

End Sub

Private Sub Timer1_Timer()
''''Dim i As Integer

''''i = DoEvents

Call refresh_values

End Sub
