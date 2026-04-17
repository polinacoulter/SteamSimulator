VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form Alarm_set_point 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Set Point"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   9765
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   3375
      Left            =   0
      Top             =   60
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5953
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   0
         TabIndex        =   0
         Top             =   60
         Width           =   9255
         Begin VB.CommandButton Command1 
            Caption         =   "Exit"
            Height          =   375
            Index           =   4
            Left            =   5580
            TabIndex        =   29
            Top             =   2700
            Width           =   1215
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   120
            Top             =   2640
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   60
            Picture         =   "Alarm_set_point.frx":0000
            ScaleHeight     =   555
            ScaleWidth      =   615
            TabIndex        =   28
            Top             =   120
            Width           =   615
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Current"
            ForeColor       =   &H000000FF&
            Height          =   2055
            Index           =   0
            Left            =   3360
            TabIndex        =   23
            Top             =   540
            Width           =   1155
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Width           =   225
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   900
               Width           =   225
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   25
               Top             =   1320
               Width           =   225
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   24
               Top             =   1680
               Width           =   225
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Default"
            Height          =   2055
            Left            =   2100
            TabIndex        =   18
            Top             =   540
            Width           =   1155
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   22
               Top             =   1680
               Width           =   225
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   225
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   900
               Width           =   225
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "   0"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   19
               Top             =   1320
               Width           =   225
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "New"
            Height          =   2055
            Left            =   4620
            TabIndex        =   1
            Top             =   540
            Width           =   4335
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   13
               Top             =   1620
               Width           =   1095
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   12
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   11
               Top             =   780
               Width           =   1095
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Change"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   1680
               TabIndex        =   9
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Change"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   1680
               TabIndex        =   8
               Top             =   780
               Width           =   1215
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Change"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   1680
               TabIndex        =   7
               Top             =   1200
               Width           =   1215
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Change"
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   1680
               TabIndex        =   6
               Top             =   1620
               Width           =   1215
            End
            Begin VB.CommandButton FORCE_default 
               Caption         =   "Default"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   3000
               TabIndex        =   5
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton FORCE_default 
               Caption         =   "Default"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   3000
               TabIndex        =   4
               Top             =   780
               Width           =   1215
            End
            Begin VB.CommandButton FORCE_default 
               Caption         =   "Default"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   3000
               TabIndex        =   3
               Top             =   1200
               Width           =   1215
            End
            Begin VB.CommandButton FORCE_default 
               Caption         =   "Default"
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   3000
               TabIndex        =   2
               Top             =   1620
               Width           =   1215
            End
            Begin ComCtl2.UpDown UpDown2 
               Height          =   375
               Index           =   0
               Left            =   1320
               TabIndex        =   14
               Top             =   360
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   327681
               OrigLeft        =   6300
               OrigTop         =   1860
               OrigRight       =   6540
               OrigBottom      =   2235
               Enabled         =   -1  'True
            End
            Begin ComCtl2.UpDown UpDown2 
               Height          =   375
               Index           =   1
               Left            =   1320
               TabIndex        =   15
               Top             =   780
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   327681
               OrigLeft        =   6300
               OrigTop         =   1860
               OrigRight       =   6540
               OrigBottom      =   2235
               Enabled         =   -1  'True
            End
            Begin ComCtl2.UpDown UpDown2 
               Height          =   375
               Index           =   2
               Left            =   1320
               TabIndex        =   16
               Top             =   1200
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   327681
               OrigLeft        =   6300
               OrigTop         =   1860
               OrigRight       =   6540
               OrigBottom      =   2235
               Enabled         =   -1  'True
            End
            Begin ComCtl2.UpDown UpDown2 
               Height          =   375
               Index           =   3
               Left            =   1320
               TabIndex        =   17
               Top             =   1620
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   327681
               OrigLeft        =   6300
               OrigTop         =   1860
               OrigRight       =   6540
               OrigBottom      =   2235
               Enabled         =   -1  'True
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HIGH HIGH Set Point:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   -60
            TabIndex        =   34
            Top             =   960
            Width           =   1830
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HIGH Set Point:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   33
            Top             =   1380
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LOW Set Point:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   660
            TabIndex        =   32
            Top             =   1800
            Width           =   1110
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LOW LOW Set Point:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   31
            Top             =   2220
            Width           =   1530
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   780
            TabIndex        =   30
            Top             =   180
            Width           =   3135
         End
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
Dim UnitStr As String
Dim SetPoint As Single
Dim AlarmSetPoint As Single


Debug.Print "alarm_HILO(1,98) = " & ALARM_HILO(1, 98)
Debug.Print "alarm_HILO(2,98) = " & ALARM_HILO(2, 98)

'Text2.text = UpDown2.Value & " " & UnitStr

'--------------------------------- hi hi -------------------------------------
If nAlarm_HIHI = 0 Then
  Label1(0).Visible = False
  Label2(0).Visible = False
  Command1(0).Visible = False
  ' 2/12/00
  FORCE_default(0).Visible = False
Else


    UnitStr = Mid$(alarmtxt(nAlarm_HIHI), 6, 4)
    AlarmSetPoint = ALARM_HILO(1, nAlarm_HIHI)
    If SI_units Then 'convert units to SI
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint = AlarmSetPoint / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint = (AlarmSetPoint - 32) * 5# / 9#
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint = AlarmSetPoint * 25.4
            End If
    Else
            SetPoint = AlarmSetPoint
    End If
    
    If InStr(1, UnitStr, "%", 1) > 0 Then
            ' 7/1/00
          ' Label2(0).Caption = Format(SetPoint * 100#, "###0.00") & " " & UnitStr     ' MUST BE A TANK
            Label2(0).Caption = Format(SetPoint, "###0.00") & " " & UnitStr     ' MUST BE A TANK
    Else
            Label2(0).Caption = Format(SetPoint, "###0.00") & " " & UnitStr                  ' IS NOT A TANK
    End If
End If


'--------------------------------- hi -------------------------------------


If nAlarm_HI = 0 Then
  Label1(1).Visible = False
  Label2(1).Visible = False
  Command1(1).Visible = False
 ' 2/12/00
  FORCE_default(1).Visible = False

Else
    UnitStr = Mid$(alarmtxt(nAlarm_HI), 6, 4)
    AlarmSetPoint = ALARM_HILO(1, nAlarm_HI)
    If SI_units Then 'convert units to SI
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint = AlarmSetPoint / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint = (AlarmSetPoint - 32) * 5# / 9#
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint = AlarmSetPoint * 25.4
            End If
    Else
            SetPoint = AlarmSetPoint
    End If
    
    If InStr(1, UnitStr, "%", 1) > 0 Then
            ' 7/1/00
         '  Label2(1).Caption = Format(SetPoint * 100#, "###0.00") & " " & UnitStr     ' MUST BE A TANK
            Label2(1).Caption = Format(SetPoint, "###0.00") & " " & UnitStr      ' MUST BE A TANK
    Else
            Label2(1).Caption = Format(SetPoint, "###0.00") & " " & UnitStr                  ' IS NOT A TANK
    End If

End If

'--------------------------------- lO -------------------------------------


If nAlarm_LO = 0 Then
  Label1(2).Visible = False
  Label2(2).Visible = False
  Command1(2).Visible = False
  ' 2/12/00
  FORCE_default(2).Visible = False

Else
    UnitStr = Mid$(alarmtxt(nAlarm_LO), 6, 4)
    AlarmSetPoint = ALARM_HILO(2, nAlarm_LO)
    If SI_units Then 'convert units to SI
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint = AlarmSetPoint / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint = (AlarmSetPoint - 32) * 5# / 9#
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint = AlarmSetPoint * 25.4
            End If
    Else
            SetPoint = AlarmSetPoint
    End If
    
    If InStr(1, UnitStr, "%", 1) > 0 Then
            ' 7/1/00
          ' Label2(2).Caption = Format(SetPoint * 100#, "###0.00") & " " & UnitStr     ' MUST BE A TANK
            Label2(2).Caption = Format(SetPoint, "###0.00") & " " & UnitStr      ' MUST BE A TANK
    Else
            Label2(2).Caption = Format(SetPoint, "###0.00") & " " & UnitStr                  ' IS NOT A TANK
    End If
End If

'--------------------------------- LO LO  -------------------------------------


If nAlarm_LOLO = 0 Then
  Label1(3).Visible = False
  Label2(3).Visible = False
  Command1(3).Visible = False
  ' 2/12/00
  FORCE_default(3).Visible = False

Else
UnitStr = Mid$(alarmtxt(nAlarm_LOLO), 6, 4)
    AlarmSetPoint = ALARM_HILO(2, nAlarm_LOLO)
    If SI_units Then 'convert units to SI
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint = AlarmSetPoint / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint = (AlarmSetPoint - 32) * 5# / 9#
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint = AlarmSetPoint * 25.4
            End If
    Else
            SetPoint = AlarmSetPoint
    End If
    
    If InStr(1, UnitStr, "%", 1) > 0 Then
            ' 7/1/00
           'Label2(3).Caption = Format(SetPoint * 100#, "###0.00") & " " & UnitStr     ' MUST BE A TANK
            Label2(3).Caption = Format(SetPoint, "###0.00") & " " & UnitStr      ' MUST BE A TANK
    Else
            Label2(3).Caption = Format(SetPoint, "###0.00") & " " & UnitStr                  ' IS NOT A TANK
    End If
    
End If

End Sub

Private Sub Command1_Click(Index As Integer)
'15/10/99
Dim New_set_point As String
Dim set_point As Single
Dim SetPoint As Single

Dim int1 As Integer
Dim this_alarm_text As String
Dim UnitStr As String
Dim AlarmSetPoint As Single

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

            set_point = UpDown2(0).Value / SetPointMult(0)
            UnitStr = Mid$(alarmtxt(nAlarm_HIHI), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
            
              ' 7/1/00
               '    Call setHILO(nAlarm_HIHI, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                    Call setHILO(nAlarm_HIHI, set_point)       ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_HIHI, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_HIHI, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_HIHI, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_HIHI, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_HIHI), 10, Len(alarmtxt(nAlarm_HIHI)))
            Call Logging("Student Changed HI-HI Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)

ElseIf Index = 1 Then

' change the HIGH Set-point
            set_point = UpDown2(1).Value / SetPointMult(1)
            UnitStr = Mid$(alarmtxt(nAlarm_HI), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
                  ' 7/1/00
                  'Call setHILO(nAlarm_HI, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                   Call setHILO(nAlarm_HI, set_point)       ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_HI, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_HI, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_HI, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_HI, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_HI), 10, Len(alarmtxt(nAlarm_HI)))
            Call Logging("Student Changed HI Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)
   
ElseIf Index = 2 Then


            set_point = UpDown2(2).Value / SetPointMult(2)
            UnitStr = Mid$(alarmtxt(nAlarm_LO), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
                   ' 7/1/00
                 ' Call setHILO(nAlarm_LO, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                   Call setHILO(nAlarm_LO, set_point)        ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_LO, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_LO, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_LO, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_LO, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_LO), 10, Len(alarmtxt(nAlarm_LO)))
            Call Logging("Student Changed LO Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)
   
ElseIf Index = 3 Then

' change the LOW LOW Set-point


            set_point = UpDown2(3).Value / SetPointMult(3)
            UnitStr = Mid$(alarmtxt(nAlarm_LOLO), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
            
                     ' 7/1/00
                  ' Call setHILO(nAlarm_LOLO, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                    Call setHILO(nAlarm_LOLO, set_point)        ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_LOLO, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_LOLO, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_LOLO, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_LOLO, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_LOLO), 10, Len(alarmtxt(nAlarm_LOLO)))
            Call Logging("Student Changed LO-LO Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)

ElseIf Index = 4 Then

  Unload Me
  
End If


End Sub
Sub JFSetPointRefresh()
Dim UnitStr As String
Dim SetPoint(0 To 3) As Single
Dim AlarmSetPoint(0 To 3) As Single
Dim i As Integer
Dim Units As String



If nAlarm_HIHI <> 0 Then
    UnitStr = Mid$(alarmtxt(nAlarm_HIHI), 6, 4)
    AlarmSetPoint(0) = ALARM_HILO(1, nAlarm_HIHI)
End If

If nAlarm_HI <> 0 Then
    UnitStr = Mid$(alarmtxt(nAlarm_HI), 6, 4)
    AlarmSetPoint(1) = ALARM_HILO(1, nAlarm_HI)
End If

If nAlarm_LO <> 0 Then
    UnitStr = Mid$(alarmtxt(nAlarm_LO), 6, 4)
    AlarmSetPoint(2) = ALARM_HILO(2, nAlarm_LO)
End If

If nAlarm_LOLO <> 0 Then
    UnitStr = Mid$(alarmtxt(nAlarm_LOLO), 6, 4)
    AlarmSetPoint(3) = ALARM_HILO(2, nAlarm_LOLO)
End If

For i = 0 To 3
    If SI_units Then 'convert units to SI before displaying
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                Units = "bar "
                SetPoint(i) = AlarmSetPoint(i) / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                Units = "deg C"
                SetPoint(i) = (AlarmSetPoint(i) - 32) * 5# / 9#
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                Units = "mm"
                SetPoint(i) = (AlarmSetPoint(i) * 25.4)
                
                
            ' === 21/07/09 ADDED the next 3 lines: ===
            Else
                SetPoint(i) = AlarmSetPoint(i)
                Units = UnitStr
            '=====================================
                
            End If
    Else
            SetPoint(i) = AlarmSetPoint(i)
            Units = UnitStr
    End If
    
    If InStr(1, UnitStr, "%", 1) > 0 Then
        ' 7/1/00
          ' Label2(i).Caption = Format(SetPoint(i) * 100#, "###0.00") & " " & Units     ' MUST BE A TANK
            Label2(i).Caption = Format(SetPoint(i), "###0.00") & " " & Units      ' MUST BE A TANK

    Else
            Label2(i).Caption = Format(SetPoint(i), "###0.00") & " " & Units                  ' IS NOT A TANK
    End If
    Text2(i).text = Format(UpDown2(i).Value / SetPointMult(i), "###0.00") & " " & Units


Next

End Sub
Sub InitializeAlarmSetPoints()
Dim SetPoint(0 To 3) As Single
Dim AlarmSetPoint As Single
Dim DefaultSetPoint(0 To 3) As Single
Dim DSP As Single
Dim Mytext As String
Dim i As Integer
Dim UnitStr As String
Dim SP(0 To 3) As Single

' 7/1/00
Dim setpoint_val As Integer

' 8/4/98


'Call refresh_values
For i = 0 To 3
    Label1(i).Visible = False
    Label2(i).Visible = False
    Text2(i).Visible = False
    UpDown2(i).Visible = False
    Command1(i).Visible = False
  ' 2/12/00
  FORCE_default(i).Visible = False


Next

If nAlarm_HIHI <> 0 Then
    Mytext = Mid$(alarmtxt(nAlarm_HIHI), 10, Len(alarmtxt(nAlarm_HIHI)))
    UnitStr = Mid$(alarmtxt(nAlarm_HIHI), 6, 4)
    AlarmSetPoint = ALARM_HILO(1, nAlarm_HIHI)
    DSP = Default_ALARM_HILO(1, nAlarm_HIHI)
    If SI_units Then 'convert units to SI before displaying
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint(0) = AlarmSetPoint / 14.5
                DefaultSetPoint(0) = DSP / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint(0) = (AlarmSetPoint - 32) * 5# / 9#
                DefaultSetPoint(0) = (DSP - 32) * 5 / 9
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint(0) = (AlarmSetPoint * 25.4)
                DefaultSetPoint(0) = (DSP * 25.4)
            '14/02/00 Tank levels reading 0 in SI units
            ElseIf InStr(1, UnitStr, "%", 1) > 0 Then 'No conversion for %
                UnitStr = "%"
                SetPoint(0) = (AlarmSetPoint)
                DefaultSetPoint(0) = (DSP)

            End If
    Else
            SetPoint(0) = AlarmSetPoint
            DefaultSetPoint(0) = DSP
    End If
    
            ' 7/1/00
    'If InStr(1, UnitStr, "%", 1) > 0 Then
    '    SetPoint(0) = SetPoint(0) * 100
    '    DefaultSetPoint(0) = DefaultSetPoint(0) * 100
    'End If
    
    Label4(0).Caption = Format(DefaultSetPoint(0), "###00.00") & " " & UnitStr
    
   '  29/4/00 Had to do a similar thing in Warsash...
   ' UpDown2(0).max = DefaultSetPoint(0) * 2
   ' UpDown2(0).Min = DefaultSetPoint(0) / 2
   
   ' === 26/09/02 ===
   '  UpDown2(0).max = DefaultSetPoint(0) * 5
   '  UpDown2(0).Min = DefaultSetPoint(0) / 5
      UpDown2(0).max = DefaultSetPoint(0) * 50
      UpDown2(0).Min = DefaultSetPoint(0) / 50

  ' 7/1/00  DO NOT CRASH THE PROGRAM...
    If ((SetPoint(0) <= UpDown2(0).max) And (SetPoint(0) >= UpDown2(0).Min)) Then
      UpDown2(0).Value = SetPoint(0)
    End If
   
    Label1(0).Visible = True
    Label2(0).Visible = True
    Text2(0).Visible = True
    UpDown2(0).Visible = True
    Command1(0).Visible = True
  ' 2/12/00
  FORCE_default(0).Visible = True

End If

If nAlarm_HI <> 0 Then
    Mytext = Mid$(alarmtxt(nAlarm_HI), 10, Len(alarmtxt(nAlarm_HI)))
    UnitStr = Mid$(alarmtxt(nAlarm_HI), 6, 4)
    AlarmSetPoint = ALARM_HILO(1, nAlarm_HI)
    DSP = Default_ALARM_HILO(1, nAlarm_HI)

    If SI_units Then 'convert units to SI before displaying
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint(1) = AlarmSetPoint / 14.5
                DefaultSetPoint(1) = DSP / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint(1) = (AlarmSetPoint - 32) * 5# / 9#
                DefaultSetPoint(1) = (DSP - 32) * 5 / 9
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint(1) = (AlarmSetPoint * 25.4)
                DefaultSetPoint(1) = (DSP * 25.4)
                
             ' 21/07/09 Tank levels reading 0 in SI units
            ElseIf InStr(1, UnitStr, "%", 1) > 0 Then 'No conversion for %
                UnitStr = "%"
                SetPoint(1) = AlarmSetPoint
                DefaultSetPoint(1) = DSP
                
            End If
    Else
            SetPoint(1) = AlarmSetPoint
            DefaultSetPoint(1) = DSP
    End If
    
            ' 7/1/00
    'If InStr(1, UnitStr, "%", 1) > 0 Then
    '    SetPoint(1) = SetPoint(1) * 100
    '    DefaultSetPoint(1) = DefaultSetPoint(1) * 100
    'End If
    
    Label4(1).Caption = Format(DefaultSetPoint(1), "###00.00") & " " & UnitStr
  
  '  29/4/00 Had to do a similar thing in Warsash...
  '  UpDown2(1).max = DefaultSetPoint(1) * 2
  '  UpDown2(1).Min = DefaultSetPoint(1) / 2
  
   ' === 26/09/02 ===
   '  UpDown2(1).max = DefaultSetPoint(1) * 5
   '  UpDown2(1).Min = DefaultSetPoint(1) / 5
      UpDown2(1).max = DefaultSetPoint(1) * 50
      UpDown2(1).Min = DefaultSetPoint(1) / 50

  ' 7/1/00  DO NOT CRASH THE PROGRAM...
    If ((SetPoint(1) <= UpDown2(1).max) And (SetPoint(1) >= UpDown2(1).Min)) Then

      UpDown2(1).Value = SetPoint(1)
      
    End If
    
    Label1(1).Visible = True
    Label2(1).Visible = True
    Text2(1).Visible = True
    UpDown2(1).Visible = True
    Command1(1).Visible = True
  ' 2/12/00
  FORCE_default(1).Visible = True

End If

If nAlarm_LO <> 0 Then
    Mytext = Mid$(alarmtxt(nAlarm_LO), 10, Len(alarmtxt(nAlarm_LO)))
    UnitStr = Mid$(alarmtxt(nAlarm_LO), 6, 4)
    AlarmSetPoint = ALARM_HILO(2, nAlarm_LO)
    DSP = Default_ALARM_HILO(2, nAlarm_LO)

    If SI_units Then 'convert units to SI before displaying
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint(2) = AlarmSetPoint / 14.5
                DefaultSetPoint(2) = DSP / 14.5
                Debug.Print "DSP = " & DSP & "DefaultSetPoint(2) = " & DefaultSetPoint(2)
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint(2) = (AlarmSetPoint - 32) * 5# / 9#
                DefaultSetPoint(2) = (DSP - 32) * 5 / 9
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint(2) = (AlarmSetPoint * 25.4)
                DefaultSetPoint(2) = (DSP * 25.4)
                
             ' 21/07/09 Tank levels reading 0 in SI units
            ElseIf InStr(1, UnitStr, "%", 1) > 0 Then 'No conversion for %
                UnitStr = "%"
                SetPoint(2) = AlarmSetPoint
                DefaultSetPoint(2) = DSP
                                               
            End If
    Else
            SetPoint(2) = AlarmSetPoint
            DefaultSetPoint(2) = DSP
    End If
    
            ' 7/1/00
    'If InStr(1, UnitStr, "%", 1) > 0 Then
    '    SetPoint(2) = SetPoint(2) * 100
    '    DefaultSetPoint(2) = DefaultSetPoint(2) * 100
    'End If
    
    Label4(2).Caption = Format(DefaultSetPoint(2), "###00.00") & " " & UnitStr
  
  '  29/4/00 Had to do a similar thing in Warsash...
  '  UpDown2(2).max = DefaultSetPoint(2) * 2
  '  UpDown2(2).Min = DefaultSetPoint(2) / 2
  
   ' === 26/09/02 ===
   '  UpDown2(2).max = DefaultSetPoint(2) * 5
   '  UpDown2(2).Min = DefaultSetPoint(2) / 5
      UpDown2(2).max = DefaultSetPoint(2) * 50
      UpDown2(2).Min = DefaultSetPoint(2) / 50
    
      ' 7/1/00  DO NOT CRASH THE PROGRAM...
    If ((SetPoint(2) <= UpDown2(2).max) And (SetPoint(2) >= UpDown2(2).Min)) Then
    
      UpDown2(2).Value = SetPoint(2)
      
    End If
    
    Label1(2).Visible = True
    Label2(2).Visible = True
    Text2(2).Visible = True
    UpDown2(2).Visible = True
    Command1(2).Visible = True
  ' 2/12/00
  FORCE_default(2).Visible = True


End If

If nAlarm_LOLO <> 0 Then
    Mytext = Mid$(alarmtxt(nAlarm_LOLO), 10, Len(alarmtxt(nAlarm_LOLO)))
    UnitStr = Mid$(alarmtxt(nAlarm_LOLO), 6, 4)
    AlarmSetPoint = ALARM_HILO(2, nAlarm_LOLO)
    DSP = Default_ALARM_HILO(2, nAlarm_LOLO)

    If SI_units Then 'convert units to SI before displaying
            If InStr(1, UnitStr, "psi", 1) > 0 Then
                UnitStr = "bar "
                SetPoint(3) = AlarmSetPoint / 14.5
                DefaultSetPoint(3) = DSP / 14.5
            ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                UnitStr = "deg C"
                SetPoint(3) = (AlarmSetPoint - 32) * 5# / 9#
                DefaultSetPoint(3) = (DSP - 32) * 5 / 9
            ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                UnitStr = "mm"
                SetPoint(3) = (AlarmSetPoint * 25.4)
                DefaultSetPoint(3) = (DSP * 25.4)
                
                 
             ' 21/07/09 Tank levels reading 0 in SI units
            ElseIf InStr(1, UnitStr, "%", 1) > 0 Then 'No conversion for %
                UnitStr = "%"
                SetPoint(3) = AlarmSetPoint
                DefaultSetPoint(3) = DSP
               
                
            End If
    Else
            SetPoint(3) = AlarmSetPoint
            DefaultSetPoint(3) = DSP
    End If
    
            ' 7/1/00
    'If InStr(1, UnitStr, "%", 1) > 0 Then
    '    SetPoint(3) = SetPoint(3) * 100
    '    DefaultSetPoint(3) = DefaultSetPoint(3) * 100
    'End If
    
    Label4(3).Caption = Format(DefaultSetPoint(3), "###00.00") & " " & UnitStr
  
   '  29/4/00 Had to do a similar thing in Warsash...
   ' UpDown2(3).max = DefaultSetPoint(3) * 2
   ' UpDown2(3).Min = DefaultSetPoint(3) / 2
   
   ' === 26/09/02 ===
   '  UpDown2(3).max = DefaultSetPoint(3) * 5
   '  UpDown2(3).Min = DefaultSetPoint(3) / 5
      UpDown2(3).max = DefaultSetPoint(3) * 50
      UpDown2(3).Min = DefaultSetPoint(3) / 50


  ' 7/1/00  DO NOT CRASH THE PROGRAM...
    If ((SetPoint(3) <= UpDown2(3).max) And (SetPoint(3) >= UpDown2(3).Min)) Then

      UpDown2(3).Value = SetPoint(3)
      
    End If
    
    
    Label1(3).Visible = True
    Label2(3).Visible = True
    Text2(3).Visible = True
    UpDown2(3).Visible = True
    Command1(3).Visible = True
  ' 2/12/00
  FORCE_default(3).Visible = True


End If

'Set Increment depending on value

For i = 0 To 3
SetPointMult(i) = 1

If SetPoint(i) < 0 Then
    SP(i) = -SetPoint(i)
Else
    SP(i) = SetPoint(i)
End If
If SP(i) < 1 Then
    UpDown2(i).Increment = 1
    SetPointMult(i) = 100
ElseIf SP(i) < 10 Then
    UpDown2(i).Increment = 1
    SetPointMult(i) = 10


ElseIf SP(i) < 200 Then
    UpDown2(i).Increment = 1
    SetPointMult(i) = 1

Else
    UpDown2(i).Increment = 5
    SetPointMult(i) = 1

End If
   
   '  29/4/00 Had to do a similar thing in Warsash...
   ' UpDown2(i).max = DefaultSetPoint(i) * SetPointMult(i) * 2
   ' UpDown2(i).Min = DefaultSetPoint(i) * SetPointMult(i) / 2
   
   ' === 26/09/02 ===
   '  UpDown2(i).max = DefaultSetPoint(i) * SetPointMult(i) * 5
   '  UpDown2(i).Min = DefaultSetPoint(i) * SetPointMult(i) / 5
      UpDown2(i).max = DefaultSetPoint(i) * SetPointMult(i) * 50
      UpDown2(i).Min = DefaultSetPoint(i) * SetPointMult(i) / 50
    
  ' 7/1/00  DO NOT CRASH THE PROGRAM...
    setpoint_val = SetPoint(i) * SetPointMult(i)
    If ((setpoint_val <= UpDown2(i).max) And (setpoint_val >= UpDown2(i).Min)) Then
    
      UpDown2(i).Value = SetPoint(i) * SetPointMult(i)
      
    End If

Next

Label3.Caption = Mytext
'Call JFSetPointRefresh
End Sub

Private Sub FORCE_default_Click(Index As Integer)

'15/10/99
Dim New_set_point As String
Dim set_point As Single
Dim SetPoint As Single

Dim int1 As Integer
Dim this_alarm_text As String
Dim UnitStr As String
Dim AlarmSetPoint As Single

' === 2/12/00 ===
   
   Control_PC_check
   If This_Control_status = False Then Exit Sub


' 7/4/98
'''int1 = DoEvents

If Index = 0 Then

' change the HIGH HIGH Set-point

           ' set_point = UpDown2(0).Value / SetPointMult(0)
           ' set_point = Val(Label4(0)) / SetPointMult(0)
             set_point = Val(Label4(0))
            
            UnitStr = Mid$(alarmtxt(nAlarm_HIHI), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
            
              ' 7/1/00
               '    Call setHILO(nAlarm_HIHI, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                    Call setHILO(nAlarm_HIHI, set_point)       ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_HIHI, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_HIHI, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_HIHI, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_HIHI, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_HIHI), 10, Len(alarmtxt(nAlarm_HIHI)))
            Call Logging("Student Changed HI-HI Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)

ElseIf Index = 1 Then

' change the HIGH Set-point

           '  set_point = UpDown2(1).Value / SetPointMult(1)
           '  set_point = Val(Label4(1)) / SetPointMult(1)
              set_point = Val(Label4(1))
           
            UnitStr = Mid$(alarmtxt(nAlarm_HI), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
                  ' 7/1/00
                  'Call setHILO(nAlarm_HI, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                   Call setHILO(nAlarm_HI, set_point)       ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_HI, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_HI, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_HI, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_HI, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_HI), 10, Len(alarmtxt(nAlarm_HI)))
            Call Logging("Student Changed HI Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)
   
ElseIf Index = 2 Then


            'set_point = UpDown2(2).Value / SetPointMult(2)
             set_point = Val(Label4(2)) / SetPointMult(2)
             set_point = Val(Label4(2))

            UnitStr = Mid$(alarmtxt(nAlarm_LO), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
                   ' 7/1/00
                 ' Call setHILO(nAlarm_LO, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                   Call setHILO(nAlarm_LO, set_point)        ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_LO, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_LO, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_LO, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_LO, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_LO), 10, Len(alarmtxt(nAlarm_LO)))
            Call Logging("Student Changed LO Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)
   
ElseIf Index = 3 Then

' change the LOW LOW Set-point


            ' set_point = UpDown2(3).Value / SetPointMult(3)
            ' set_point = Val(Label4(3)) / SetPointMult(3)
              set_point = Val(Label4(3))

            UnitStr = Mid$(alarmtxt(nAlarm_LOLO), 6, 4)
            'AlarmSetPoint = ALARM_HILO(1, nAlarm_LO)
            If InStr(1, UnitStr, "%", 1) > 0 Then
            
                     ' 7/1/00
                  ' Call setHILO(nAlarm_LOLO, set_point * 0.01)       ' MUST BE A TANK So SAVE VALUE / 100
                    Call setHILO(nAlarm_LOLO, set_point)        ' MUST BE A TANK So SAVE VALUE / 100
            Else
                If SI_units Then 'convert units from SI to ANSI before saving
                    If InStr(1, UnitStr, "psi", 1) > 0 Then
                        UnitStr = "bar "
                        Call setHILO(nAlarm_LOLO, set_point * 14.5)   'Entered in Bar, save in psi
                    ElseIf InStr(1, UnitStr, "degF", 1) > 0 Then
                        UnitStr = "deg C"
                        Call setHILO(nAlarm_LOLO, ((set_point * 9) / 5) + 32) 'Entered in C , save in F
                    ElseIf InStr(1, UnitStr, "in", 1) > 0 Then
                        UnitStr = "mm"
                        Call setHILO(nAlarm_LOLO, (set_point / 25.4)) 'Entered in C , save in F
                    End If
                Else
                    Call setHILO(nAlarm_LOLO, set_point)              ' IS NOT A TANK
                End If
            End If
            this_alarm_text = Mid$(alarmtxt(nAlarm_LOLO), 10, Len(alarmtxt(nAlarm_LOLO)))
            Call Logging("Student Changed LO-LO Setpoint for Alarm No " & Format(nAlarm_LO, "###0") & " " & this_alarm_text & " to " & Format(set_point, "####0.00") & Space(1) & UnitStr, LogChangeAlarmSP)

ElseIf Index = 4 Then

  Unload Me
  
End If




End Sub

Private Sub Form_Load()
Dim i As Integer

Me.Height = 3780
'Me.Width = 7440

' 2/12/00
'Me.Width = 7950
Me.Width = 9255

Call InitializeAlarmSetPoints
LastUnitType = SI_units  'keep track of Unit selection in case there's a change



' === 20/01/03 PUT IN:===
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
    For i = 0 To 3
        Command1(i).Enabled = False
        FORCE_default(i).Enabled = False
    Next i
Else
    For i = 0 To 3
        Command1(i).Enabled = True
        FORCE_default(i).Enabled = True
    Next i
End If


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)



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

'Call refresh_values
If LastUnitType <> SI_units Then  're-initialize everything if there's a change in unit
    Call InitializeAlarmSetPoints
    LastUnitType = SI_units
End If
Debug.Print "updown2(2).max = " & UpDown2(2).max
Call JFSetPointRefresh

End Sub

