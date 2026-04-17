VERSION 5.00
Begin VB.Form AlarmV2_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Active Alarms"
   ClientHeight    =   7440
   ClientLeft      =   3675
   ClientTop       =   5430
   ClientWidth     =   8385
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   8385
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2340
      TabIndex        =   3
      Text            =   "Active Alarms"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Alarm Reset"
      Height          =   735
      Left            =   120
      Picture         =   "AlarmV2_form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to clear any non-active alarms from list"
      Top             =   60
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6360
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   8115
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6540
      Top             =   120
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "AlarmV2_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer

Sub catch_alarm()

' 30/10/97 moved to sub catch_alarms (MainInterfacefrm)

End Sub

Private Sub Command1_Click()

If medium_speed Then

' 17/6/98
   If Portugal_flag Then
      nDI = 439 ' PORTUGAL    ( To reset SLOWDOWN/SHUTDOWN trips )
   Else
      nDI = 334 ' VIGO/FERROL ( To reset SLOWDOWN/SHUTDOWN trips )
   End If
   
Else
   nDI = 15  ' ( To reset LSS SLOWDOWN/SHUTDOWN trips )
End If

' 10/12/98
'set_DI_flag = True
Call setPCDI(nDI)
set_dummy_focus = True


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

' 29/12/97
    Me.Height = alarm_height
     Me.Width = alarm_width

' 3/4/98  RWJ added:
old_alarm_latest(1) = -123   ' To ensure an up-date of the alarms after form load
    
Call get_alarm_text
set_dummy_focus = True

    
End Sub

Private Sub Form_Resize()

'29/12/97
    If Me.WindowState = 0 Then ' If normal
        If Me.Width < 4000 Then Me.Width = 4000
        If Me.Height < 1000 Then Me.Height = 1000
    End If

    If Me.WindowState <> 1 Then ' If not minimised
        List1.Width = Me.ScaleWidth - List1.Left * 2
        List1.Height = Me.ScaleHeight - List1.Top
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    alarm_height = Me.Height
    alarm_width = Me.Width
'    Debug.Print Me.Height & "  " & Me.Width & "  " & Me.ScaleHeight & "  " & Me.ScaleWidth

End Sub

Private Sub RESET_Click()

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim LatestTime As Date
Dim AllowedTime As Date
Dim LatestId As Integer
Dim RepeatedTime As Boolean
Dim Alarm_set_point As Single
Dim one_char As String * 1
Dim AlarmLoggedOK As Byte

' 28/5/98
Dim this_alarm_text As String
Dim mystr As String
Dim UnitStr As String * 4
Dim SetPoint As Single


Dim HIGH_setpoint As Integer
Dim alarm_no As Integer
AlarmChanges = False



Debug.Print " In alarm timer 405 text = " & alarmtxt(405)

'compare past and present alarm arrays to see if there is a change. If there is, then clear list and refresh
For i = 1 To 451
    If old_alarm_latest(i) <> alarm_latest(i) Then
        AlarmChanges = True
        old_alarm_latest(i) = alarm_latest(i)
    End If
Next

'dummybox.SetFocus
'dummybox.SetFocus
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

If AlarmChanges Then

    List1.Clear
    List1.AddItem "  TIME    ID     Description                              Value"
     For i = 1 To 450

        alarm_no = alarm_latest(i)
  
        If alarm_no > 0 And alarm_no < 451 Then
  
            If alarm(alarm_no) Then
            
                If ALARM_HILO(1, alarm_no) < 999.999 Then
                    HIGH_setpoint = True
                Else
                    HIGH_setpoint = False
                End If
                mystr = alarmtxt(alarm_no)
                UnitStr = Mid$(mystr, 6, 9)
                If alarm_no = 405 Then
                 alarm_no = 405
                End If



                If HIGH_setpoint Then
                    If InStr(1, UnitStr, "%", 1) > 0 Then
                        Alarm_set_point = ALARM_HILO(1, alarm_no) * 100#      ' MUST BE A TANK
                    Else
                        Alarm_set_point = ALARM_HILO(1, alarm_no)             '  IS NOT A TANK
                    End If
                    one_char = ">"
                Else
                    If InStr(1, UnitStr, "%", 1) > 0 Then
                        Alarm_set_point = ALARM_HILO(2, alarm_no) * 100#      ' MUST BE A TANK
                    Else
                        Alarm_set_point = ALARM_HILO(2, alarm_no)             '  IS NOT A TANK
                    End If
                    one_char = "<"
                End If
                
                mystr = alarmtxt(alarm_no)
               
                If alarm_no = 378 Then
                 alarm_no = 378
                End If
                
                
                'UnitStr = Mid$(mystr, 6, 9)
                If UnitStr <> "    " Then
                    If SI_units Then 'convert units to SI
                        'If UnitStr = "psi " Then
                        If InStr(1, mystr, "psi", 1) > 0 Then
                            UnitStr = "bar "
                            SetPoint = Alarm_set_point / 14.5
                        ElseIf InStr(1, mystr, "degF", 1) > 0 Then
                            UnitStr = "deg C"
                            SetPoint = (Alarm_set_point - 32) * 5# / 9#
                        End If
                    Else
                        SetPoint = Alarm_set_point
                    End If
                    this_alarm_text = alarm_time(i) & "  " & Format$(alarm_no, "000") _
                    & "   " & Mid$(mystr, 10, Len(mystr)) & "  " & one_char & Format(SetPoint, "###0.0") & " " & UnitStr
                    List1.AddItem this_alarm_text
                Else
                    'List1.AddItem alarm_time(i) & "  " & Format$(alarm_no, "000") _
                    '& "   " & this_alarm_text
                    this_alarm_text = alarm_time(i) & "  " & Format$(alarm_no, "000") _
                    & "   " & Mid$(mystr, 10, Len(mystr))
                    List1.AddItem this_alarm_text
 
                    
                End If
                
            End If
            Debug.Print "#" & i & "UnitStr = " & UnitStr
        End If
    Next i
End If



End Sub
