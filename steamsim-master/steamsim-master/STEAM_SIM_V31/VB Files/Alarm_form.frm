VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form Alarm_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Active Alarms"
   ClientHeight    =   11790
   ClientLeft      =   3675
   ClientTop       =   5430
   ClientWidth     =   9360
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11790
   ScaleWidth      =   9360
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   11475
      Left            =   0
      Top             =   60
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   20241
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2520
         Top             =   60
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
         Height          =   10410
         Left            =   0
         TabIndex        =   6
         Top             =   960
         Width           =   8355
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C000&
         Caption         =   "Alarm Reset"
         Height          =   795
         Left            =   240
         Picture         =   "Alarm_form.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Click to clear any non-active alarms from list"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Alarm_silence 
         BackColor       =   &H008080FF&
         Caption         =   "Alarm Silence MECD"
         Height          =   855
         Left            =   1320
         Picture         =   "Alarm_form.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click to clear any non-active alarms from list"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000C000&
         Caption         =   "Alarm Accept MECD"
         Height          =   855
         Left            =   180
         Picture         =   "Alarm_form.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click to clear any non-active alarms from list"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000C000&
         Caption         =   "Alarm Accept Elec"
         Height          =   855
         Left            =   6000
         Picture         =   "Alarm_form.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Click to clear any non-active alarms from list"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Alarm Silence Elec"
         Height          =   855
         Left            =   7140
         Picture         =   "Alarm_form.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click to clear any non-active alarms from list"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5220
         TabIndex        =   0
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox dummybox 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Active Alarms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3180
         TabIndex        =   10
         Top             =   60
         Width           =   2235
      End
      Begin VB.Label Totel_lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4260
         TabIndex        =   9
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Number: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   540
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Alarm_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer

Sub catch_alarm()

' 30/10/97 moved to sub catch_alarms (MainInterfacefrm)

End Sub

Private Sub Alarm_silence_Click()


' 7/9/00
Call setPCDI(326)



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

    ' For LSS Simulators....
    '  nDI = 15  ' ( To reset LSS SLOWDOWN/SHUTDOWN trips )
      
    ' For STEAM SIM ....
    '  nDI = ***
   
End If

' 10/12/98
'set_DI_flag = True
Call setPCDI(nDI)
set_dummy_focus = True


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

' === 11/12/00 ===
' Accept MECD Panel # 2 Alarms
Call setPCDI(93)


' === 6/12/01 ===
' We only need to use ONE PCDI for ALL of them...
'' Accept MECD Panel # 6 Alarms
'Call setPCDI(325)
'
'' Accept Electrical Alarms
'Call setPCDI(708)



End Sub

Private Sub Command4_Click()

' === 6/12/01  14.088 ===
' Accept Elec Alarms
Call setPCDI(708)


End Sub

Private Sub Command5_Click()

' === 6/12/01  14.087 ===
' Silence Elec Alarms
Call setPCDI(714)


End Sub

Private Sub Form_Load()

' 29/12/97
    Me.Height = alarm_height
     Me.Width = alarm_width

' 3/4/98  RWJ added:
old_alarm_latest(1) = -123   ' To ensure an up-date of the alarms after form load
    
Call get_alarm_text
set_dummy_focus = True

' 7/9/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    Alarm_silence.Visible = False
    
Else


      ' === 7/01/02 ===
      ' Alarm_silence.Visible = True
     
     If PTT Then
     
     ' === 13/11/02 ===
     '   Alarm_silence.Visible = False
     '   Command3.Visible = False
         Alarm_silence.Visible = True
         Command3.Visible = True
         Command3.Caption = "Alarm Accept"
         Alarm_silence.Caption = "Alarm Silence"
        
         Command4.Visible = False
         Command5.Visible = False
     Else
        Alarm_silence.Visible = True
        Command3.Visible = True
        Command4.Visible = True
        Command5.Visible = True
     End If
    
End If

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

    
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
Dim MyStr As String
Dim UnitStr As String * 4
Dim SetPoint As Single


Dim HIGH_setpoint As Integer
Dim alarm_no As Integer

'08/09/00
Dim NumberOfAlarms As Integer

AlarmChanges = False



'12/12/00 Debug.Print " In alarm timer 405 text = " & alarmtxt(405)








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
    NumberOfAlarms = 0
    List1.clear
    List1.AddItem "  TIME    ID     Description                              Value"
     For i = 1 To 450

        alarm_no = alarm_latest(i)
  
        If alarm_no > 0 And alarm_no < 451 Then
  
            If alarm(alarm_no) Then
                NumberOfAlarms = NumberOfAlarms + 1
                If ALARM_HILO(1, alarm_no) < 999.999 Then
                    HIGH_setpoint = True
                Else
                    HIGH_setpoint = False
                End If
                
                MyStr = alarmtxt(alarm_no)
                
                ' 10/5/00 Taken from Alarm_set_point which works OK...
                'UnitStr = Mid$(alarmtxt(nAlarm_HIHI), 6, 4)
                
                '  10/5/00 SO...
                'UnitStr = Mid$(mystr, 6, 9)
                 UnitStr = Mid$(MyStr, 6, 4)
                
                
                If alarm_no = 405 Then
                 alarm_no = 405
                End If



                If HIGH_setpoint Then
                    If InStr(1, UnitStr, "%", 1) > 0 Then
                    
                     '  11/5/00  CMA Tank Alarms are stored in % ( NOT 0 to 1.0 )
                     '  Alarm_set_point = ALARM_HILO(1, alarm_no) * 100#      ' MUST BE A TANK
                        Alarm_set_point = ALARM_HILO(1, alarm_no)             ' MUST BE A TANK
                        
                    Else
                        Alarm_set_point = ALARM_HILO(1, alarm_no)             '  IS NOT A TANK
                    End If
                    one_char = ">"
                Else
                    If InStr(1, UnitStr, "%", 1) > 0 Then
                    
                      ' 11/5/00  CMA Tank Alarms are stored in % ( NOT 0 to 1.0 )
                      ' Alarm_set_point = ALARM_HILO(2, alarm_no) * 100#      ' MUST BE A TANK
                        Alarm_set_point = ALARM_HILO(2, alarm_no)             ' MUST BE A TANK
                       
                    Else
                        Alarm_set_point = ALARM_HILO(2, alarm_no)             '  IS NOT A TANK
                    End If
                    one_char = "<"
                End If
                
                MyStr = alarmtxt(alarm_no)
               
                If alarm_no = 378 Then
                 alarm_no = 378
                End If
                
                
                'UnitStr = Mid$(mystr, 6, 9)
                If UnitStr <> "    " Then
                    If SI_units Then 'convert units to SI
                        'If UnitStr = "psi " Then
                        If InStr(1, MyStr, "psi", 1) > 0 Then
                            UnitStr = "bar "
                            SetPoint = Alarm_set_point / 14.5
                        ElseIf InStr(1, MyStr, "degF", 1) > 0 Then
                            UnitStr = "deg C"
                            SetPoint = (Alarm_set_point - 32) * 5# / 9#
                        End If
                    Else
                        SetPoint = Alarm_set_point
                    End If
                    this_alarm_text = alarm_time(i) & "  " & Format$(alarm_no, "000") _
                    & "   " & Mid$(MyStr, 10, Len(MyStr)) & "  " & one_char & Format(SetPoint, "###0.0") & " " & UnitStr
                    List1.AddItem this_alarm_text
                Else
                    'List1.AddItem alarm_time(i) & "  " & Format$(alarm_no, "000") _
                    '& "   " & this_alarm_text
                    this_alarm_text = alarm_time(i) & "  " & Format$(alarm_no, "000") _
                    & "   " & Mid$(MyStr, 10, Len(MyStr))
                    List1.AddItem this_alarm_text
 
                    
                End If
                
            End If
            Debug.Print "#" & i & "UnitStr = " & UnitStr
        End If
    Next i
    ' === 13/11/02 ===
    'Text3.text = str(NumberOfAlarms)
     Totel_lbl.Caption = str(NumberOfAlarms)
    
End If



End Sub
