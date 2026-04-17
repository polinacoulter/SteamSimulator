VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{C17D71FA-613B-11D3-AA68-00C04FA979C0}#6.0#0"; "EOT.ocx"
Begin VB.Form Frm_EOT_MECD 
   Caption         =   "MECD Telegraph"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   3885
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7335
      Left            =   120
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   12938
      Begin VB.Frame Frame2 
         Caption         =   "Steering Gear #1"
         Height          =   1635
         Left            =   120
         TabIndex        =   13
         Top             =   5400
         Width           =   1575
         Begin VB.CommandButton sbyPUMP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "S/B"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1140
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "STOP"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "START"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin PumpPrj.Pump HavenPump1 
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            CentrifugalPump =   0   'False
            OrientLeft      =   0   'False
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Steering Gear #2"
         Height          =   1635
         Left            =   1740
         TabIndex        =   8
         Top             =   5400
         Width           =   1575
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "START"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "STOP"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton sbyPUMP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "S/B"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1140
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin PumpPrj.Pump HavenPump1 
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            CentrifugalPump =   0   'False
            OrientLeft      =   0   'False
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   60
         Top             =   3720
      End
      Begin EOT.EOTctl EOTctl1 
         Height          =   4335
         Left            =   -120
         TabIndex        =   0
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7646
         ThrottleVAlue   =   170
         BackColor       =   65535
         TelegraphValue  =   120
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   5040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Bridge_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bridge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Bridge_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bridge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label ER_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Engine Room"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   4560
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label ER_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Engine Room"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   4560
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Transfer_requested_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Transfer Request"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   4560
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Transfer_requested_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transfer Request"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   4560
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "Frm_EOT_MECD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim form_resize_flag As Boolean
Dim OLD_A_input_0 As Integer
Dim OLD_A_input_1 As Integer
' === 12/11/02 ===
Dim set_dummy_focus As Integer
'
'



Private Sub Bridge_label_Click(Index As Integer)


Call setPCDI(306)


End Sub

Private Sub EOTctl1_OnChange()


' === 12/11/02 ===
'If (PTT) Then
'
'    ' IF PTT then we do NOT use the real Hardware pot
'
'    ' Send to Fortran  ( NB VB I/O program is NOT running )
'    Call setAI(1, EOTctl1.TelegraphValue)
'    Call setAI(0, EOTctl1.ThrottleVAlue)
'
'    ' Send to VB ( NB VB I/O program is NOT running )
'    A_INPUT(1) = EOTctl1.TelegraphValue
'    A_INPUT(0) = EOTctl1.ThrottleVAlue
'
'End If


If (PTT) Then
    
    ' === 03/12/02 REMOVED ===
    'If I_AM_HOST Then
    '        Call setAI(1, EOTctl1.TelegraphValue)
    '        Call setAI(0, EOTctl1.ThrottleVAlue)
    '        A_INPUT(1) = EOTctl1.TelegraphValue
    '        A_INPUT(0) = EOTctl1.ThrottleVAlue
    'Else
    '    If This_PC_number = 1 Then
    '        Call setAI(1, EOTctl1.TelegraphValue)
    '        Call setAI(0, EOTctl1.ThrottleVAlue)
    '        A_INPUT(1) = EOTctl1.TelegraphValue
    '        A_INPUT(0) = EOTctl1.ThrottleVAlue
    '    Else
    '        Me.Caption = "MECD Telegraph MONITORING"
    '    End If
    'End If
    
    ' === 03/12/02 ADDED ===
    If EOTctl1.TelegraphValue <> VB_More_integers(2) Then
            Call setAI(1, EOTctl1.TelegraphValue)
    End If
    If EOTctl1.ThrottleVAlue <> VB_More_integers(1) Then
            Call setAI(0, EOTctl1.ThrottleVAlue)
    End If
    
    
Else
        ' Use the HARDWARE REAL POT
End If

    
    
 ' 1/10/99
    form_resize_flag = True
    
    
End Sub

Private Sub ER_label_Click(Index As Integer)

Call setPCDI(308)

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_EOT_MECD)



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


  'set equal to real eot
' 1/10/99
'  EOTctl1.TelegraphValue = A_OUTPUT(131)
'  EOTctl1.ThrottleVAlue = A_OUTPUT(130)


' === 29/11/02 ===============
If PTT Then
  EOTctl1.TelegraphValue = VB_More_integers(2)
  EOTctl1.ThrottleVAlue = VB_More_integers(1)
End If
' === 29/11/02 end of edit ===


'c       === 7/01/02 For the EOT MECD_ER_DO ===
'             LSS_LOGICALS(96) = DO(1061)


'If D_OUTPUT(1061) Then
If LSS_LOGICALS(96) Then
    EOTctl1.BackColor = &HFFFF&        ' YELLOW
Else
    EOTctl1.BackColor = &HFFFFFF       ' WHITE
End If

' === 03/12/02 ===
VB_More_integers_OLD(1) = -123 ' To ensure an update
VB_More_integers_OLD(2) = -123 ' To ensure an update
  
  
End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_EOT_MECD)
    HavenStretch1.Visible = True
End If


End Sub

Private Sub OFFpump_Click(Index As Integer)

' === 12/11/02 1.05 ===
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 12))  ' Steering Gear # 1 STOP (Pump number 12)
Else
   Call setPCDI(PumpIO_LSS(4, 25))  ' Steering Gear # 2 STOP (Pump number 25)
End If



End Sub

Private Sub ONpump_Click(Index As Integer)


' === 12/11/02 1.05 ===

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 12))   ' Steering Gear # 1 START (Pump number 12)
Else
   Call setPCDI(PumpIO_LSS(2, 25))   ' Steering Gear # 2 START (Pump number 25)
End If


End Sub

Private Sub sbyPUMP_Click(Index As Integer)

' === 12/11/02 1.05 ===
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

   'Note:  PumpIO(5=Stand-bye,nPump)
If Index = 0 Then
    Call setPCDI(PumpIO_LSS(5, 12))  ' Steering Gear # 1 STBY (Pump number 12)
Else
    Call setPCDI(PumpIO_LSS(5, 25))  ' Steering Gear # 2 STBY (Pump number 25)
End If




End Sub

Private Sub Timer1_Timer()


' === 12/11/02 1.05 ===
Me.Caption = "MECD Telegraph"

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

Call update_pump_LSS(Me, 0, 12)     ' Steering Gear # 1
Call update_pump_LSS(Me, 1, 25)     ' Steering Gear # 2

' ================



' 1/10/99
' EOTctl1.InstructorTelegarphValue = A_INPUT(1)
' EOTctl1.InstructorThrottleVAlue = A_INPUT(0)
  
  
' === 03/12/02 ============================================

'If (PTT) Then
'' MOVE to change event
'' 12/11/99   Set the AI pot values ( If PTT )
''    Call setAI(1, EOTctl1.TelegraphValue)
''    Call setAI(0, EOTctl1.ThrottleVAlue)
'Else
'' 12/11/99   Follow the MECD pots ( If Full Mission Sim )
'  EOTctl1.TelegraphValue = A_INPUT(1)
'  EOTctl1.ThrottleVAlue = A_INPUT(0)
'End If

If (VB_More_integers_OLD(1) <> VB_More_integers(1)) Then
    VB_More_integers_OLD(1) = VB_More_integers(1)
      EOTctl1.ThrottleVAlue = VB_More_integers(1)
End If


If (VB_More_integers_OLD(2) <> VB_More_integers(2)) Then

    VB_More_integers_OLD(2) = VB_More_integers(2)

    ' === 06/12/02 TELEGRAPH "creeps" up-wards so =========================
    '    EOTctl1.TelegraphValue = VB_More_integers(2)
    If VB_More_integers(2) > EOT_instructor(1) Then
    
     ' === 06/12/02 ===
     '  EOTctl1.TelegraphValue = 250      ' Full Ahead
        EOTctl1.TelegraphValue = 240      ' Full Ahead
        
    ElseIf VB_More_integers(2) > EOT_instructor(2) Then
    
     ' === 06/12/02 ===
        EOTctl1.TelegraphValue = 220      ' Half Ahead
        
    ElseIf VB_More_integers(2) > EOT_instructor(3) Then
    
     ' === 06/12/02 ===
     '  EOTctl1.TelegraphValue = 200      ' Slow Ahead
        EOTctl1.TelegraphValue = 190      ' Slow Ahead
        
    ElseIf VB_More_integers(2) > EOT_instructor(4) Then
    
     ' === 06/12/02 ===
     ' EOTctl1.TelegraphValue = 180      ' Dead Slow Ahead
       EOTctl1.TelegraphValue = 170      ' Dead Slow Ahead
        
    ElseIf VB_More_integers(2) > EOT_instructor(5) Then
        EOTctl1.TelegraphValue = 150      ' Stand-by
    ElseIf VB_More_integers(2) > EOT_instructor(6) Then
        EOTctl1.TelegraphValue = 120      ' Stop
    ElseIf VB_More_integers(2) > EOT_instructor(7) Then
        EOTctl1.TelegraphValue = 100      ' FWE
    ElseIf VB_More_integers(2) > EOT_instructor(8) Then
        EOTctl1.TelegraphValue = 70      ' D Slow Astern
    ElseIf VB_More_integers(2) > EOT_instructor(9) Then
        EOTctl1.TelegraphValue = 50      ' Slow Astern
    ElseIf VB_More_integers(2) > EOT_instructor(10) Then
        EOTctl1.TelegraphValue = 30      ' Half Astern
    Else
        EOTctl1.TelegraphValue = 8     ' Full Astern
    End If
    ' === 06/12/02 end of edit ===============
    
End If


' === 03/12/02 end of edit =================================



'c       === 7/01/02 ===
'      More_integers(1) = AI(0)
'      More_integers(2) = AI(1)
'      More_integers(3) = AO(131)
'      More_integers(4) = AO(130)


'       === 7/01/02 ===
' EOTctl1.InstructorTelegarphValue = A_OUTPUT(131)
' EOTctl1.InstructorThrottleVAlue = A_OUTPUT(130)
  EOTctl1.InstructorTelegarphValue = VB_More_integers(3)
  EOTctl1.InstructorThrottleVAlue = VB_More_integers(4)
 
 


'       === 7/01/02 ===
'If form_resize_flag Or (A_INPUT(0) <> OLD_A_input_0) Or (A_INPUT(1) <> OLD_A_input_1) Then
If form_resize_flag Or (VB_More_integers(1) <> OLD_A_input_0) Or (VB_More_integers(2) <> OLD_A_input_1) Then
  
  '05/06/07 ===== next 3 lines commented out& next 2 added
  'Form_Resize
  'form_resize_flag = False
  'Label1.Caption = "RE-SIZED"
  EOTctl1.Visible = False
  EOTctl1.Visible = True
  
Else
  Label1.Caption = ""
End If

'       === 7/01/02 ===
'OLD_A_input_0 = A_INPUT(0)
'OLD_A_input_1 = A_INPUT(1)
OLD_A_input_0 = VB_More_integers(1)
OLD_A_input_1 = VB_More_integers(2)
  
' === 7/01/02 ===
'c       === 7/01/02 For the EOT MECD_bridge_DO ===
'             LSS_LOGICALS(93) = DO(1057)
'c       === 7/01/02 For the EOT MECD_transfer_DO ===
'             LSS_LOGICALS(94) = DO(1058)
'c       === 7/01/02 For the EOT MECD_transfer_DO + 1 ===
'             LSS_LOGICALS(95) = DO(1059)
'c       === 7/01/02 For the EOT MECD_ER_DO ===
'             LSS_LOGICALS(96) = DO(1061)




' === 7/01/02 ===
'If D_OUTPUT(1057) Then
If LSS_LOGICALS(93) Then
  Bridge_label(0).Visible = False
  Bridge_label(1).Visible = True
Else
  Bridge_label(1).Visible = False
  Bridge_label(0).Visible = True
End If

' === 7/01/02 ===
'If D_OUTPUT(1061) Then
If LSS_LOGICALS(96) Then
  ER_label(0).Visible = False
  ER_label(1).Visible = True
Else
  ER_label(1).Visible = False
  ER_label(0).Visible = True
End If


' === 7/01/02 ===
'If OLD_DO_ER <> D_OUTPUT(1061) Then

' === 7/01/02 ===
 'If D_OUTPUT(1061) Then
  If LSS_LOGICALS(96) Then
    EOTctl1.BackColor = &HFFFF&        ' YELLOW
  Else
    EOTctl1.BackColor = &HFFFFFF       ' WHITE
  End If
  
'End If

  OLD_DO_ER = D_OUTPUT(1061)


' === 7/01/02 ===
'If D_OUTPUT(1058) Then
  If LSS_LOGICALS(94) Then
  Transfer_requested_label(0).Visible = Not Transfer_requested_label(0).Visible
  Transfer_requested_label(1).Visible = Not Transfer_requested_label(0).Visible
'ElseIf D_OUTPUT(1059) Then
ElseIf LSS_LOGICALS(95) Then
  Transfer_requested_label(0).Visible = False
  Transfer_requested_label(1).Visible = True
Else
  Transfer_requested_label(0).Visible = True
End If


  
End Sub

Private Sub Transfer_requested_label_Click(Index As Integer)

Call setPCDI(307)

End Sub
