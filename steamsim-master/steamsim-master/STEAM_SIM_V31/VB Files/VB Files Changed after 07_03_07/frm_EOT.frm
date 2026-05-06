VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{C17D71FA-613B-11D3-AA68-00C04FA979C0}#6.0#0"; "EOT.ocx"
Begin VB.Form Frm_EOT 
   Caption         =   "Bridge Telegraph"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   3780
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7995
      Left            =   60
      Top             =   60
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   14102
      Begin VB.TextBox Shaft_rpm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1200
         TabIndex        =   21
         Text            =   "99 rpm"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Vship 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1200
         TabIndex        =   20
         Text            =   "6.3 Knotts"
         Top             =   60
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Steering Gear #2"
         Height          =   1635
         Left            =   1740
         TabIndex        =   14
         Top             =   6300
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
            Index           =   1
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Index           =   1
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   16
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
            Index           =   1
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin PumpPrj.Pump HavenPump1 
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   18
            Top             =   840
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            CentrifugalPump =   0   'False
            OrientLeft      =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Steering Gear #1"
         Height          =   1635
         Left            =   120
         TabIndex        =   8
         Top             =   6300
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
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Index           =   0
            Left            =   300
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   11
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
            Index           =   0
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
            Index           =   0
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
         Left            =   120
         Top             =   4380
      End
      Begin EOT.EOTctl EOTctl1 
         Height          =   4335
         Left            =   -60
         TabIndex        =   0
         Top             =   900
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7646
         InstructorTelegarphValue=   150
         BackColor       =   12632319
         TelegraphValue  =   130
      End
      Begin VB.TextBox dummybox 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Shaft rpm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Sea Speed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   5940
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
         Top             =   5460
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
         Top             =   5460
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
         Top             =   5460
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
         Top             =   5460
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
         Top             =   5460
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
         Top             =   5460
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "Frm_EOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim set_dummy_focus As Integer
'
' === 07/06/02 ===
'Private EOT_instructor(11) As Integer
'
'



Private Sub Bridge_label_Click(Index As Integer)

'  integer BRIDGE_transfer_DI
'  data  BRIDGE_transfer_DI /4401/
'  integer BRIDGE_bridge_DI
'  data  BRIDGE_bridge_DI /4402/
'  integer BRIDGE_ER_DI
'  data  BRIDGE_ER_DI /4403/

' 14/10/99
'Call setPCDI(3636)
Call setPCDI(4402)


End Sub

Private Sub EOTctl1_OnChange()

  ' Fortran writes to the A_OUTPUT array...
  '  A_OUTPUT(131) = EOTctl1.TelegraphValue
  '  A_OUTPUT(130) = EOTctl1.ThrottleVAlue
    
    ' 20/7/00
    If Not Fault(1003) Then
      Call set_AO_LSS(131, EOTctl1.TelegraphValue)
    End If
    
    Call set_AO_LSS(130, EOTctl1.ThrottleVAlue)
    
End Sub

Private Sub ER_label_Click(Index As Integer)

'  integer BRIDGE_transfer_DI
'  data  BRIDGE_transfer_DI /4401/
'  integer BRIDGE_bridge_DI
'  data  BRIDGE_bridge_DI /4402/
'  integer BRIDGE_ER_DI
'  data  BRIDGE_ER_DI /4403/

' 14/10/99
'Call setPCDI(3638)
Call setPCDI(4403)

End Sub

Private Sub Form_Load()

' === 12/11/01 ===
Dim FileNo As Integer
Dim IN_string As String
Dim i As Integer

' === 11/12/01 ===
'On Error Resume Next
On Error GoTo ErrH

Call JYAScreenInit(Me, id_EOT)



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


  'set equal to real eot
  EOTctl1.TelegraphValue = A_OUTPUT(131)
  EOTctl1.ThrottleVAlue = A_OUTPUT(130)
  

'c       === 7/01/02 For the EOT MECD_bridge_DO ===
'             LSS_LOGICALS(93) = DO(1057)


'If D_OUTPUT(1057) Then
If LSS_LOGICALS(93) Then
    EOTctl1.BackColor = &HC0C0FF       ' PINK
Else
    EOTctl1.BackColor = &HFFFFFF       ' WHITE
End If


' === 11/12/01 ===


' === 07/06/02 =================

' Now done in HOST.frm
'FileNo = FreeFile
'Open "c:\STEAM_SIM\text\MECD_EOT.txt" For Input As #FileNo
''
'For i = 1 To 10
'    Input #FileNo, IN_string
'    EOT_instructor(i) = Val(IN_string)
'Next i
''
'Close #FileNo

' ===============================

Exit Sub

ErrH:
resp = MsgBox("There was an error reading C:\Steam_SIm\MECD_EOT.txt" & Chr(10) & _
       "You must ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
       Stop


  
End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_EOT)
    HavenStretch1.Visible = True
End If


End Sub

Private Sub OFFpump_Click(Index As Integer)


' === 14/11/01 ===

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

' === 14/11/01 ===

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


' === 14/11/01 ===

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

' === 14/11/01 ===
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' === 14/12/02 19.161 ===
Vship.text = Format(LSS_reals(54), "##0.0") & " Knotts"
Shaft_rpm.text = Format(LSS_reals(80), "##0") & " rpm"


' === 14/11/01 ===
Call update_pump_LSS(Me, 0, 12)     ' Steering Gear # 1
Call update_pump_LSS(Me, 1, 25)     ' Steering Gear # 2

'c       === 7/01/02 ===
'      More_integers(1) = AI(0)
'      More_integers(2) = AI(1)
'      More_integers(3) = AO(131)
'      More_integers(4) = AO(130)


' === 11/12/01 Calibration Error so do the following: ===
'  EOTctl1.InstructorTelegarphValue = A_INPUT(1)

' === 7/01/02 ===
' A_INPUT(1) changed to VB_More_integers(2)....
If VB_More_integers(2) > EOT_instructor(1) Then
    EOTctl1.InstructorTelegarphValue = 250      ' Full Ahead
ElseIf VB_More_integers(2) > EOT_instructor(2) Then
    EOTctl1.InstructorTelegarphValue = 220      ' Half Ahead
ElseIf VB_More_integers(2) > EOT_instructor(3) Then
    EOTctl1.InstructorTelegarphValue = 200      ' Slow Ahead
ElseIf VB_More_integers(2) > EOT_instructor(4) Then
    EOTctl1.InstructorTelegarphValue = 180      ' Dead Slow Ahead
ElseIf VB_More_integers(2) > EOT_instructor(5) Then
    EOTctl1.InstructorTelegarphValue = 150      ' Stand-by
ElseIf VB_More_integers(2) > EOT_instructor(6) Then
    EOTctl1.InstructorTelegarphValue = 120      ' Stop
ElseIf VB_More_integers(2) > EOT_instructor(7) Then
    EOTctl1.InstructorTelegarphValue = 100      ' FWE
ElseIf VB_More_integers(2) > EOT_instructor(8) Then
    EOTctl1.InstructorTelegarphValue = 70      ' D Slow Astern
ElseIf VB_More_integers(2) > EOT_instructor(9) Then
    EOTctl1.InstructorTelegarphValue = 50      ' Slow Astern
ElseIf VB_More_integers(2) > EOT_instructor(10) Then
    EOTctl1.InstructorTelegarphValue = 30      ' Half Astern
Else
    EOTctl1.InstructorTelegarphValue = 8     ' Full Astern
End If

  
  
' === 7/01/02 ===
' A_INPUT(0) changed to VB_More_integers(1)....
  
'  EOTctl1.InstructorThrottleVAlue = A_INPUT(0)
   EOTctl1.InstructorThrottleVAlue = VB_More_integers(1)
  
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
'If OLD_DO_BRIDGE <> D_OUTPUT(1057) Then

'  Label1.Caption = D_OUTPUT(1057)

' If D_OUTPUT(1057) Then
  If LSS_LOGICALS(93) Then
    EOTctl1.BackColor = &HC0C0FF       ' PINK
  Else
    EOTctl1.BackColor = &HFFFFFF       ' WHITE
  End If
  
'Else
'  Label1.Caption = " "
'End If

  OLD_DO_BRIDGE = D_OUTPUT(1057)


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

'  integer BRIDGE_transfer_DI
'  data  BRIDGE_transfer_DI /4401/
'  integer BRIDGE_bridge_DI
'  data  BRIDGE_bridge_DI /4402/
'  integer BRIDGE_ER_DI
'  data  BRIDGE_ER_DI /4403/

' 14/10/99
'Call setPCDI(3637)
Call setPCDI(4401)

End Sub
