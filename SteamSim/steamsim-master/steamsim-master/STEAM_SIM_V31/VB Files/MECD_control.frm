VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{D6ECA430-4BD6-11D3-AA45-00C04FA979C0}#20.0#0"; "SquareGauge.ocx"
Object = "{C17D71FA-613B-11D3-AA68-00C04FA979C0}#6.0#0"; "EOT.ocx"
Begin VB.Form MECD_control 
   Caption         =   "MECD"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   12150
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   10695
      Left            =   180
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   18865
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4380
         Top             =   8880
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Evaporator"
         Height          =   435
         Left            =   9300
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "LO Pump"
         Height          =   435
         Left            =   1140
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bunkers"
         Height          =   435
         Left            =   2640
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7485
         Width           =   1335
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "F.O. Services"
         Height          =   435
         Left            =   1140
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Main Turbine"
         Height          =   435
         Left            =   1140
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "SSDG"
         Height          =   435
         Left            =   1140
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Port Burner Management"
         Height          =   435
         Left            =   2640
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Main Circ Pump"
         Height          =   435
         Left            =   2640
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6750
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Main Swbd"
         Height          =   435
         Left            =   9240
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Port Boiler"
         Height          =   435
         Left            =   1140
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Control Air"
         Height          =   435
         Left            =   7740
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bilge Pump"
         Height          =   435
         Left            =   2640
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fuel to Burners"
         Height          =   435
         Left            =   9240
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Boiler Fronts"
         Height          =   435
         Left            =   7740
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Main Condenser"
         Height          =   435
         Left            =   9240
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "SSTG"
         Height          =   435
         Left            =   2640
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Stbd.Boiler"
         Height          =   435
         Left            =   9240
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "SW Pump"
         Height          =   435
         Left            =   7740
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Emergency swbd"
         Height          =   435
         Left            =   7740
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Stbd. Burner Management"
         Default         =   -1  'True
         Height          =   435
         Left            =   7740
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5280
         Width           =   1335
      End
      Begin SquareGauge.Gauge Gauge4 
         Height          =   1755
         Left            =   660
         TabIndex        =   0
         Top             =   3000
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   3096
         CentreCircleColour=   255
         ScaleText       =   "psi"
      End
      Begin SquareGauge.Gauge Gauge5 
         Height          =   1755
         Left            =   2700
         TabIndex        =   1
         Top             =   3000
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   3096
         CentreCircleColour=   255
         ScaleText       =   "psi"
      End
      Begin SquareGauge.Gauge Gauge6 
         Height          =   1755
         Left            =   7380
         TabIndex        =   2
         Top             =   3060
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   3096
         CentreCircleColour=   255
         ScaleText       =   "psi"
      End
      Begin SquareGauge.Gauge Gauge2 
         Height          =   2295
         Left            =   7080
         TabIndex        =   4
         Top             =   180
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4048
         CentreCircleColour=   255
         ScaleText       =   "psi"
      End
      Begin SquareGauge.Gauge Gauge3 
         Height          =   2295
         Left            =   2160
         TabIndex        =   5
         Top             =   150
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4048
         CentreCircleColour=   255
         ScaleText       =   "psi"
      End
      Begin SquareGauge.Gauge Gauge7 
         Height          =   1755
         Left            =   9540
         TabIndex        =   6
         Top             =   420
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   3096
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge8 
         Height          =   1755
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   3096
         CentreCircleColour=   255
         ScaleText       =   ""
      End
      Begin SquareGauge.Gauge Gauge10 
         Height          =   1755
         Left            =   9420
         TabIndex        =   17
         Top             =   3060
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   3096
         CentreCircleColour=   255
         ScaleText       =   "psi"
      End
      Begin SquareGauge.Gauge Gauge1 
         Height          =   2295
         Left            =   4680
         TabIndex        =   3
         Top             =   180
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4048
         CentreCircleColour=   255
         ScaleText       =   "in Hg"
      End
      Begin SquareGauge.Gauge Gauge9 
         Height          =   2295
         Left            =   4740
         TabIndex        =   36
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4048
         CentreCircleColour=   255
         ScaleText       =   "in Hg"
      End
      Begin EOT.EOTctl EOTctl1 
         Height          =   4335
         Left            =   4140
         TabIndex        =   37
         Top             =   5160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7646
         BackColor       =   65535
         TelegraphValue  =   130
      End
      Begin VB.Label Label8 
         Caption         =   "Label1"
         Height          =   255
         Left            =   6660
         TabIndex        =   38
         Top             =   10200
         Visible         =   0   'False
         Width           =   615
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
         Left            =   5340
         TabIndex        =   43
         Top             =   9720
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
         Left            =   6360
         TabIndex        =   41
         Top             =   9720
         Width           =   915
         WordWrap        =   -1  'True
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
         Left            =   4380
         TabIndex        =   39
         Top             =   9720
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   10920
         Top             =   4860
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   8760
         Top             =   4920
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4620
         Top             =   4860
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6240
         Top             =   2520
         Width           =   255
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2040
         Top             =   4920
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   8940
         Top             =   2580
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4140
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "RPM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4200
         TabIndex        =   18
         Top             =   5820
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "S.W. pressure"
         Height          =   315
         Left            =   7560
         TabIndex        =   13
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Main Circ Pump Pressure"
         Height          =   315
         Left            =   2640
         TabIndex        =   12
         Top             =   4860
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "L.O. pressure"
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Vacuum"
         Height          =   315
         Left            =   5280
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Port Boiler Pressure"
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         Top             =   2580
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Stbd. Boiler Pressure"
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   2580
         Width           =   1635
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
         Left            =   6360
         TabIndex        =   42
         Top             =   9720
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
         Left            =   5340
         TabIndex        =   44
         Top             =   9720
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
         Left            =   4380
         TabIndex        =   40
         Top             =   9720
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "MECD_control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim form_resize_flag As Boolean
Dim OLD_A_input_0 As Integer
Dim OLD_A_input_1 As Integer

Private Sub Bridge_label_Click(index As Integer)

Call setPCDI(306)

End Sub

Private Sub Command1_Click()

'BOILER_PORT.Show
Port_Boiler.Show

End Sub

Private Sub Command19_Click()

Bunkers.Show

End Sub

Private Sub EOTctl1_OnChange()

 ' 1/10/99
    form_resize_flag = True

End Sub

Private Sub ER_label_Click(index As Integer)

Call setPCDI(308)

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_EOT_MECD)

If D_OUTPUT(1061) Then
    EOTctl1.BackColor = &HFFFF&        ' YELLOW
Else
    EOTctl1.BackColor = &HFFFFFF       ' WHITE
End If

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_EOT_MECD)
    HavenStretch1.Visible = True
End If


End Sub

Private Sub Timer1_Timer()

' 1/10/99
' EOTctl1.InstructorTelegarphValue = A_INPUT(1)
' EOTctl1.InstructorThrottleVAlue = A_INPUT(0)
  
  
' 1/10/99   Follow the MECD
 EOTctl1.TelegraphValue = A_INPUT(1)
 EOTctl1.ThrottleVAlue = A_INPUT(0)
  
 EOTctl1.InstructorTelegarphValue = A_OUTPUT(131)
 EOTctl1.InstructorThrottleVAlue = A_OUTPUT(130)


If form_resize_flag Or (A_INPUT(0) <> OLD_A_input_0) Or (A_INPUT(1) <> OLD_A_input_1) Then
  Form_Resize
  form_resize_flag = False
  Label1.Caption = "RE-SIZED"
Else
  Label1.Caption = ""
End If

OLD_A_input_0 = A_INPUT(0)
OLD_A_input_1 = A_INPUT(1)
  
'   2/10/99
If D_OUTPUT(1057) Then
  Bridge_label(0).Visible = False
  Bridge_label(1).Visible = True
Else
  Bridge_label(1).Visible = False
  Bridge_label(0).Visible = True
End If

If D_OUTPUT(1061) Then
  ER_label(0).Visible = False
  ER_label(1).Visible = True
Else
  ER_label(1).Visible = False
  ER_label(0).Visible = True
End If

If OLD_DO_ER <> D_OUTPUT(1061) Then
  If D_OUTPUT(1061) Then
    EOTctl1.BackColor = &HFFFF&        ' YELLOW
  Else
    EOTctl1.BackColor = &HFFFFFF       ' WHITE
  End If
End If
  OLD_DO_ER = D_OUTPUT(1061)

If D_OUTPUT(1058) Then
  Transfer_requested_label(0).Visible = Not Transfer_requested_label(0).Visible
  Transfer_requested_label(1).Visible = Not Transfer_requested_label(0).Visible
ElseIf D_OUTPUT(1059) Then
  Transfer_requested_label(0).Visible = False
  Transfer_requested_label(1).Visible = True
Else
  Transfer_requested_label(0).Visible = True
End If


End Sub

Private Sub Transfer_requested_label_Click(index As Integer)

Call setPCDI(307)

End Sub
