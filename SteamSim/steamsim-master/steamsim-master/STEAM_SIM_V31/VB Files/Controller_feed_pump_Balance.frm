VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{17FE8C5E-5183-11D3-AA4D-00C04FA979C0}#7.0#0"; "HorizontalGauge.ocx"
Begin VB.Form Controller_feed_pump_Balance 
   Caption         =   "Feed Pump Set Point"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   4695
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   2475
      Left            =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4366
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3120
         Top             =   1380
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Left            =   300
         TabIndex        =   1
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   195
         Left            =   3360
         TabIndex        =   0
         Top             =   1920
         Width           =   255
      End
      Begin HorizontalGauge.HGauge H_gauge 
         Height          =   795
         Left            =   780
         TabIndex        =   2
         ToolTipText     =   "Balance Gauge"
         Top             =   300
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1402
         MinorTicks      =   1
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "psi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3180
         TabIndex        =   7
         Top             =   540
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Feed Pump Set Point"
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   2220
         X2              =   2340
         Y1              =   1680
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   2220
         X2              =   2340
         Y1              =   1440
         Y2              =   1560
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   1500
         X2              =   1620
         Y1              =   1560
         Y2              =   1680
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   1500
         X2              =   1620
         Y1              =   1560
         Y2              =   1440
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   2295
         Left            =   60
         Top             =   120
         Width           =   3795
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Balance"
         Height          =   255
         Left            =   -180
         TabIndex        =   3
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label INC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1980
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label DEC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1380
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
   End
End
Attribute VB_Name = "Controller_feed_pump_Balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INC_MAN_Mousedown As Boolean
Dim DEC_MAN_Mousedown As Boolean


Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()

Controller_FEED_PUMP_STBD.Show

End Sub

Private Sub Command2_Click()

Controller_FEED_PUMP_PORT.Show

End Sub

Private Sub DEC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DEC_MAN_Mousedown = True
    DEC_label.BorderStyle = 0

End Sub

Private Sub DEC_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 DEC_MAN_Mousedown = False
 DEC_label.BorderStyle = 1

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_feed_pump_balance)

'   Horizontal BALANCE Gauge
    
' === 12/12/01 14.129 ===
'    H_gauge.GaugeText 1, "Stbd"
'    H_gauge.GaugeText 2, ""
'    H_gauge.GaugeText 3, "Both"
'    H_gauge.GaugeText 4, ""
'    H_gauge.GaugeText 5, "Port"
    
    H_gauge.GaugeText 1, "0"
    H_gauge.GaugeText 2, ""
    H_gauge.GaugeText 3, "100"
    H_gauge.GaugeText 4, ""
    H_gauge.GaugeText 5, "200"
    
    H_gauge.Value = 127  '  MIDDLE  ( Scaled 0 to 255 )

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_feed_pump_balance)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub INC_label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  ' BALANCE
  INC_MAN_Mousedown = True
  INC_label.BorderStyle = 0

End Sub

Private Sub INC_label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 INC_MAN_Mousedown = False
 INC_label.BorderStyle = 1

End Sub

Private Sub Timer1_Timer()

'=== 12/12/01 14.129 ===
'Re-define Balance_feed_pump as the Set-Point for BOTH feed pumps (nom 100 psi)


' BALANCE CONTROL
     
If INC_MAN_Mousedown Then
    
      If (H_gauge.Value < 255) Then H_gauge.Value = H_gauge.Value + 1
     ' H_gauge.Value = H_gauge.Value + 1
                
      Call setCONTROLLERvars(210, H_gauge.Value)     ' Balance_feed_pump
                  
ElseIf DEC_MAN_Mousedown Then
    
      If (H_gauge.Value > 0) Then H_gauge.Value = H_gauge.Value - 1
     'H_gauge.Value = H_gauge.Value - 1
          
      Call setCONTROLLERvars(210, H_gauge.Value)       ' Balance_feed_pump
Else
      H_gauge.Value = Format(LSS_reals(485), "##0")     ' BOTTOM Gauge
End If


End Sub
