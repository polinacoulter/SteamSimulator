VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form PID_constants 
   Caption         =   "PID Constants"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   5010
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   2415
      Left            =   120
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2640
         Top             =   1320
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4140
         OleObjectBlob   =   "PID_constants.frx":0000
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Delta_sign_18.doc"
         TabIndex        =   15
         Top             =   1740
         Width           =   255
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         OleObjectBlob   =   "PID_constants.frx":5018
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Percent_sign_18.doc"
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Word.Document.8"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4140
         OleObjectBlob   =   "PID_constants.frx":A030
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam_Sim\Text\Integral_sign_18.doc"
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Title_label 
         Alignment       =   2  'Center
         Caption         =   "Plant Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Pband_label 
         Alignment       =   2  'Center
         Caption         =   "25 %"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Ti_label 
         Alignment       =   2  'Center
         Caption         =   "0.0 sec"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Pb  Proportional Band"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Ti  Integral Time"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   2
         X1              =   960
         X2              =   1080
         Y1              =   840
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         Index           =   2
         X1              =   960
         X2              =   1080
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         Index           =   2
         X1              =   240
         X2              =   360
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         Index           =   2
         X1              =   240
         X2              =   360
         Y1              =   720
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   1
         X1              =   960
         X2              =   1080
         Y1              =   1440
         Y2              =   1320
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         Index           =   1
         X1              =   960
         X2              =   1080
         Y1              =   1200
         Y2              =   1320
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         Index           =   1
         X1              =   240
         X2              =   360
         Y1              =   1320
         Y2              =   1440
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         Index           =   1
         X1              =   240
         X2              =   360
         Y1              =   1320
         Y2              =   1200
      End
      Begin VB.Label Td_label 
         Alignment       =   2  'Center
         Caption         =   "0.0 sec"
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   960
         X2              =   1080
         Y1              =   2040
         Y2              =   1920
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   960
         X2              =   1080
         Y1              =   1800
         Y2              =   1920
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   240
         X2              =   360
         Y1              =   1920
         Y2              =   2040
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   240
         X2              =   360
         Y1              =   1920
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Td  Derivative Time "
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label INC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label DEC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label INC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label DEC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label INC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label DEC_label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "PID_constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INC_Mousedown(3) As Boolean
Dim DEC_Mousedown(3) As Boolean



Private Sub DEC_label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DEC_Mousedown(Index) = True
    DEC_label(Index).BorderStyle = 0

End Sub

Private Sub DEC_label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 DEC_Mousedown(Index) = False
 DEC_label(Index).BorderStyle = 1

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, id_PID_constants)

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_PID_constants)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub INC_label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  INC_Mousedown(Index) = True
  INC_label(Index).BorderStyle = 0

End Sub

Private Sub INC_label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 INC_Mousedown(Index) = False
 INC_label(Index).BorderStyle = 1

End Sub

Private Sub Timer1_Timer()

    '  n_Ti = 172
    'n_Pband = 173
   ' Global Ti_value As Single
'Global Pband_value As Single


' === Integral Time ( Ti ) ===
If INC_Mousedown(1) Then
    
      If (Ti_value < 60#) Then Ti_value = Ti_value + 1#
                                  
      Call set_CONTROLLER_PID_constants(n_Ti, Ti_value, nLOOPNUMBER)
                                
ElseIf DEC_Mousedown(1) Then
    
      If (Ti_value > 0#) Then Ti_value = Ti_value - 1#
      
      Call set_CONTROLLER_PID_constants(n_Ti, Ti_value, nLOOPNUMBER)
          
Else
      Ti_value = LSS_reals(449)
      Ti_label.Caption = Format(Ti_value, "##0.0") & " sec"
End If
  
' === Proportional Band ( Pband ) ===
If INC_Mousedown(2) Then
    
      If (Pband_value < 60#) Then Pband_value = Pband_value + 1#
                                  
      Call set_CONTROLLER_PID_constants(n_Pband, Pband_value * 0.01, nLOOPNUMBER)
                                
ElseIf DEC_Mousedown(2) Then
    
      If (Pband_value > 0#) Then Pband_value = Pband_value - 1#
      
      Call set_CONTROLLER_PID_constants(n_Pband, Pband_value * 0.01, nLOOPNUMBER)
          
Else
      Pband_value = LSS_reals(450) * 100#
      Pband_label.Caption = Format(Pband_value, "##0.0") & " %"
End If


Select Case nLOOPNUMBER

  Case 1
    Title_label.Caption = "Drum Level STBD"
  Case 2
    Title_label.Caption = "Drum Level PORT"
    
  Case 3
    Title_label.Caption = "Steam Temperature STBD"
  Case 4
    Title_label.Caption = "Steam Temperature PORT"
    
  Case 5
    Title_label.Caption = "Fuel Oil STBD"
  Case 6
    Title_label.Caption = "Fuel Oil PORT"
    
  Case 7
    Title_label.Caption = "Combustion Air STBD"
  Case 8
    Title_label.Caption = "Combustion Air PORT"
  
  Case 9
    Title_label.Caption = "Boiler Master STBD"
  Case 10
    Title_label.Caption = "Boiler Master PORT"
  
  
  Case 11
    Title_label.Caption = "Plant Master"
    
    
  Case 12
    Title_label.Caption = "Feed Pump STBD"
  Case 13
    Title_label.Caption = "Feed Pump PORT"


End Select

 
  

End Sub
