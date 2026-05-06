VERSION 5.00
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.0#0"; "HavenPump v1_42.ocx"
Begin VB.Form Elec_ctrl 
   Caption         =   "Electrical Control Panel"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   5625
   Begin VB.Frame Frame2 
      Caption         =   "Shore Supply"
      Height          =   1695
      Left            =   2760
      TabIndex        =   23
      Top             =   840
      Width           =   1215
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1140
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON/OFF"
         Height          =   255
         Index           =   5
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   660
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pre-Lub"
      Height          =   1635
      Left            =   60
      TabIndex        =   14
      Top             =   2640
      Width           =   3975
      Begin VB.CommandButton sbypump 
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
         Height          =   255
         Index           =   1
         Left            =   2880
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         Height          =   255
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton sbypump 
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
         Height          =   255
         Index           =   4
         Left            =   1680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Index           =   4
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         Height          =   255
         Index           =   4
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton sbypump 
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
         Height          =   255
         Index           =   3
         Left            =   480
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         Height          =   255
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "SSTG#2"
         Height          =   195
         Left            =   2760
         TabIndex        =   26
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "SSTG#1"
         Height          =   195
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "SSDG"
         Height          =   195
         Left            =   420
         TabIndex        =   18
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2280
      Top             =   240
   End
   Begin VB.Frame DA1_frame 
      Caption         =   "EDA"
      Height          =   1695
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   1215
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         Height          =   255
         Index           =   0
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Index           =   0
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   780
         Width           =   615
      End
      Begin VB.CommandButton sbypump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AUTO"
         Height          =   315
         Index           =   0
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame DA1_frame 
      Caption         =   "SSDG"
      Height          =   2535
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         Height          =   255
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         Height          =   255
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   780
         Width           =   615
      End
      Begin VB.CommandButton MAN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MAN"
         Height          =   255
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton AUTO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AUTO"
         Height          =   315
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1380
         Width           =   615
      End
      Begin VB.CommandButton SBY1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/B 1"
         Height          =   255
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1740
         Width           =   615
      End
      Begin VB.CommandButton SBY2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/B 2"
         Height          =   255
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2040
         Width           =   615
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Top             =   180
      Width           =   435
   End
   Begin PumpPrj.Pump HavenPump1 
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      LineCol         =   16711680
   End
   Begin PumpPrj.Pump HavenPump1 
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      LineCol         =   16711680
   End
End
Attribute VB_Name = "Elec_ctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim set_dummy_focus As Integer

Private Sub OFF_Click(Index As Integer)


End Sub

Private Sub ON_Click(Index As Integer)



End Sub

Private Sub AUTO_Click(Index As Integer)

set_dummy_focus = True
Control_PC_check
If This_Control_status = False Then Exit Sub

If Index = 1 Then
    Call setPCDI(EAUTOSW(1))    ' D/A 1 to AUTO
ElseIf Index = 2 Then
    Call setPCDI(EAUTOSW(2))    ' D/A 2 to AUTO
End If


End Sub

Private Sub Command1_Click()


End Sub

Private Sub form_Load()

   set_dummy_focus = True
   
    Me.Height = elec_ctrl_height
    Me.Width = elec_ctrl_width

End Sub

Private Sub Form_Resize()

'11/3/98
If Me.WindowState = 0 Then ' If normal
        If Me.Width < 4000 Then Me.Width = 4000
        If Me.Height < 1000 Then Me.Height = 1000
End If

'If Me.WindowState <> 1 Then ' If not minimised
'        List1.Width = Me.ScaleWidth - List1.Left * 2
'        List1.Height = Me.ScaleHeight - List1.Top
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)


' 11/8/00
'    alarm_height = Me.Height
'     alarm_width = Me.Width
    elec_ctrl_height = Me.Height
     elec_ctrl_width = Me.Width

End Sub

Private Sub MAN_Click(Index As Integer)

set_dummy_focus = True
Control_PC_check
If This_Control_status = False Then Exit Sub

If Index = 1 Then
    Call setPCDI(EMANSW(1))    ' D/A 1 to MANUAL
ElseIf Index = 2 Then
    Call setPCDI(EMANSW(2))    ' D/A 2 to MANUAL
End If

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True
Control_PC_check
If This_Control_status = False Then Exit Sub
   
If medium_speed Then
  If Index = 0 Then
   Call setPCDI(1107)    ' Emerg ALT OFF
  ElseIf Index = 1 Then
    Call setPCDI(ESTOPSW(1))    ' STOP D/A 1
  ElseIf Index = 2 Then
    Call setPCDI(ESTOPSW(2))    ' STOP D/A 2
  ElseIf Index = 3 Then
  '   Note:  PumpIO(1=ON,nPump)
   Call setPCDI(PumpIO(2, 41))     ' D/A 1 PRE-LUB OFF
  ElseIf Index = 4 Then
  '   Note:  PumpIO(1=ON,nPump)
   Call setPCDI(PumpIO(2, 42))     ' D/A 2 PRE-LUB OFF
  ElseIf Index = 5 Then
    ss_supply = False
  End If
Else
  If Index = 0 Then
   Call setPCDI(789)    ' Emerg ALT OFF
  End If
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 4/6/98
'If Index = 0 Then
'   Call setPCDI(1108)    ' Emerg ALT ON
'ElseIf Index = 1 Then
'    Call setPCDI(ESTARTSW(1))    ' START D/A 1
'ElseIf Index = 2 Then
'    Call setPCDI(ESTARTSW(2))    ' START D/A 2
'ElseIf Index = 3 Then
''   Note:  PumpIO(1=ON,nPump)
'   Call setPCDI(PumpIO(1, 41))   ' START D/A 1 Pre-lub pump
'ElseIf Index = 4 Then
''   Note:  PumpIO(1=ON,nPump)
'   Call setPCDI(PumpIO(1, 42))   ' START D/A 2 Pre-lub pump
'ElseIf Index = 5 Then
'    ss_supply = True
'End If

If medium_speed Then

   Select Case Index
      Case 0
         Call setPCDI(1108)           ' Emerg ALT ON
      Case 1
         Call setPCDI(ESTARTSW(1))    ' START D/A 1
      Case 2
         Call setPCDI(ESTARTSW(2))    ' START D/A 2
      Case 3
'        Note:  PumpIO(1=ON,nPump)
         Call setPCDI(PumpIO(1, 41))   ' START D/A 1 Pre-lub pump
      Case 4
'        Note:  PumpIO(1=ON,nPump)
         Call setPCDI(PumpIO(1, 42))   ' START D/A 2 Pre-lub pump
      Case 5
         ss_supply = True
    End Select
    
Else

   Select Case Index
      Case 0
         Call setPCDI(781)           ' Emerg ALT ON

   End Select

End If


End Sub

Private Sub SBY1_Click(Index As Integer)

set_dummy_focus = True
Control_PC_check
If This_Control_status = False Then Exit Sub

If Index = 1 Then
    Call setPCDI(ESTBY1SW(1))    ' D/A 1 to STAND-BY 1
ElseIf Index = 2 Then
    Call setPCDI(ESTBY1SW(2))    ' D/A 2 to STAND-BY 1
End If

End Sub

Private Sub SBY2_Click(Index As Integer)

set_dummy_focus = True
Control_PC_check
If This_Control_status = False Then Exit Sub

If Index = 1 Then
    Call setPCDI(ESTBY2SW(1))    ' D/A 1 to STAND-BY 2
ElseIf Index = 2 Then
    Call setPCDI(ESTBY2SW(2))    ' D/A 2 to STAND-BY 2
End If

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
If medium_speed Then
  If Index = 0 Then
     Call setPCDI(1100)    ' Emerg ALT AUTO
  End If
Else
  If Index = 0 Then
     Call setPCDI(785)    ' Emerg ALT AUTO
  End If
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' Emergency ALT ON and OFF
If elec_lamps(14) Then
   ONpump(0).BackColor = &HFF00&      ' GREEN
   OFFpump(0).BackColor = &HC0C0C0    ' GREY
'   ALT_image(0).Picture = GEN_on.Picture
Else
   ONpump(0).BackColor = &HC0C0C0     ' GREY
   OFFpump(0).BackColor = &HFFFFFF    ' WHITE
'   ALT_image(0).Picture = GEN_off.Picture
End If

' Emergency ALT AUTO
If elec_lamps(16) Then
' 26/3/98
'  sbyPUMP(0).BackColor = &HFF00FF     ' PURPLE
   sbyPUMP(0).BackColor = &HFF80FF     ' LIGHT PURPLE
Else
   sbyPUMP(0).BackColor = &HC0C0C0     ' GREY
End If

' SHORE SUPPLY ON and OFF
If elec_lamps(18) Then
   ONpump(5).BackColor = &HFF00&      ' GREEN
   OFFpump(5).BackColor = &HC0C0C0    ' GREY
Else
   ONpump(5).BackColor = &HC0C0C0     ' GREY
   OFFpump(5).BackColor = &HFFFFFF    ' WHITE
End If


' 21/9/99 Only ONE D/A  ( SSDG )
'For i = 1 To 2
For i = 1 To 1

' DIESEL ALTs ON and OFF
  If elec_lamps(i + 20) Then   ' STARTLP=true
     ONpump(i).BackColor = &HFF00&      ' GREEN
     OFFpump(i).BackColor = &HC0C0C0    ' GREY
  Else
     ONpump(i).BackColor = &HC0C0C0     ' GREY
     OFFpump(i).BackColor = &HFFFFFF    ' WHITE
  End If
  
' DIESEL ALTs MAN and AUTO
  If elec_lamps(i + 24) Then   ' AUTOLP=true
     MAN(i).BackColor = &HC0C0C0       ' GREY
     AUTO(i).BackColor = &HFF00FF      ' PURPLE
  Else
     MAN(i).BackColor = &HFFFF&        ' YELLOW
     AUTO(i).BackColor = &HC0C0C0      ' GREY
  End If
  
' DIESEL ALTs STAND-BY 1
  If elec_lamps(i + 26) Then   ' STBY1LP=true
     SBY1(i).BackColor = &HFFFF&        ' YELLOW
  Else
     SBY1(i).BackColor = &HC0C0C0       ' GREY
  End If
  
' DIESEL ALTs STAND-BY 2
  If elec_lamps(i + 28) Then   ' STBY2LP=true
     SBY2(i).BackColor = &HFFFF&        ' YELLOW
  Else
     SBY2(i).BackColor = &HC0C0C0       ' GREY
  End If
  
Next i


' Note: update_pumps(syxxx,index,nPump)
Call update_pump(Elec_ctrl, 3, 41)      ' D/A 1 PRE-LUB
'Call update_pump(Elec_ctrl, 4, 42)      ' D/A 2 PRE-LUB

End Sub
