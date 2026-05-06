VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Begin VB.Form Elec_ctrl_LSS 
   Caption         =   "Electrical Control Panel"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4140
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   4140
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4275
      Left            =   60
      Top             =   60
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   7541
      Begin VB.Frame DA1_frame 
         Caption         =   "SSDG"
         Height          =   2535
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1215
         Begin VB.CommandButton SBY2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "S/B 2"
            Height          =   255
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton SBY1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "S/B"
            Height          =   255
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1740
            Width           =   615
         End
         Begin VB.CommandButton AUTO 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AUTO"
            Height          =   315
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Check Synch MUST be ON and air and fuel MUST be available"
            Top             =   1380
            Width           =   615
         End
         Begin VB.CommandButton MAN 
            BackColor       =   &H00C0C0C0&
            Caption         =   "MAN"
            Height          =   255
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   255
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   780
            Width           =   615
         End
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ON"
            Height          =   255
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame DA1_frame 
         Caption         =   "EDA"
         Height          =   1635
         Index           =   0
         Left            =   1260
         TabIndex        =   16
         Top             =   840
         Width           =   1215
         Begin VB.CommandButton sbypump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AUTO"
            Height          =   315
            Index           =   0
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   255
            Index           =   0
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   780
            Width           =   615
         End
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ON"
            Height          =   255
            Index           =   0
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1980
         Top             =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Pre-Lub"
         Height          =   1635
         Left            =   0
         TabIndex        =   3
         Top             =   2520
         Width           =   3855
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ON"
            Height          =   255
            Index           =   3
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   660
            Width           =   615
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   255
            Index           =   3
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   960
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
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1320
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ON"
            Height          =   255
            Index           =   4
            Left            =   1620
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   660
            Width           =   615
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   255
            Index           =   4
            Left            =   1620
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1320
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ON"
            Height          =   255
            Index           =   2
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   660
            Width           =   615
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   255
            Index           =   2
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
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
            Index           =   2
            Left            =   2880
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1320
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            Caption         =   "SSDG"
            Height          =   195
            Left            =   420
            TabIndex        =   15
            Top             =   360
            Width           =   555
         End
         Begin VB.Label PRE_lub_label 
            Caption         =   "SSTG#1"
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   14
            Top             =   360
            Width           =   675
         End
         Begin VB.Label PRE_lub_label 
            Caption         =   "SSTG#2"
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   13
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Shore Supply"
         Height          =   1635
         Left            =   2580
         TabIndex        =   0
         Top             =   840
         Width           =   1215
         Begin VB.CommandButton ONpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ON/OFF"
            Height          =   255
            Index           =   5
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   660
            Width           =   795
         End
         Begin VB.CommandButton OFFpump 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   255
            Index           =   5
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         LineCol         =   16711680
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   4
         Left            =   3180
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         LineCol         =   16711680
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   2
         Left            =   3660
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         LineCol         =   16711680
      End
      Begin VB.TextBox dummybox 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   540
         TabIndex        =   27
         Top             =   60
         Width           =   435
      End
   End
End
Attribute VB_Name = "Elec_ctrl_LSS"
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


' 5/1/00
' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


If Index = 1 Then
    Call setPCDI(EAUTOSW(1))    ' D/A 1 to AUTO
ElseIf Index = 2 Then
    Call setPCDI(EAUTOSW(2))    ' D/A 2 to AUTO
End If


End Sub

Private Sub Command1_Click()


End Sub

Private Sub Form_Load()

   set_dummy_focus = True
   
    Me.Height = elec_ctrl_height
    Me.Width = elec_ctrl_width

' === 30/09/02 17.008 ===
If Not I_AM_INSTRUCTOR And Not PTT Then
    ' I MUST be an ACT Terminal (Full Mission) so
    ONpump(4).Visible = False
    ONpump(2).Visible = False
    OFFpump(4).Visible = False
    OFFpump(2).Visible = False
    PRE_lub_label(0).Visible = False
    PRE_lub_label(1).Visible = False
Else
    ONpump(4).Visible = True
    ONpump(2).Visible = True
    OFFpump(4).Visible = True
    OFFpump(2).Visible = True
    PRE_lub_label(0).Visible = True
    PRE_lub_label(1).Visible = True
End If
' =======================



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

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

' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If



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
   
' 5/1/00
' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
   

If Index = 0 Then
  ' 21/9/99
  'Call setPCDI(789)    ' Emerg ALT OFF
   Call setPCDI(745)    ' Emerg ALT OFF
ElseIf Index = 1 Then

    Call setPCDI(ESTOPSW(1))    ' STOP SSDG
    
ElseIf Index = 2 Then

         Call setPCDI(PumpIO_LSS(4, 21))    ' STOP SSTG # 2 Pre-lub pump
   
ElseIf Index = 3 Then
'        Note:  PumpIO(4=OFF,nPump)
         Call setPCDI(PumpIO_LSS(4, 8))    ' STOP SSDG Pre-lub pump
         
ElseIf Index = 4 Then

'        Note:  PumpIO(4=OFF,nPump)
         Call setPCDI(PumpIO_LSS(4, 20))    ' STOP SSTG # 1 Pre-lub pump

ElseIf Index = 5 Then

    ss_supply = False
    
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 5/1/00
' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If


Select Case Index
  Case 0
     ' 21/9/99
       ' Call setPCDI(781)           ' Emerg ALT ON
         Call setPCDI(746)           ' Emerg ALT ON
  Case 1
         Call setPCDI(ESTARTSW(1))    ' START SSDG
         
  Case 2
        
        ' 23/11/99
        Call setPCDI(PumpIO_LSS(2, 21))  '  START SSTG # 2 Pre-lub pump
       
  Case 3
      '  Note:  PumpIO(2=ON,nPump)
         Call setPCDI(PumpIO_LSS(2, 8))   ' START SSDG Pre-lub pump
         
  Case 4
        ' Note:  PumpIO(2=ON,nPump)
        
        ' 23/11/99
        Call setPCDI(PumpIO_LSS(2, 20))  '  START SSTG # 1 Pre-lub pump
        
  Case 5
         ss_supply = True

End Select


' === 25/03/02 REMOVED the next line ===
'       Call setPCDI(PumpIO_LSS(2, 20))


End Sub

Private Sub SBY1_Click(Index As Integer)

set_dummy_focus = True
Control_PC_check
If This_Control_status = False Then Exit Sub


' 5/1/00
' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

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


' 5/1/00
' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

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
   
' 5/1/00
' 11/8/00
'If Not PTT Then Exit Sub
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
   
If Index = 0 Then
    ' 21/9/99
    'Call setPCDI(785)    ' Emerg ALT AUTO
     Call setPCDI(747)    ' Emerg ALT AUTO
End If

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim single_char As String * 1

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' Emergency ALT ON and OFF
'If elec_lamps(14) Then
single_char = Mid$(PC_INT_STRING, 394, 1)
If single_char = "1" Then
   ONpump(0).BackColor = &HFF00&      ' GREEN
   OFFpump(0).BackColor = &HC0C0C0    ' GREY
'   ALT_image(0).Picture = GEN_on.Picture
Else
   ONpump(0).BackColor = &HC0C0C0     ' GREY
   OFFpump(0).BackColor = &HFFFFFF    ' WHITE
'   ALT_image(0).Picture = GEN_off.Picture
End If

' Emergency ALT AUTO
'If elec_lamps(16) Then
single_char = Mid$(PC_INT_STRING, 395, 1)
If single_char = "1" Then
   sbypump(0).BackColor = &HFF80FF     ' LIGHT PURPLE
Else
   sbypump(0).BackColor = &HC0C0C0     ' GREY
End If

' SHORE SUPPLY ON and OFF
'If elec_lamps(18) Then
single_char = Mid$(PC_INT_STRING, 415, 1)
If single_char = "1" Then
   ONpump(5).BackColor = &HFF00&      ' GREEN
   OFFpump(5).BackColor = &HC0C0C0    ' GREY
Else
   ONpump(5).BackColor = &HC0C0C0     ' GREY
   OFFpump(5).BackColor = &HFFFFFF    ' WHITE
End If


'For i = 1 To 2
For i = 1 To 1

' DIESEL ALTs ON and OFF
'  If elec_lamps(i + 20) Then   ' STARTLP=true
     single_char = Mid$(PC_INT_STRING, 425 + i, 1)
     If single_char = "2" Or single_char = "3" Then
     ONpump(i).BackColor = &HFF00&      ' GREEN
     OFFpump(i).BackColor = &HC0C0C0    ' GREY
  Else
     ONpump(i).BackColor = &HC0C0C0     ' GREY
     OFFpump(i).BackColor = &HFFFFFF    ' WHITE
  End If
  
' DIESEL ALTs MAN and AUTO
'  If elec_lamps(i + 24) Then   ' AUTOLP=true
'    single_char = Mid$(PC_INT_STRING, 400 + i, 1)
     single_char = Mid$(PC_INT_STRING, 402 + i, 1)
     If single_char = "1" Then
     MAN(i).BackColor = &HC0C0C0       ' GREY
     AUTO(i).BackColor = &HFF00FF      ' PURPLE
  Else
     MAN(i).BackColor = &HFFFF&        ' YELLOW
     AUTO(i).BackColor = &HC0C0C0      ' GREY
  End If
  
' DIESEL ALTs STAND-BY 1
'  If elec_lamps(i + 26) Then   ' STBY1LP=true
     single_char = Mid$(PC_INT_STRING, 404 + i, 1)
     If single_char = "1" Then
     SBY1(i).BackColor = &HFFFF&        ' YELLOW
  Else
     SBY1(i).BackColor = &HC0C0C0       ' GREY
  End If
  
' DIESEL ALTs STAND-BY 2
'  If elec_lamps(i + 28) Then   ' STBY2LP=true
     single_char = Mid$(PC_INT_STRING, 406 + i, 1)
     If single_char = "1" Then
     SBY2(i).BackColor = &HFFFF&        ' YELLOW
  Else
     SBY2(i).BackColor = &HC0C0C0       ' GREY
  End If
  
Next i


' Note: update_pumps(syxxx,index,nPump)
Call update_pump_LSS(Elec_ctrl_LSS, 3, 8)    ' SSDG PRE-LUB

' 23/11/99
Call update_pump_LSS(Elec_ctrl_LSS, 4, 20)   ' SSTG # 1 PRE-LUB
' 23/11/99
Call update_pump_LSS(Elec_ctrl_LSS, 2, 21)   ' SSTG # 2 PRE-LUB

End Sub
