VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form FEED_Pump_Control 
   Caption         =   "MECD Feed Pump Control"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9645
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   5355
      Left            =   0
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   9446
      Begin VB.CommandButton Water_IN_CLOSE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CLOSE"
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
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3420
         Width           =   675
      End
      Begin VB.CommandButton Water_IN_OPEN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPEN"
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
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3120
         Width           =   675
      End
      Begin VB.CommandButton Steam_OUT_CLOSE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CLOSE"
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
         Index           =   1
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Width           =   675
      End
      Begin VB.CommandButton Steam_OUT_OPEN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPEN"
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
         Index           =   1
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2220
         Width           =   675
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   6000
         Top             =   0
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2700
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2700
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2400
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2700
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton AUTO_VV 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
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
         Index           =   1
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton AUTO_VV 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
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
         Index           =   2
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton OPEN_VV 
         BackColor       =   &H0000FF00&
         Caption         =   "OPEN"
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
         Index           =   1
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton OPEN_VV 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPEN"
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
         Index           =   2
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton CLOSE_VV 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
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
         Index           =   1
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   675
      End
      Begin VB.CommandButton CLOSE_VV 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
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
         Index           =   2
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Steam Reg Outlet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1020
         TabIndex        =   20
         Top             =   1440
         Width           =   555
      End
      Begin VB.Shape Feed_Steam_IN_CLOSED 
         BackColor       =   &H00000040&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   315
         Index           =   0
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   315
      End
      Begin VB.Shape Feed_Steam_OUT_OPEN 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   315
         Index           =   0
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4395
         Left            =   180
         Top             =   480
         Width           =   5655
      End
      Begin VB.Shape Feed_Steam_IN_CLOSED 
         BackColor       =   &H00000040&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   315
         Index           =   1
         Left            =   480
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   315
      End
      Begin VB.Shape Feed_Steam_IN_OPEN 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   315
         Index           =   1
         Left            =   480
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   315
      End
      Begin VB.Shape Feed_Water_OUT_CLOSED 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   315
         Index           =   1
         Left            =   4380
         Shape           =   3  'Circle
         Top             =   3420
         Width           =   315
      End
      Begin VB.Shape Feed_Water_OUT_OPEN 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   315
         Index           =   1
         Left            =   4380
         Shape           =   3  'Circle
         Top             =   3060
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Steam Reg Inlet "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Feed WaterOutlet"
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
         Left            =   4020
         TabIndex        =   18
         Top             =   3840
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Feed Water Inlet"
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
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Steam Exh "
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
         Left            =   3960
         TabIndex        =   14
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label RECIRC_label 
         Alignment       =   2  'Center
         Caption         =   "Reciculation Valve # 1"
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
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label RECIRC_label 
         Alignment       =   2  'Center
         Caption         =   "Reciculation Valve # 2"
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
         Index           =   2
         Left            =   6780
         TabIndex        =   7
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Main Feed Pump # 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   3120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FEED_Pump_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub AUTO_VV_Click(Index As Integer)

set_dummy_focus = True

Select Case Index

  Case 1
    Call setPCDI(11)
  Case 2
    Call setPCDI(99)
    
End Select


End Sub

Private Sub CLOSE_VV_Click(Index As Integer)

set_dummy_focus = True

'    CLOSE
Select Case Index

  Case 1
    Call setPCDI(12)
  Case 2
    Call setPCDI(100)
    
End Select

End Sub

Private Sub Form_Load()

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub ONpump_Click(Index As Integer)


' === 09/11/04 ===
Select Case Index
    Case 0
                Call setPCDI(DI_ON_VVIO_LSS(84))
    Case 1
                Call setPCDI(DI_ON_VVIO_LSS(83))
End Select

End Sub

Private Sub OPEN_VV_Click(Index As Integer)

set_dummy_focus = True


' === 09/11/04 ===   OPEN
Select Case Index

  Case 1
    Call setPCDI(10)
  Case 2
    Call setPCDI(98)
    
End Select

End Sub

Private Sub Steam_IN_OPEN_Click(Index As Integer)

End Sub

Private Sub Steam_OUT_CLOSE_Click(Index As Integer)


set_dummy_focus = True


' === 09/11/04 ===   OPEN
Select Case Index

  Case 1
    Call setPCDI(16)
  Case 2
    Call setPCDI(113)
    
End Select

End Sub

Private Sub Steam_OUT_OPEN_Click(Index As Integer)

set_dummy_focus = True


' === 09/11/04 ===   OPEN
Select Case Index

  Case 1
    Call setPCDI(15)
  Case 2
    Call setPCDI(112)
    
End Select


End Sub

Private Sub Timer1_Timer()


' === Feed Pump re-circ valves===
' OPEN re-circ valves
If vv_status_LSS(85) = 3 Then
  OPEN_VV(1).BackColor = &HFF00&            ' GREEN
Else
  OPEN_VV(1).BackColor = &HC0C0C0           ' GREY
End If
If vv_status_LSS(86) = 3 Then
  OPEN_VV(2).BackColor = &HFF00&            ' GREEN
Else
  OPEN_VV(2).BackColor = &HC0C0C0           ' GREY
End If

' CLOSE re-circ valves
If vv_status_LSS(85) = 0 Then
  CLOSE_VV(1).BackColor = &H8080FF           ' PINK
Else
  CLOSE_VV(1).BackColor = &HC0C0C0           ' GREY
End If
If vv_status_LSS(86) = 0 Then
  CLOSE_VV(2).BackColor = &H8080FF           ' PINK
Else
  CLOSE_VV(2).BackColor = &HC0C0C0           ' GREY
End If

' Auto re-circ valves
If (D_OUTPUT(147)) Then
  AUTO_VV(1).BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_VV(1).BackColor = &HC0C0C0           ' GREY
End If
If (D_OUTPUT(495)) Then
  AUTO_VV(2).BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_VV(2).BackColor = &HC0C0C0           ' GREY
End If

'c     === 10/11/04 Feed Pump # 1 Steam REG IN IN valve (Indication Only) ===
'    if( DO(151) ) then
'        More_integers(9) = 1
'    Else
'        More_integers(9) = 0
'    End If
'
If VB_More_integers(9) > 0 Then
    Feed_Steam_IN_OPEN(1).BackColor = vbGreen
    Feed_Steam_IN_CLOSED(1).BackColor = &H40&       ' DARK RED
Else
    Feed_Steam_IN_OPEN(1).BackColor = &H4000&       ' DARK GREEN
    Feed_Steam_IN_CLOSED(1).BackColor = vbRed
End If
'c     === 10/11/04 Feed Pump # 1 Water OUT valve (Indication Only) ===
'    if( DO(163) ) then
'        More_integers(10) = 1
'    Else
'        More_integers(10) = 0
'    End If
If VB_More_integers(10) > 0 Then
    Feed_Water_OUT_OPEN(1).BackColor = vbGreen
    Feed_Water_OUT_CLOSED(1).BackColor = &H40&      ' DARK RED
Else
    Feed_Water_OUT_OPEN(1).BackColor = &H4000&      ' DARK GREEN
    Feed_Water_OUT_CLOSED(1).BackColor = vbRed
End If
'     === 10/11/04 Feed Pump # 1 Steam REG OUT valve (Indication Only) ===
If vv_status_LSS(97) = 3 Then
Else
End If



End Sub

Private Sub Water_IN_CLOSE_Click(Index As Integer)

set_dummy_focus = True


' === 09/11/04 ===   OPEN
Select Case Index

  Case 1
    Call setPCDI(18)
  Case 2
    Call setPCDI(115)
    
End Select


End Sub

Private Sub Water_IN_OPEN_Click(Index As Integer)

set_dummy_focus = True


' === 09/11/04 ===   OPEN
Select Case Index

  Case 1
    Call setPCDI(17)
  Case 2
    Call setPCDI(114)
    
End Select


End Sub

