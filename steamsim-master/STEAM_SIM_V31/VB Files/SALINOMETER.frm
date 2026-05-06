VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{17FE8C5E-5183-11D3-AA4D-00C04FA979C0}#7.0#0"; "HorizontalGauge.ocx"
Begin VB.Form SALINOMETER 
   Caption         =   "SALINOMETER"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   4830
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   8281
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1800
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   60
         Top             =   900
      End
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2820
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1980
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton SALLY 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin HorizontalGauge.HGauge Salinometer_gauge 
         Height          =   555
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Salinity ppm"
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   979
         MajorTicks      =   6
         MinorTicks      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Test"
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   16
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "SALINITY    ppm"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   14
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Distilled Water Tank"
         Height          =   255
         Index           =   5
         Left            =   1260
         TabIndex        =   13
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Distilling Plant"
         Height          =   255
         Index           =   4
         Left            =   1260
         TabIndex        =   12
         Top             =   2820
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Port Boiler Drum"
         Height          =   255
         Index           =   3
         Left            =   1260
         TabIndex        =   11
         Top             =   2400
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Starboard Boiler Drum"
         Height          =   255
         Index           =   2
         Left            =   1260
         TabIndex        =   10
         Top             =   1980
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "Atmospheric Drain Tank"
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SALINITY MONITOR"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   4020
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4335
         Left            =   660
         Top             =   60
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Condenser Hotwell"
         Height          =   255
         Index           =   0
         Left            =   1260
         TabIndex        =   2
         Top             =   1140
         Width           =   1935
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "SALINOMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer

Private Sub Form_Load()


On Error Resume Next
Call JYAScreenInit(Me, idSalinometer)


    Salinometer_gauge.GaugeText 1, "0"
    Salinometer_gauge.GaugeText 2, "1"
    Salinometer_gauge.GaugeText 3, "2"
    Salinometer_gauge.GaugeText 4, "3"
    Salinometer_gauge.GaugeText 5, "4"
    Salinometer_gauge.GaugeText 6, "5"
    
    Salinometer_gauge.Value = 0    ' ZERO

    
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)
  
End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idSalinometer)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub SALLY_Click(Index As Integer)

' 14/9/00
set_dummy_focus = True

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

  Call setPCDI(178 + Index)

End Sub

Private Sub Timer1_Timer()
Dim sally_gauge As Single


'c   === 14/9/00 ===
'    equivalence(STEAM_logicals(156),SALINITY_lamp(1))
'    logical*1           SALINITY_lamp(6)

Dim i As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' 14/9/00

'c     SALINITY_gauge_value is in ppm
'              PC_REALS(585) = SALINITY_gauge_value
sally_gauge = LSS_reals(585) * 255# / 5#

If sally_gauge > 255# Then
  Salinometer_gauge.Value = 255
ElseIf sally_gauge < 0# Then
  Salinometer_gauge.Value = 0
Else
  Salinometer_gauge.Value = sally_gauge
End If

' 14/9/00
For i = 0 To 5
  If VB_steam_logicals(156 + i) Then
    SALLY(i).BackColor = &HFF00&          ' green
  Else
    SALLY(i).BackColor = &H808080     ' grey
  End If
Next i



End Sub
