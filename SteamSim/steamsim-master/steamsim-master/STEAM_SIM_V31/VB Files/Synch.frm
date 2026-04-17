VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form frm_synch 
   Caption         =   "Synchroscope and Synchronising Phasors"
   ClientHeight    =   6855
   ClientLeft      =   1500
   ClientTop       =   1830
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6855
   ScaleWidth      =   8655
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10821
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   120
         Top             =   3600
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   60
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Text2"
         Top             =   300
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   900
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox Text5 
         Height          =   255
         Left            =   0
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton SynchExit 
         Caption         =   "Exit"
         Height          =   675
         Left            =   6540
         TabIndex        =   30
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On line"
         Height          =   675
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5160
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LOCK_SCOPE"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RUNALL"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2400
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Synchroscope / Phasor Diagram"
         Height          =   675
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5160
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Caption         =   "Incoming Machine"
         Height          =   1635
         Left            =   1260
         TabIndex        =   19
         Top             =   4080
         Width           =   1515
         Begin VB.CommandButton ALT_select 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SSTG 1"
            Height          =   315
            Index           =   3
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox dummybox 
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   480
            TabIndex        =   25
            Top             =   600
            Width           =   195
         End
         Begin VB.CommandButton ALT_select 
            Height          =   315
            Index           =   4
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton ALT_select 
            Caption         =   "SSDG"
            Height          =   315
            Index           =   1
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   900
            Width           =   735
         End
         Begin VB.CommandButton ALT_select 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SSTG 2"
            Height          =   315
            Index           =   2
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   300
            Width           =   735
         End
         Begin VB.CommandButton ALT_select 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OFF"
            Height          =   315
            Index           =   0
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Check Synch"
         Height          =   795
         Left            =   2820
         TabIndex        =   16
         Top             =   4080
         Width           =   1395
         Begin VB.CommandButton Check_OFF 
            Caption         =   "OFF"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   300
            Width           =   555
         End
         Begin VB.CommandButton Check_ON 
            Caption         =   "ON"
            Height          =   315
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame frame3 
         Caption         =   "Synchroscope"
         Height          =   795
         Left            =   2820
         TabIndex        =   13
         Top             =   4920
         Width           =   1395
         Begin VB.CommandButton Synch_OFF 
            Caption         =   "OFF"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   300
            Width           =   555
         End
         Begin VB.CommandButton Synch_ON 
            Caption         =   "ON"
            Height          =   315
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame frame4 
         Caption         =   "Parameters"
         Height          =   2655
         Left            =   4440
         TabIndex        =   4
         Top             =   540
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Bus-Bar"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   420
            Width           =   615
         End
         Begin VB.Label VRUN_label 
            Caption         =   "xxx V"
            Height          =   195
            Left            =   960
            TabIndex        =   11
            Top             =   420
            Width           =   915
         End
         Begin VB.Label FRUN_label 
            Caption         =   "xx.x Hz"
            Height          =   195
            Left            =   960
            TabIndex        =   10
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Incoming"
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label VINC_label 
            Caption         =   "xxx V"
            Height          =   195
            Left            =   960
            TabIndex        =   8
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label FINC_label 
            Caption         =   "xx.x Hz"
            Height          =   195
            Left            =   960
            TabIndex        =   7
            Top             =   1620
            Width           =   915
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase Diff"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label PH_diff 
            Caption         =   "xxx deg"
            Height          =   195
            Left            =   1020
            TabIndex        =   5
            Top             =   2040
            Width           =   795
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FAIL_SCOPE"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox Text6 
         Height          =   255
         Left            =   60
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Text6"
         Top             =   1500
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text7 
         Height          =   255
         Left            =   60
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "Text7"
         Top             =   1740
         Visible         =   0   'False
         Width           =   1635
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   315
         Index           =   4
         Left            =   3480
         TabIndex        =   35
         Top             =   540
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   327682
         Max             =   100
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   315
         Index           =   3
         Left            =   3420
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   327682
         Max             =   100
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Index           =   2
         Left            =   5160
         TabIndex        =   37
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   327682
         Max             =   100
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Index           =   1
         Left            =   5100
         TabIndex        =   38
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   327682
         Max             =   1300
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Index           =   0
         Left            =   5100
         TabIndex        =   39
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   327682
         Max             =   1300
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Index           =   5
         Left            =   4500
         TabIndex        =   40
         Top             =   4440
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   873
         _Version        =   327682
         Max             =   100
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Index           =   6
         Left            =   5760
         TabIndex        =   41
         Top             =   4440
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   873
         _Version        =   327682
         Max             =   100
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   495
         Index           =   7
         Left            =   60
         TabIndex        =   42
         Top             =   4800
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   873
         _Version        =   327682
         Max             =   1000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   2580
         X2              =   2580
         Y1              =   960
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   2580
         X2              =   1260
         Y1              =   2400
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   3900
         X2              =   2580
         Y1              =   3360
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         Index           =   3
         X1              =   2100
         X2              =   2580
         Y1              =   1080
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   1860
         X2              =   2580
         Y1              =   3600
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Index           =   5
         X1              =   4020
         X2              =   2580
         Y1              =   3000
         Y2              =   2400
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incoming voltage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   4140
         TabIndex        =   51
         Top             =   300
         Width           =   2235
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Voltage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   50
         Top             =   1500
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency difference"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   4140
         TabIndex        =   49
         Top             =   2820
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label Fdiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fdiff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6480
         TabIndex        =   48
         Top             =   2820
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label V_ref 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "V_ref"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6480
         TabIndex        =   47
         Top             =   1500
         Width           =   870
         WordWrap        =   -1  'True
      End
      Begin VB.Label V_inc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "V_inc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         TabIndex        =   46
         Top             =   300
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Phase difference"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   4140
         TabIndex        =   45
         Top             =   4140
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Phase_diff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Phdiff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6300
         TabIndex        =   44
         Top             =   4140
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   4
         Height          =   3375
         Left            =   960
         Top             =   540
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "SYNCHROSCOPE"
         Height          =   195
         Left            =   2040
         TabIndex        =   43
         Top             =   600
         Width           =   1515
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   2535
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   1140
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frm_synch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim set_dummy_focus As Integer
' === 22/01/02 15.047 ===
Private shape2_original_width As Integer
Private shape2_new_width As Integer
Private Stretch_factor As Single

Private Sub ALT_select_Click(Index As Integer)
Dim i As Integer

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 12/3/98
For i = 1 To 4

      
     If I_AM_HOST Or Not In_a_group Then
     
     Else ' I am a CLIENT So we need to sleep a while....
          ' === 25/03/02 Messages sent too quickly from Client to Host ===
          Call Sleep(200)
     End If
     
     
      If i = Index Then
        Call setDI(SYNCHSELSW(i), True)       ' SELECT Incoming Freq
      Else
        Call setDI(SYNCHSELSW(i), False)      ' DE-SELECT Incoming Freq
      End If
      
Next i

End Sub

Private Sub Check_OFF_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

Call setDI(CHSYNCHSW, False)   ' Check Synch OFF

End Sub

Private Sub Check_ON_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

Call setDI(CHSYNCHSW, True)   ' Check Synch ON

End Sub

Private Sub Command1_Click()
Dim i As Integer


synch_In_a_group = Not synch_In_a_group
     
     
Call setup_screen

' 26/8/98 CRASHES but why?
' set_dummy_focus = True


End Sub



Private Sub Command2_Click()

LOCK_SCOPE = Not LOCK_SCOPE

End Sub

Private Sub Command3_Click()

RUNALL = Not RUNALL

End Sub

Private Sub Command4_Click()


synchroscope = Not synchroscope

Call setup_screen

' 26/8/98  CRASHES but why?
'set_dummy_focus = True

End Sub

Private Sub Command5_Click()

FAIL_SCOPE = Not FAIL_SCOPE

End Sub

Private Sub Form_Load()
Dim i As Integer


' === 22/01/02 Moved from below ===
frm_synch.Height = 6360
frm_synch.ScaleHeight = 5955

frm_synch.Width = 7560
frm_synch.ScaleWidth = 7440


' === 22/01/02 ===
On Error Resume Next
Call JYAScreenInit(Me, id_Synchro)

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


' === 22/01/02 ===
'Private shape2_original_width As Integer
'Private shape2_new_width As Integer
'Private Stretch_factor As Single
shape2_original_width = Shape2.Width
Stretch_factor = 1#


' 21/8/98
If medium_speed Then
   ALT_select(0).Visible = False
   ALT_select(3).Caption = "OFF"
Else
   ALT_select(0).Visible = True
   
'   22/11/99
'  ALT_select(3).Caption = "T/A"
   ALT_select(3).Caption = "SSTG 1"
   
End If


' 11/3/98
   AIPHASE_model = 0
   
' 26/3/98
   AIPHASE_model_old = -123    ' To ensure an update
   
' 26/3/98
Tx_delay = 860

' 10/3/98
            synch_count = 5
  
' === 22/01/02 Moved up ===
'frm_synch.Height = 6360
'frm_synch.ScaleHeight = 5955
'
'frm_synch.Width = 7560
'frm_synch.ScaleWidth = 7440



' 28/2/97   Initial conditions:

AIPHASE = -10#

Static_radius = 1500
Moving_radius = 1500

' === 22/01/02 ===
'Private shape2_original_width As Integer
'Private shape2_new_width As Integer
'Private Stretch_factor As Single

'centreX = 2520
'centreY = 2400
centreX = 2520 * Stretch_factor
centreY = 2400 * Stretch_factor

 Slider2(0).Value = 1500
 Slider2(1).Value = 1500
 Slider2(2).Value = 25

' 10/3/98
Call setup_screen

End Sub


Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, id_Synchro)
    HavenStretch1.Visible = True
End If

' === 22/01/02 ===
'Private shape2_original_width As Integer
'Private shape2_new_width As Integer
'Private Stretch_factor As Single
shape2_new_width = Shape2.Width
Stretch_factor = shape2_new_width / shape2_original_width

centreX = 2520 * Stretch_factor
centreY = 2400 * Stretch_factor


End Sub

Private Sub Slider2_Change(Index As Integer)

If Index = 3 Then
   VINC = Slider2(3).Value * 4#
ElseIf Index = 4 Then
   VRUN = Slider2(4).Value * 4#
ElseIf Index = 5 Then
   fInc = Slider2(5).Value * 0.1 + 45#
ElseIf Index = 6 Then
   FRUN = Slider2(6).Value * 0.1 + 45#
ElseIf Index = 7 Then
   Tx_delay = Slider2(7).Value
End If


End Sub

Private Sub Synch_OFF_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub

' 3 WAY SWITCH:
Call setDI(SYNCHOFFSW, True)    '  Synchroscope OFF
Call setDI(SELSYNCHSW, False)
Call setDI(SELLAMPSSW, False)

End Sub

Private Sub Synch_ON_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 3 WAY SWITCH:
Call setDI(SYNCHOFFSW, False)
Call setDI(SELSYNCHSW, True)    '  Synchroscope ON
Call setDI(SELLAMPSSW, False)

End Sub

Private Sub SynchExit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim Aradians As Single
Dim i As Integer
Dim IAI_PHASE As Integer

' 24/11/99
Dim AI_val As Single
Dim IAI_val As Integer


' 21/8/98
Dim Alt_selected As Integer

' 2/5/97 now GLOBAL Dim speed As Single
Dim AI_STEP As Single
  
If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

'Tx_delay = 1000#    '    A wild guess
Tx_delay = Slider2(7).Value


Text3.text = "Tx_delay = " & Tx_delay

' Get the values from the Fortran model
If medium_speed Then
    VRUN = elec_reals(29)      ' VEMBS
    VINC = elec_reals(61)      ' VVESYI
    FRUN = elec_reals(30)      ' HEMBS
    fInc = elec_reals(60)      ' HHESYI
Else
    VRUN = LSS_reals(712)
    VINC = LSS_reals(714)
    FRUN = LSS_reals(713)
    fInc = LSS_reals(715)
End If


' 8/6/98
If medium_speed Then
   ' 11/3/98
   If elec_lamps(31) Then
      LOCK_SCOPE = True
   Else
      LOCK_SCOPE = False
   End If
   If elec_lamps(32) Then
      FAIL_SCOPE = True
   Else
      FAIL_SCOPE = False
   End If
   If elec_lamps(41) Then
      RUNALL = True
   Else
      RUNALL = False
   End If
Else
   If LSS_LOGICALS(9) Then
      LOCK_SCOPE = True
   Else
      LOCK_SCOPE = False
   End If
   If LSS_LOGICALS(11) Then
      FAIL_SCOPE = True
   Else
      FAIL_SCOPE = False
   End If
   If LSS_LOGICALS(10) Then
      RUNALL = True
   Else
      RUNALL = False
   End If
End If


' 21/8/98
Alt_selected = False

For i = 1 To 4

  If medium_speed Then
  
   ' Synchroscope selection ( i=3=OFF , 1=DA1 , 2=DA2 , 4=SHAFT )
    If elec_lamps(i + 32) Then
         ALT_select(i).BackColor = &HFF00&      ' GREEN
    Else
         ALT_select(i).BackColor = &HC0C0C0     ' GREY
    End If
    
  Else
  
   ' Synchroscope selection ( i=3=T/A , 1=DA1 , 2=DA2 , 4=SHAFT )
    If LSS_LOGICALS(i + 11) Then
        Alt_selected = True
         ALT_select(i).BackColor = &HFF00&      ' GREEN
    Else
         ALT_select(i).BackColor = &HC0C0C0     ' GREY
    End If
    
  End If

Next i

' 21/8/98
If medium_speed = False Then
    If Alt_selected Then  ' ( Slow Speed Incoming m/c OFF )
          ALT_select(0).BackColor = &HC0C0C0     ' GREY
    Else
          ALT_select(0).BackColor = &HFF00&      ' GREEN
    End If
End If

' Synchroscope ON/OFF
If medium_speed Then
  If elec_lamps(39) Then    ' SYNCHROSCOPE is SELECTED
    Synch_ON.BackColor = &HFF00&       ' GREEN
    Synch_OFF.BackColor = &HC0C0C0       ' GREY
  Else
    Synch_OFF.BackColor = &HFF00&       ' GREEN
    Synch_ON.BackColor = &HC0C0C0       ' GREY
  End If
Else
  If LSS_LOGICALS(19) Then    ' SYNCHROSCOPE is SELECTED
    Synch_ON.BackColor = &HFF00&       ' GREEN
    Synch_OFF.BackColor = &HC0C0C0       ' GREY
  Else
    Synch_OFF.BackColor = &HFF00&       ' GREEN
    Synch_ON.BackColor = &HC0C0C0       ' GREY
  End If
End If

' Check Scope ON/OFF
If medium_speed Then
  If elec_lamps(37) Then    ' CHECK SCOPE is ON
    Check_OFF.BackColor = &HC0C0C0     ' GREY
    Check_ON.BackColor = &HFF00&       ' GREEN
  Else
    Check_OFF.BackColor = &HFF00&       ' GREEN
    Check_ON.BackColor = &HC0C0C0       ' GREY
  End If
Else
  If LSS_LOGICALS(16) Then    ' CHECK SCOPE is ON
    Check_OFF.BackColor = &HC0C0C0     ' GREY
    Check_ON.BackColor = &HFF00&       ' GREEN
  Else
    Check_OFF.BackColor = &HFF00&       ' GREEN
    Check_ON.BackColor = &HC0C0C0       ' GREY
  End If
End If



If synch_In_a_group Then

        
        ' === 22/02/06 REMOVE NEXT LINE... ===
        ' If I_AM_HOST Then     ' 26/3/98
         
              synch_count = synch_count + 1
              If synch_count >= 10 Then
                   ' 12/3/98
                    If AIPHASE > 0# And AIPHASE < 256# Then
                        IAI_PHASE = AIPHASE
                        
                        If medium_speed Then
                          Call setAI(59, IAI_PHASE)    ' Every 500mS
                        Else
                        
                            ' 24/11/99  LSS Diesel uses AI(35) , STEAM_SIM uses AI(47)
'                               Call setAI(35, IAI_PHASE)    ' Every 500mS

                            If (PTT) Then
 '                              Note: WieldMuller AI(47) is approx 88 at 12 o'clock ( NOT 127 )
 '                                  AI_val = VB_AI(47) * 127# / 88#
                                    AI_val = IAI_PHASE * 88# / 127#
                                    IAI_val = AI_val
                                Call setAI(47, IAI_val)    ' Every 500mS
                                
                            End If
                            
                       End If
                        
                        
                    End If
                 synch_count = 0
              End If
              
        ' === 22/02/06 REMOVE NEXT LINE... ===
        ' End If
              
   If medium_speed Then
      VRUN_label.Caption = Format(elec_reals(29), "##0") & " V"
      VINC_label.Caption = Format(elec_reals(61), "##0") & " V"
      FRUN_label.Caption = Format(elec_reals(30), "##0.00") & " Hz"
      FINC_label.Caption = Format(elec_reals(60), "##0.00") & " Hz"
   Else
      VRUN_label.Caption = Format(LSS_reals(712), "##0") & " V"
      VINC_label.Caption = Format(LSS_reals(714), "##0") & " V"
      FRUN_label.Caption = Format(LSS_reals(713), "##0.00") & " Hz"
      FINC_label.Caption = Format(LSS_reals(715), "##0.00") & " Hz"
   End If

      If LOCK_SCOPE Then
           Text4.text = "LOCK_SCOPE=True"
      Else
           Text4.text = "LOCK_SCOPE=False"
      End If
      If RUNALL Then
           Text5.text = "RUNALL=True"
      Else
           Text5.text = "RUNALL=False"
      End If

        

 
' 10/3/98     Static_radius = VRUN * 4#       ' MAX value = 2000
              Static_radius = VRUN * 4# * 0.64
         
            V_ref.Caption = Format$(VRUN, "##0") & " V"         '  volts
          
     If Not RUNALL Then
             Moving_radius = 0#      ' PARK the Scope
             Line1(0).BorderWidth = 2
             Line1(1).BorderWidth = 2
             Line1(2).BorderWidth = 2
     Else
        If LOCK_SCOPE Then            ' NOTE: LOCK_SCOPE is an inverted signal
          
 ' 10/3/98        Moving_radius = VINC * 4#          '  So display the MOVING radius
                  Moving_radius = VINC * 4# * 0.64   '  So display the MOVING radius
                
               Line1(0).BorderWidth = 2
               Line1(1).BorderWidth = 2
               Line1(2).BorderWidth = 2
              
         Else
                 Moving_radius = 0#          ' Synchro-scope is at 12 o'clock
               Line1(0).BorderWidth = 4
               Line1(1).BorderWidth = 4
               Line1(2).BorderWidth = 4
         End If
     End If
     
        V_inc.Caption = Format$(VINC, "##0") & " V"       '  volts
      
                   Speed = fInc - FRUN             ' Speed in Hz
               
         '  6/5/97
    ' 12/3/98   If speed > 0.001 And speed < 0.04 Then speed = 0.04       '  Note: Digital S'Scope resolution is 0.04 Hz
    '           If speed < -0.001 And speed > -0.04 Then speed = -0.04
                If Speed > 0# And Speed < 0.04 Then Speed = 0.04          '  Note: Digital S'Scope resolution is 0.04 Hz
                If Speed < 0# And Speed > -0.04 Then Speed = -0.04
           
         
                 ' 26/3/98
                 ' 24/11/99
                 '  AIPHASE_model = VB_pots(59)
                 
                 If medium_speed Then
                 
                    AIPHASE_model = VB_pots(59)  '  (gets transmitted around the network)
                    
                 Else
                    ' Note:
                    ' A_INPUT(47) does NOT get transmitted around the network
                    '   VB_AI(47) does get transmitted around the network
                    
                    ' === 22/02/06 ADDED Next line for Texas (complaint student stand alone can't parallel)
                    If update_VBAI(47) Then
                        update_VBAI(47) = False

                        ' === 12/12/05 REMOVE nex line and the End If===
                        'If (Not PTT) Then
                        '    Note: WieldMuller AI(47) is approx 88 at 12 o'clock ( NOT 127 )
                            AI_val = VB_AI(47) * 127# / 88#
                            AIPHASE_model = AI_val
                        
                        'End If
                        
                    End If
                    
                 End If
         
' =============== 10/3/98 TESTING (SIMULATE THE PHASE SIGNAL FROM THE SYNCHRO_SCOPE)
If medium_speed Then
  GoTo skip_test:
End If

' 24/11/99
If (PTT) Then

               AIPHASE_model = AIPHASE_model + Speed * 255# * (Timer1.Interval * 0.001)
again:
            If AIPHASE_model > 255# Then AIPHASE_model = AIPHASE_model - 255#
            If AIPHASE_model > 255# Then GoTo again:
again1:
           
            If AIPHASE_model < 0# Then AIPHASE_model = AIPHASE_model + 255#
            If AIPHASE_model < 0# Then GoTo again1:
            
End If

skip_test:
                 Text1.text = AIPHASE
                 Text2.text = AIPHASE_model
 ' ===============

            If AIPHASE_model <> AIPHASE_model_old Then
          
                 AIPHASE_model_old = AIPHASE_model
                 
'  NOTE: Tx_delay is the transmission time in milli-seconds from the FORTRAN model to this VB program
   ' 6/5/97
   ' 26/3/98
   '          If AIPHASE_model < 200 And AIPHASE_model > 80 Then
   '              AIPHASE = (AIPHASE_model + AIPHASE_model_old) * 0.5 + speed * 255# * (Tx_delay * 0.001)
   '          End If
              
                  AIPHASE = (AIPHASE_model + AIPHASE_model_old) * 0.5 + Speed * 255# * (Tx_delay * 0.001)
        
         
            End If
        
Else

' 10/3/98
'        Static_radius = Slider2(1).Value    ' MAX value = 2000
'        Moving_radius = Slider2(0).Value    ' MAX value = 2000

        Static_radius = Slider2(1).Value    ' MAX value = 1300
        Moving_radius = Slider2(0).Value    ' MAX value = 1300
        
'   10/3/98
'     V_ref.Caption = Format$(Static_radius * 0.25, "##0") & " V"       ' MAX value = 500 volts
'     V_inc.Caption = Format$(Moving_radius * 0.25, "##0") & " V"       ' MAX value = 500 volts
      V_ref.Caption = Format$(Static_radius * 0.3846, "##0") & " V"     ' MAX value = 500 volts
      V_inc.Caption = Format$(Moving_radius * 0.3846, "##0") & " V"     ' MAX value = 500 volts
      
        Speed = (Slider2(2).Value - 50) * 0.01              ' Speed in Hz   ( range -0.5Hz  to + 0.5Hz )
                          
            synch_count = synch_count + 1
         If synch_count >= 5 Then
             synch_count = 0
         End If
          
End If

If Abs(Speed) < 5# Then
    AI_STEP = Speed * 255# * (Timer1.Interval * 0.001)
    AIPHASE = AIPHASE + AI_STEP
    Text6.text = "speed=" & Speed
Else
    AIPHASE = 180#
    Text6.text = "speed>5.0 " & Speed
End If
                        
If AIPHASE > 255# Then AIPHASE = AIPHASE - 255#
If AIPHASE < 0# Then AIPHASE = AIPHASE + 255#
    
Fdiff.Caption = Format$(Speed, "0.00") & " Hz"
                        
                        
If synch_In_a_group Then

     If VINC > 100# And VRUN > 100# Then
       If RUNALL Then
           If LOCK_SCOPE Then
              Text7.text = "AIPHASE=" & AIPHASE
            Else
              AIPHASE = 127.5               ' Go to 12 o'clock
              Text7.text = "AIPHASE=127"
            End If
       Else
               AIPHASE = 240               ' PARK SCOPE
               Text7.text = "AIPHASEa=240"
       End If
       
     Else
          AIPHASE = 240
          Text7.text = "AIPHASEb=240"     ' PARK SCOPE
     End If
     
Else
     If Val(V_ref.Caption) > 100 And Val(V_inc.Caption) > 100 Then
         ' DO nothing
     Else
               AIPHASE = 240               ' PARK SCOPE
               Text7.text = "AIPHASEc=240"
     End If
End If
                                  
                        
' 10/3/98
If synchroscope Then
   
     Line1(0).Visible = False
     Line1(1).Visible = False
     Line1(2).Visible = False
     Line1(4).Visible = False
     Line1(5).Visible = False
     
     Moving_radius = 440# / 0.3846

Else

   Line1(0).Visible = True
   Line1(1).Visible = True
   Line1(2).Visible = True
   Line1(4).Visible = True
   Line1(5).Visible = True
   
End If

' 28/2/97  Plot the MOVING PHASORS
RED_angle = AIPHASE * 360# / 255# + 270#
If RED_angle > 360# Then RED_angle = RED_angle - 360#

BLUE_angle = RED_angle + 240#
If BLUE_angle > 360# Then BLUE_angle = BLUE_angle - 360#

YELLOW_angle = RED_angle + 120#
If YELLOW_angle > 360# Then YELLOW_angle = YELLOW_angle - 360#


'   Define the RED angle in radians ( -PI to + PI )
Aradians = (180# - RED_angle) * 3.141596 / 180#

PH_diff.Caption = Format$((AIPHASE / 255# - 0.5) * 360#, "###0") & " deg"  ' Chr$(179)          '  Phase difference in degrees
Phase_diff.Caption = Format$((AIPHASE / 255# - 0.5) * 360#, "###0") & " deg"  ' Chr$(179)          '  Phase difference in degrees



' === 22/01/02 ===
' ADDED Stretch_factor

' Plot the RED line
Line1(3).X1 = centreX
Line1(3).Y1 = centreY
Line1(3).X2 = centreX + Moving_radius * Cos(Aradians) * Stretch_factor
Line1(3).Y2 = centreY - Moving_radius * Sin(Aradians) * Stretch_factor

'   Define the BLUE angle in radians ( -PI to + PI )
Aradians = (180# - BLUE_angle) * 3.141596 / 180#
' Plot the RED line
Line1(4).X1 = centreX
Line1(4).Y1 = centreY
Line1(4).X2 = centreX + Moving_radius * Cos(Aradians) * Stretch_factor
Line1(4).Y2 = centreY - Moving_radius * Sin(Aradians) * Stretch_factor

'   Define the YELLOW angle in radians ( -PI to + PI )
Aradians = (180# - YELLOW_angle) * 3.141596 / 180#
' Plot the RED line
Line1(5).X1 = centreX
Line1(5).Y1 = centreY
Line1(5).X2 = centreX + Moving_radius * Cos(Aradians) * Stretch_factor
Line1(5).Y2 = centreY - Moving_radius * Sin(Aradians) * Stretch_factor


' 28/2/97  Plot the STATIC PHASORS
RED_angle = 90#

BLUE_angle = RED_angle + 240#
If BLUE_angle > 360# Then BLUE_angle = BLUE_angle - 360#

YELLOW_angle = RED_angle + 120#
If YELLOW_angle > 360# Then YELLOW_angle = YELLOW_angle - 360#


'   Define the RED angle in radians ( -PI to + PI )
Aradians = (180# - RED_angle) * 3.141596 / 180#
' Plot the RED line
Line1(0).X1 = centreX
Line1(0).Y1 = centreY
Line1(0).X2 = centreX + Static_radius * Cos(Aradians) * Stretch_factor
Line1(0).Y2 = centreY - Static_radius * Sin(Aradians) * Stretch_factor

'   Define the BLUE angle in radians ( -PI to + PI )
Aradians = (180# - BLUE_angle) * 3.141596 / 180#
' Plot the RED line
Line1(1).X1 = centreX
Line1(1).Y1 = centreY
Line1(1).X2 = centreX + Static_radius * Cos(Aradians) * Stretch_factor
Line1(1).Y2 = centreY - Static_radius * Sin(Aradians) * Stretch_factor

'   Define the YELLOW angle in radians ( -PI to + PI )
Aradians = (180# - YELLOW_angle) * 3.141596 / 180#
' Plot the RED line
Line1(2).X1 = centreX
Line1(2).Y1 = centreY
Line1(2).X2 = centreX + Static_radius * Cos(Aradians) * Stretch_factor
Line1(2).Y2 = centreY - Static_radius * Sin(Aradians) * Stretch_factor



End Sub
Sub setup_screen()
Dim i As Integer

If synch_In_a_group Then
        
 ' 10/3/98
         Label(0).Visible = False
         Label(2).Visible = False
         Label(3).Visible = False
         Label(4).Visible = False
         Phase_diff.Visible = False
         
         Fdiff.Visible = False
         V_ref.Visible = False
         V_inc.Visible = False
         
         Frame1.Visible = True
         Frame2.Visible = True
         Frame3.Visible = True
         Frame4.Visible = True
         
         Command1.BackColor = &HFF00&      ' GREEN
         Command1.Caption = "ON LINE"
         For i = 0 To 2
             Slider2(i).Visible = False
         Next i
            
Else
    
         Label(0).Visible = True
         Label(2).Visible = True
         Label(3).Visible = True
         Label(4).Visible = True
        Phase_diff.Visible = True
         Fdiff.Visible = True
         V_ref.Visible = True
         V_inc.Visible = True
         
          Frame1.Visible = False
          Frame2.Visible = False
          Frame3.Visible = False
          Frame4.Visible = False
  
        Command1.BackColor = &HFF&         ' RED
        Command1.Caption = "OFF LINE"
         For i = 0 To 2
             Slider2(i).Visible = True
         Next i
        
End If
      
If synchroscope Then
  Label1.Caption = "Synchroscope"
  Shape1.Visible = True
  Shape2.Visible = True
Else
  Label1.Caption = "Phasor Diagram"
  Shape1.Visible = False
  Shape2.Visible = False
End If

End Sub


