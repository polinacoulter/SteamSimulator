VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form Steam_Calculator 
   Caption         =   "Steam Calculator"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   10170
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6615
      Left            =   240
      Top             =   180
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   11668
      Begin VB.CommandButton Command9 
         Caption         =   "Calculate"
         Height          =   675
         Left            =   3720
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   14
         Text            =   "40.0        Bar"
         Top             =   720
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Text            =   "3094.0    kJ/kg"
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   5160
         TabIndex        =   12
         Text            =   "1.000"
         Top             =   1140
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   11
         Text            =   "                  degC"
         Top             =   720
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1140
         TabIndex        =   10
         Text            =   "40.0        Bar"
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1140
         TabIndex        =   9
         Text            =   "350.0       degC"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1140
         TabIndex        =   8
         Text            =   "1.000"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate"
         Height          =   675
         Left            =   3720
         TabIndex        =   7
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   5160
         TabIndex        =   6
         Text            =   "    kJ/kg"
         Top             =   2400
         Width           =   1995
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calculate"
         Height          =   675
         Left            =   3720
         TabIndex        =   5
         Top             =   4500
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   4
         Text            =   "       250.3     degC"
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   5220
         TabIndex        =   3
         Text            =   "     Bar"
         Top             =   4680
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Calculate"
         Height          =   675
         Left            =   3720
         TabIndex        =   2
         Top             =   5340
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   1200
         TabIndex        =   1
         Text            =   "40.0        Bar"
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   5220
         TabIndex        =   0
         Text            =   "           degC"
         Top             =   5520
         Width           =   1995
      End
      Begin VB.Line Line8 
         BorderWidth     =   5
         X1              =   60
         X2              =   60
         Y1              =   4080
         Y2              =   6240
      End
      Begin VB.Line Line5 
         BorderWidth     =   5
         X1              =   60
         X2              =   60
         Y1              =   3840
         Y2              =   120
      End
      Begin VB.Line Line7 
         BorderWidth     =   5
         X1              =   60
         X2              =   8640
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Pressure"
         Height          =   255
         Left            =   300
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Parameter"
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Result"
         Height          =   255
         Left            =   5580
         TabIndex        =   32
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Enthalpy"
         Height          =   255
         Left            =   300
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Dryness"
         Height          =   255
         Left            =   7200
         TabIndex        =   30
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Temperature"
         Height          =   255
         Left            =   7200
         TabIndex        =   29
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Pressure"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Temperature *"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Dryness"
         Height          =   255
         Left            =   -60
         TabIndex        =   26
         Top             =   2820
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Enthalpy"
         Height          =   255
         Left            =   7260
         TabIndex        =   25
         Top             =   2460
         Width           =   735
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   60
         X2              =   8640
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   8640
         X2              =   8640
         Y1              =   120
         Y2              =   3840
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Temperature"
         Height          =   255
         Left            =   60
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Pressure"
         Height          =   255
         Left            =   7320
         TabIndex        =   23
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Pressure"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Temperature"
         Height          =   255
         Left            =   7320
         TabIndex        =   21
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Saturation Pressure/Temperature"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   4200
         Width           =   2595
      End
      Begin VB.Line Line6 
         BorderWidth     =   5
         X1              =   60
         X2              =   8640
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderWidth     =   5
         X1              =   8640
         X2              =   8640
         Y1              =   4080
         Y2              =   6240
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Parameter"
         Height          =   255
         Left            =   1500
         TabIndex        =   19
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Result"
         Height          =   255
         Left            =   5820
         TabIndex        =   18
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   60
         X2              =   8640
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         X1              =   60
         X2              =   8640
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label Label18 
         Caption         =   "Temperature *  Enter the EXACT Saturation Temperature and you will be able to enter a Dryness value"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3300
         Width           =   7755
      End
      Begin VB.Label Label19 
         Caption         =   "Temperature *  If you enter ANY other value then the Dryness factor will be calculated for you"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3540
         Width           =   7755
      End
   End
End
Attribute VB_Name = "Steam_Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim P1_steam As Single
Dim T1_steam As Single
Dim H1_steam As Single

Dim q1_dryness As Single
Dim T1_sat As Single
Dim hf_1 As Single
Dim hfg_1 As Single
Dim hg_1 As Single


P1_steam = Val(Text1(4).text)
T1_steam = Val(Text1(5).text)

q1_dryness = Val(Text1(6).text)


Call get_Tsat(P1_steam, T1_sat)

' === 5/11/99 =======================

If (T1_steam > T1_sat) Then

    Text1(6).text = " 1.000  [Super-heated Steam]"
       q1_dryness = Val(Text1(6).text)
    Text1(6).Enabled = False
    
ElseIf (T1_steam < T1_sat) Then

    Text1(6).text = " 0.000  [Water]"
       q1_dryness = Val(Text1(6).text)
    Text1(6).Enabled = False
    Text1(7).text = " "
    Exit Sub
Else

    Text1(6).Enabled = True
        
    If (q1_dryness > 0.99999) Then
         q1_dryness = 1#
    ElseIf (q1_dryness < 0.00001) Then
         q1_dryness = 0#
    End If
    
    If (q1_dryness > 0.99999) Then
      Text1(6).text = Format(q1_dryness, "#0.000") & " [Super-heated Steam]"
    ElseIf (q1_dryness < 0.00001) Then
      Text1(6).text = Format(q1_dryness, "#0.000") & " [Water]"
    Else
      Text1(6).text = Format(q1_dryness, "#0.000") & " [Wet Steam]"
    End If
    
    
End If

' === 5/11/99 end of edit ============


Call get_hf_hfg(P1_steam, hf_1, hfg_1)


hg_1 = hf_1 + hfg_1


Call get_Hsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_Steam)

Text1(7).text = Format(H1_steam, "###0.0") & " kJ/kg"

End Sub

Private Sub Command2_Click()

Tsat = Val(Text1(8).text)
Call get_Psat(Psat, Tsat)

' 5/11/99
'Text1(9).text = Format(Psat, "###0.0") & " Bar"
Text1(9).text = Format(Psat, "###0.00") & " Bar"

End Sub

Private Sub Command3_Click()

Psat = Val(Text1(10).text)
Call get_Tsat(Psat, Tsat)

' 5/11/99
'Text1(11).text = Format(Tsat, "###0.0") & " " & Chr(176) & "C"
Text1(11).text = Format(Tsat, "###0.000 000 000") & " " & Chr(176) & "C"

End Sub

Private Sub Command9_Click()
Dim P1_steam As Single
Dim T1_steam As Single
Dim H1_steam As Single

Dim q1_dryness As Single
Dim T1_sat As Single
Dim hf_1 As Single
Dim hfg_1 As Single
Dim hg_1 As Single


' 20/4/99
P1_steam = Val(Text1(0).text)
H1_steam = Val(Text1(1).text)

Call get_Tsat(P1_steam, T1_sat)

Call get_hf_hfg(P1_steam, hf_1, hfg_1)

hg_1 = hf_1 + hfg_1

' ==== 12/11/99 ====
If H1_steam < hf_1 Then
   Text1(1).text = "Value too low"
   Text1(2).text = " "
   Text1(3).text = " "
   Exit Sub
End If
' ==================


Call get_Tsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_Steam)

Text1(2).text = Format(T1_steam, "###0.0") & " " & Chr(176) & "C "
Text1(3).text = Format(q1_dryness, "###0.000")


End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idSteam_Calc)


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idSteam_Calc)
    HavenStretch1.Visible = True
End If


End Sub

