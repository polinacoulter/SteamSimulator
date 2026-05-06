VERSION 5.00
Begin VB.Form InstructorParameters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructor Parameters"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6735
   Begin VB.TextBox txtIRQ 
      Height          =   285
      Left            =   3300
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   4080
      Width           =   675
   End
   Begin VB.TextBox txtAddr 
      Height          =   285
      Left            =   3300
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   3720
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2700
      TabIndex        =   49
      ToolTipText     =   "Save These, Sound, Screen Size And Langugae Values"
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CheckBox chkBeep 
      Caption         =   "Beep For Attention"
      Height          =   315
      Left            =   4560
      TabIndex        =   48
      ToolTipText     =   "Sound A Beep To Call Attention To New Alarms"
      Top             =   3960
      Width           =   2115
   End
   Begin VB.Frame frmSliders 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   5220
      TabIndex        =   39
      Top             =   360
      Width           =   1635
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   8
         LargeChange     =   10
         Left            =   0
         Max             =   1000
         Min             =   100
         TabIndex        =   55
         Top             =   2880
         Value           =   100
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   0
         LargeChange     =   5
         Left            =   0
         Max             =   40
         TabIndex        =   47
         Top             =   0
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   1
         LargeChange     =   5
         Left            =   0
         Max             =   50
         TabIndex        =   46
         Top             =   360
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   45
         Top             =   720
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   44
         Top             =   1080
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   43
         Top             =   1440
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   42
         Top             =   1800
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   41
         Top             =   2160
         Width           =   1395
      End
      Begin VB.HScrollBar scrValue 
         Height          =   255
         Index           =   7
         LargeChange     =   10
         Left            =   0
         Max             =   100
         TabIndex        =   40
         Top             =   2520
         Width           =   1395
      End
   End
   Begin VB.Timer tmrCheckActive 
      Interval        =   200
      Left            =   360
      Top             =   4920
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   1080
      TabIndex        =   27
      ToolTipText     =   "Reset These Options (Except Amplicon's) To Their Default Values"
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   29
      ToolTipText     =   "Discard Any Changes Just Made"
      Top             =   4920
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4260
      TabIndex        =   28
      ToolTipText     =   "Accept These Values"
      Top             =   4920
      Width           =   1155
   End
   Begin VB.Label lblName 
      Caption         =   "FO Service Tank Capacity (Tonnes)"
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   59
      Top             =   3240
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0.3 T"
      Height          =   255
      Index           =   17
      Left            =   2880
      TabIndex        =   58
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0.1-1 T"
      Height          =   255
      Index           =   16
      Left            =   3600
      TabIndex        =   57
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   56
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "Changes To These Hardware Settings Only Take Effect When The Program Is Restarted"
      Height          =   435
      Left            =   0
      TabIndex        =   54
      Top             =   4440
      Width           =   6675
   End
   Begin VB.Label Label3 
      Caption         =   "Amplicon Board IRQ (e.g. 5)"
      Height          =   255
      Left            =   60
      TabIndex        =   51
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Amplicon Board Address (e.g. 300)"
      Height          =   255
      Left            =   60
      TabIndex        =   50
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-40"
      Height          =   255
      Index           =   15
      Left            =   3720
      TabIndex        =   38
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3660
      TabIndex        =   37
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-50"
      Height          =   255
      Index           =   14
      Left            =   3720
      TabIndex        =   36
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-100%"
      Height          =   255
      Index           =   13
      Left            =   3720
      TabIndex        =   35
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-100%"
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   34
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-100%"
      Height          =   255
      Index           =   11
      Left            =   3720
      TabIndex        =   33
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-100%"
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-100%"
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   31
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "0-100%"
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   30
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "60%"
      Height          =   255
      Index           =   7
      Left            =   2940
      TabIndex        =   26
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   25
      Top             =   2880
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "CPP Sump Level"
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   24
      Top             =   2880
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "80%"
      Height          =   255
      Index           =   6
      Left            =   2940
      TabIndex        =   23
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   22
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "HTCW Expansion Tank Level"
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   21
      Top             =   2520
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "90%"
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   20
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   19
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "ME LO Sump Level"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   18
      Top             =   2160
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "80%"
      Height          =   255
      Index           =   4
      Left            =   2940
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   16
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "LTCW Expansion Tank Level"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   15
      Top             =   1800
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "20%"
      Height          =   255
      Index           =   3
      Left            =   2940
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "Fuel Oil Service Tank Level"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   12
      Top             =   1440
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "20%"
      Height          =   255
      Index           =   2
      Left            =   2940
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "Fuel Oil Settling Tank Level"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   1080
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "25"
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "Ambient E. R. Temperature"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   2715
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4500
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label lblDefault 
      Alignment       =   2  'Center
      Caption         =   "15"
      Height          =   255
      Index           =   0
      Left            =   2940
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblName 
      Caption         =   "Sea Water Temperature"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   2715
   End
End
Attribute VB_Name = "InstructorParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'18/11/97 Dim OrigValues(0 To 7) As Integer
Dim OrigValues(0 To 8) As Integer
Dim OrigBeep As Boolean

Private Sub chkBeep_Click()
    BeepForAttention = chkBeep.Value
End Sub

Private Sub cmdCancel_Click()
    ' Reset the values to those that we started with on the Form Load
    Dim i As Integer
    
'18/11/97
'   For i = 0 To 7
    For i = 0 To 8
         IPVALUES(i + 1) = OrigValues(i)
    Next i
    
    BeepForAttention = OrigBeep
    Unload Me
End Sub

Private Sub cmdDefaults_Click()
    ' Reset the default values
    scrValue(0).Value = 15
    scrValue(1).Value = 25
    scrValue(2).Value = 20
    scrValue(3).Value = 20
    scrValue(4).Value = 80
    scrValue(5).Value = 90
    scrValue(6).Value = 80
    scrValue(7).Value = 60
'18/11/97
    scrValue(8).Value = 300

End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ' Open a default file (for overwrite).
    ' Save the values here and for the volumes (via sound mixer)
    
    ' Should I save the screen sizes???
    
    Dim MyFile As Integer
    Dim i As Integer
    
    MyFile = FreeFile
    Open App.Path & "\InstructorParams.ini" For Output As MyFile
    
    ' Write the values shown on this form
'18/11/97 added:
    Write #MyFile, CFO_service_tank
    
    Write #MyFile, SWT
    Write #MyFile, TA
    Write #MyFile, ZLFST
    Write #MyFile, ZLDT
    Write #MyFile, ZLLT
    Write #MyFile, ZLS
    Write #MyFile, ZLHT
    Write #MyFile, ZLSCP
    Write #MyFile, BeepForAttention
    Write #MyFile, useEnglish
    Write #MyFile, hardwareAddress
    Write #MyFile, hardwareIRQ
    
    ' Save the volumes
    For i = 0 To 5
        Write #MyFile, sndVolume(i)
    Next i
    
    ' Save the screen sizes
    For i = 1 To MaxScreens
        Write #MyFile, ScreenSettings(i).Height
        Write #MyFile, ScreenSettings(i).Width
    Next i
    
    ' Save the users choice of logs - the user still has to open a log file before these take effect
    ' Note the WRITE is used in preference to PRINT as PRINT does not work with Boolean types properly
    Write #MyFile, frmMainInterface.mnuLogAlarms.Checked
    Write #MyFile, frmMainInterface.mnuLogAnalog.Checked
    Write #MyFile, frmMainInterface.mnuLogDigital.Checked
    Write #MyFile, frmMainInterface.mnuLogStudentActions.Checked
    Close MyFile
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
' ===== 19/11/97 =========================
Dim ii As Integer

If useEnglish Then
   Open App.Path & "\EInstructor.txt" For Input As #1
Else
   Open App.Path & "\IInstructor.txt" For Input As #1
End If
For ii = 1 To 22
   Input #1, ice_strings(ii)
Next ii
Close #1
 
 InstructorParameters.Caption = ice_strings(1)
 Label1(0).Caption = ice_strings(2)
 Label1(1).Caption = ice_strings(3)
 Label1(3).Caption = ice_strings(4)
 Label1(2).Caption = ice_strings(5)
For ii = 1 To 9
    lblName(ii - 1).Caption = ice_strings(ii + 5)
Next ii
 Label2.Caption = ice_strings(15)
 Label3.Caption = ice_strings(16)
 chkBeep.Caption = ice_strings(17)
 Label4.Caption = ice_strings(18)
 cmdDefaults.Caption = ice_strings(19)
 cmdSave.Caption = ice_strings(20)
 cmdOK.Caption = ice_strings(21)
 cmdCancel.Caption = ice_strings(22)

Close #1
'============================================
    
    ' from original dos program
    IPVALUES(1) = SWT
    IPVALUES(2) = TA
    IPVALUES(3) = ZLFST * 100
    IPVALUES(4) = ZLDT * 100
    IPVALUES(5) = ZLLT * 100
    IPVALUES(6) = ZLS * 100
    IPVALUES(7) = ZLHT * 100
    IPVALUES(8) = ZLSCP * 100
    
    IPVALUES(9) = CFO_service_tank
    
    ' Add a "Degree C" caption to these two Default values
    lblDefault(15).Caption = lblDefault(15).Caption & Chr$(176) & "C"
    lblDefault(14).Caption = lblDefault(14).Caption & Chr$(176) & "C"
    lblDefault(0).Caption = lblDefault(0).Caption & Chr$(176) & "C"
    lblDefault(1).Caption = lblDefault(1).Caption & Chr$(176) & "C"
    
    ' Load in the current values
    For i = 0 To 7
       lblValue(i).Caption = IPVALUES(i + 1)
        scrValue(i).Value = IPVALUES(i + 1)
        OrigValues(i) = IPVALUES(i + 1)
    Next i
'18/11/97
       lblValue(8).Caption = IPVALUES(9) * 0.001
        scrValue(8).Value = IPVALUES(9)
        OrigValues(8) = IPVALUES(9)
        
    OrigBeep = BeepForAttention

    If BeepForAttention Then
        chkBeep.Value = 1
    Else
        chkBeep.Value = 0
    End If
    
    ' Set the hardware values
    txtAddr.Text = hardwareAddress
    txtIRQ.Text = hardwareIRQ

    tmrCheckActive_Timer
End Sub

Private Sub scrValue_Change(Index As Integer)

'18/11/97
If Index = 8 Then
    lblValue(Index).Caption = scrValue(Index).Value * 0.001
Else
    lblValue(Index).Caption = scrValue(Index).Value
End If

    IPVALUES(Index + 1) = scrValue(Index).Value
        
    If Index < 2 Then
        lblValue(Index).Caption = lblValue(Index).Caption & Chr$(176) & "C"
    ElseIf Index = 8 Then ' 18/11/97 FO Service Tank Capacity
        lblValue(Index).Caption = lblValue(Index).Caption & " T"
    Else
        lblValue(Index).Caption = lblValue(Index).Caption & "%"
    End If
    
End Sub

Private Sub scrValue_Scroll(Index As Integer)
    Call scrValue_Change(Index)
End Sub

Private Sub tmrCheckActive_Timer()
    ' If the model is running, then turn off the slider bars, and the Default/Cancel buttons.
    ' If not running, then we can change anything.
    Dim i As Integer
    
    If Model_Run Then
        cmdDefaults.Enabled = False
        cmdCancel.Enabled = False
        frmSliders.Enabled = False
        
'18/11/97
'      For i = 0 To 7
       For i = 0 To 8
            lblValue(i).Enabled = False
        Next i
        
        ' Copy the program's values, and show them
        IPVALUES(1) = SWT
        IPVALUES(2) = TA
        IPVALUES(3) = ZLFST * 100
        IPVALUES(4) = ZLDT * 100
        IPVALUES(5) = ZLLT * 100
        IPVALUES(6) = ZLS * 100
        IPVALUES(7) = ZLHT * 100
        IPVALUES(8) = ZLSCP * 100
        
'19/11/97
         IPVALUES(9) = CFO_service_tank

'18/11/97   For i = 0 To 7
            For i = 0 To 8
                scrValue(i).Value = IPVALUES(i + 1)
            Next i
        
    Else
        If cmdDefaults.Enabled = False Then cmdDefaults.Enabled = True
        If cmdCancel.Enabled = False Then cmdCancel.Enabled = True
        If frmSliders.Enabled = False Then frmSliders.Enabled = True
        If lblValue(0).Enabled = False Then
 '18/11/97  For i = 0 To 7
            For i = 0 To 8
                lblValue(i).Enabled = True
            Next i
        End If
        ' Copy the updated values back to the main program
        SWT = IPVALUES(1)
        TA = IPVALUES(2)
        ZLFST = IPVALUES(3) / 100
        ZLDT = IPVALUES(4) / 100
        ZLLT = IPVALUES(5) / 100
        ZLS = IPVALUES(6) / 100
        ZLHT = IPVALUES(7) / 100
        ZLSCP = IPVALUES(8) / 100
'18/11/97
        CFO_service_tank = IPVALUES(9)
        
    End If
End Sub

Private Sub txtAddr_Change()
    hardwareAddress = Val(txtAddr.Text)
End Sub

Private Sub txtIRQ_Change()
    hardwareIRQ = Val(txtIRQ.Text)
End Sub
