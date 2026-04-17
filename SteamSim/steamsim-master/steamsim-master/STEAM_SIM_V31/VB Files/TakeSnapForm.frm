VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Begin VB.Form SnapForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snapshot"
   ClientHeight    =   9060
   ClientLeft      =   300
   ClientTop       =   1635
   ClientWidth     =   16035
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TakeSnapForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9060
   ScaleWidth      =   16035
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SnapTime 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   14460
      TabIndex        =   4
      Text            =   "Time"
      Top             =   2535
      Width           =   1095
   End
   Begin VB.TextBox SnapDate 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13410
      TabIndex        =   3
      Text            =   "Date"
      Top             =   2535
      Width           =   1035
   End
   Begin VB.TextBox Description 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8190
      TabIndex        =   2
      Text            =   "Description"
      Top             =   2535
      Width           =   5220
   End
   Begin VB.TextBox Course 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5985
      TabIndex        =   1
      Text            =   "Course"
      Top             =   2535
      Width           =   2205
   End
   Begin VB.TextBox Class 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Text            =   "Class"
      Top             =   2535
      Width           =   2370
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15720
      Begin VB.CommandButton TakeSnapOK 
         Caption         =   "Take Snapshot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   705
         Width           =   1335
      End
      Begin VB.TextBox SnapshotsAvailable 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   14
         Text            =   "Snapshots Available"
         Top             =   1965
         Width           =   1815
      End
      Begin VB.TextBox Number 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   255
         TabIndex        =   13
         Text            =   "Number"
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox Instructor 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1230
         TabIndex        =   12
         Text            =   "Instructor"
         Top             =   2460
         Width           =   2280
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   5520
         Left            =   240
         TabIndex        =   11
         Top             =   2865
         Width           =   15300
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   270
         LinkTimeout     =   30
         MaxLength       =   50
         TabIndex        =   10
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   285
         TabIndex        =   9
         Text            =   "Description of Selected Snapshot"
         Top             =   345
         Width           =   3015
      End
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6255
         TabIndex        =   8
         Top             =   1170
         Width           =   1335
      End
      Begin VB.TextBox SnapNumSelected 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5340
         TabIndex        =   7
         Top             =   735
         Width           =   375
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5250
         TabIndex        =   6
         Text            =   "Number"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label DEMO_label 
         Alignment       =   2  'Center
         Caption         =   "WARNING: DEMONSTRATION VERSION.  The MODEL will be permanently halted after a snapshot is re-loaded."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   2040
         TabIndex        =   16
         Top             =   1620
         Width           =   7215
      End
   End
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   2415
      Left            =   60
      Top             =   6960
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   4260
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   5520
      Y1              =   2640
      Y2              =   3120
   End
End
Attribute VB_Name = "SnapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Fill_Default_Snap_Info_File()

Dim Filename As String
Dim pathname As String
Dim wholename As String
Dim InfoString As String * 200

Dim record_num As Integer
Dim msg As String

Dim Numberstring As String
Dim InstructorString As String
Dim ClassString As String
Dim CourseString As String
Dim Datestring As String
Dim DescriptionString As String
Dim i As Integer
Dim Fnum As Integer






'On Error GoTo ErrorHandler

' 12/7/99
'pathname = "c:\HAVEN\Snapshots\"
pathname = "c:\Steam_Sim\Snapshots\"
Fnum = FreeFile(0)



If medium_speed Then
    If Not split_mode Then
        Filename = "Med_Comb_Snap_Info.dat"
    Else
        If ElecExec Then
            Filename = "Med_Elec_Snap_Info.dat"
        Else
            Filename = "Med_Diesel_Snap_Info.dat"
        End If
    End If
Else
    If Not split_mode Then
        Filename = "Slow_Comb_Snap_Info.dat"
    Else
        If ElecExec Then
            Filename = "Slow_Elec_Snap_Info.dat"
        Else
            Filename = "Slow_Diesel_Snap_Info.dat"
        End If
    End If
End If
wholename = pathname & Filename

'Set up some dummy info
InstructorString = "Instructor"
ClassString = "Class"
CourseString = "Course"
DescriptionString = "Description"
Datestring = Date
TimeString = Time

'put dummy info in a string of correct length

InfoString = String(Rec_Length, " ")
Mid(InfoString, 1, Len(Numberstring)) = Numberstring
Mid(InfoString, 10, Len(InstructorString)) = InstructorString
Mid(InfoString, 32, Len(ClassString)) = ClassString
Mid(InfoString, 54, Len(CourseString)) = CourseString
Mid(InfoString, 75, Len(DescriptionString)) = DescriptionString
Mid(InfoString, 125, Len(Datestring)) = Datestring
Mid(InfoString, 135, Len(TimeString)) = TimeString

'Debug.Print "in fill default, about to put in data ... "
'write dummy strings to file,and increment number of each record
'Open wholename For Random Access Write As #11 Len = Rec_Length
Open wholename For Random Access Write As #Fnum Len = Rec_Length
        '11/12/00 For record_num = 1 To 20
        For record_num = 1 To 25
            Mid(InfoString, 2, Len(str(record_num))) = str(record_num)
            'Put #11, record_num, InfoString
            Put #Fnum, record_num, InfoString
            'debug.print "Put record_num(" & record_num & ")"
        Next
        'debug.print "Closing snapshot info File"
'Close #11
Close #Fnum
       
        
Exit Sub
errorhandler:
DoEvents
'Close #11
Close #Fnum
'debug.print "Error Number = " & Err.Number
'debug.print "Error Description = " & Err.Description

msg = MsgBox("Error Reading:  " & wholename & Err.Description, 48)
DoEvents

Exit Sub

End Sub



Sub FillSnapList()
Dim Filename As String
Dim pathname As String
Dim wholename As String
Dim InfoString As String * 200

Dim record_num As Integer
Dim msg As String

Dim Numberstring As String
Dim InstructorString As String
Dim ClassString As String
Dim CourseString As String
Dim SnapDescription  As String * 30
Dim Datestring As String
Dim DescriptionString As String
Dim i As Integer
Dim Response As Integer

Dim Fnum As Integer





On Error GoTo errorhandler

' 12/7/99
'pathname = "c:\HAVEN\Snapshots\" ' temporary, use app.path\snapshots after development
pathname = "c:\Steam_sim\Snapshots\"
Fnum = FreeFile(0)

If medium_speed Then
    If Not split_mode Then
        Filename = "Med_Comb_Snap_Info.dat"
    Else
        If ElecExec Then
            Filename = "Med_Elec_Snap_Info.dat"
        Else
            Filename = "Med_Diesel_Snap_Info.dat"
        End If
    End If
Else
    If Not split_mode Then
        Filename = "Slow_Comb_Snap_Info.dat"
    Else
        If ElecExec Then
            Filename = "Slow_Elec_Snap_Info.dat"
        Else
            Filename = "Slow_Diesel_Snap_Info.dat"
        End If
    End If
End If
wholename = pathname & Filename

'Open wholename For Random Access Read As #10 Len = Rec_Length
Open wholename For Random Access Read As #Fnum Len = Rec_Length

        List1.clear
        '11/12/00 For record_num = 1 To 20
        For record_num = 1 To 25
            'Get #10, record_num, InfoString
            Get #Fnum, record_num, InfoString
            List1.AddItem InfoString
        Next
        'Close #10
        Close #Fnum

        If Val(InfoString) = 0 Then ' there are no details in info file
            Response = MsgBox("Could not find snapshot information in " & wholename & Chr(13) & Chr(10) & "Do you wish create default information ?", vbYesNo, "Snapshot Information")
            If Response = vbYes Then
               Call Fill_Default_Snap_Info_File
            End If
       End If

       
        
Exit Sub
errorhandler:
DoEvents
'Close #10
Close #Fnum
'debug.print "Error Number = " & Err.Number
'debug.print "Error Description = " & Err.Description

msg = MsgBox("Error :  " & Err.Number & "  " & Err.Description, 48)
DoEvents

Exit Sub
End Sub


Sub SaveSnapList()
Dim Filename As String
Dim pathname As String
Dim wholename As String
Dim InfoString As String * 200

Dim record_num As Integer
Dim msg As String

Dim Numberstring As String
Dim InstructorString As String
Dim ClassString As String
Dim CourseString As String
Dim SnapDescription  As String * 30
Dim Datestring As String
Dim TimeString As String
Dim DescriptionString As String
Dim i As Integer

Dim Fnum As Integer





On Error GoTo errorhandler

' 12/7/99
'pathname = "c:\HAVEN\Snapshots\"
pathname = "c:\Steam_Sim\Snapshots\"
Fnum = FreeFile(0)

If medium_speed Then
    If Not split_mode Then
        Filename = "Med_Comb_Snap_Info.dat"
    Else
        If ElecExec Then
            Filename = "Med_Elec_Snap_Info.dat"
        Else
            Filename = "Med_Diesel_Snap_Info.dat"
        End If
    End If
Else
    If Not split_mode Then
        Filename = "Slow_Comb_Snap_Info.dat"
    Else
        If ElecExec Then
            Filename = "Slow_Elec_Snap_Info.dat"
        Else
            Filename = "Slow_Diesel_Snap_Info.dat"
        End If
    End If
End If
wholename = pathname & Filename

'change information for snapshot being overwritten
Numberstring = Trim(SnapNumSelected.text)
InstructorString = Trim(InstructorName)
ClassString = Trim(ClassName)
CourseString = Trim(CourseName)
DescriptionString = Trim(Text1.text)
Datestring = Date
TimeString = Time
record_num = Val(Numberstring)

'11/12/00
InstructorString = UCase(InstructorString)

If InstructorString = "LLANELLI" Then

    ' === 30/11/01 ===
   If record_num > 20 And record_num <= 25 Then ' Snapshots 21 to 25 are "RESERVED"
        InstructorString = "RESERVED" 'Snapshots 21 to 25 saved under the Llanelli Instructor name are to reserved
   Else
        InstructorString = ""    ' Llanelli Snapshots 1 to 20 have NO Instructor Name
   End If
   
   I_AM_LLANELLI = True
   
Else
   If record_num > 20 And record_num <= 25 Then 'Don't let instructors other than Llanelli change snaps 21 to 25
        Exit Sub
   End If
   InstructorString = Trim(InstructorName) 'could have left it uppercase but ..
End If


'Clear a string of correct length
For i = 1 To Rec_Length
            Mid(InfoString, i, 1) = Chr(32)
Next

'setup a record string to save in file
Mid(InfoString, 1, Len(Numberstring)) = Numberstring
Mid(InfoString, 10, Len(InstructorString)) = InstructorString
Mid(InfoString, 32, Len(ClassString)) = ClassString
Mid(InfoString, 54, Len(CourseString)) = CourseString
Mid(InfoString, 75, Len(DescriptionString)) = DescriptionString
Mid(InfoString, 125, Len(Datestring)) = Datestring
Mid(InfoString, 135, Len(TimeString)) = TimeString

'now open the file for writing and save

'Open wholename For Random Access Write As #10 Len = Rec_Length
Open wholename For Random Access Write As #Fnum Len = Rec_Length
    List1.List(record_num) = InfoString
    Put #Fnum, record_num, InfoString
'Close #10
Close #Fnum



        
        
Exit Sub
errorhandler:
    DoEvents
    'Close #10
    Close #Fnum
    Debug.Print "Error Number = " & Err.Number
    Debug.Print "Error Description = " & Err.Description

    msg = MsgBox("Error :  " & Err.Number & "  " & Err.Description, 48)

Exit Sub
End Sub

Private Sub Cancel_Click()

Unload Me

End Sub

Private Sub Exit_Snapshot_Click()
Unload Me

End Sub


Private Sub Form_Load()
Dim Filename As String
Dim pathname As String
Dim wholename As String
Dim InfoString As String * 32
Dim record_num As Integer
Dim msg As String

'08/09/00
'SnapForm.Width = 15885


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)



' 21/7/00  To ensure frm_faults updates correcly...
Unload frm_faults


'08/09/00 align frame centrally
If SnapForm.Width > List1.Width Then
  'Frame1.Left = SnapForm.Left + ((SnapForm.Width - Frame1.Width) / 2)
End If

' 15/10/98
If DEMO_flag Then
  DEMO_label.Visible = True
Else
  DEMO_label.Visible = False
End If

'SnapshotsAvailable.Top = Number.Top + 300           'Align Text boxes closely
'List1.Top = Number.Top - 300
TakeSnapOK.Enabled = False
'debug.print "calling FillSnapList"

Call FillSnapList
If TakeSnapFlag Then
    SnapForm.Caption = "Save Snapshot"
    SnapForm.TakeSnapOK.Caption = "Save Snapshot"
Else
    SnapForm.Caption = "Load Snapshot"
    SnapForm.TakeSnapOK.Caption = "Load Snapshot"
End If
'debug.print " about to show SnapForm"

SnapForm.Show
        
End Sub

Private Sub List1_Click()

Dim SnapNumStr As String
SnapNumStr = Mid(List1.List(List1.ListIndex), 1, 4)

' 2/8/00  Needs MORE characters...
'Text1.text = RTrim(Mid(List1.List(List1.ListIndex), 75, 25))
Text1.text = RTrim(Mid(List1.List(List1.ListIndex), 75, 55))

'07/09/00 get all the text
Text1.text = RTrim(Mid(List1.List(List1.ListIndex), 75, Text1.MaxLength))

SnapNumSelected.text = str(Val(SnapNumStr))
Debug.Print "SnapNumStr = " & SnapNumStr
Debug.Print "SnapNumSelected.Text = " & SnapNumSelected.text

TakeSnapOK.Enabled = True


'07/09/00 select text in description text box ready for edit ...

If TakeSnapFlag Then 'select text in description text box ready for edit ...

    'If List1.ListIndex > 0 Then
        'List1.Selected(List1.ListIndex) = False
    'End If
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.text)
End If


End Sub


Private Sub SnapOptions_Click()

End Sub

Private Sub TakeSnapOK_Click()
Dim IncrementVal As Integer
Dim i As Integer
Dim SnapCommand As String



'debug.print " TakeSnapFlag = " & TakeSnapFlag
'debug.print " medium_speed = " & medium_speed
'debug.print " Val (SnapNumSelected.Text) = " & Val(SnapNumSelected.Text)


' 11/12/98
Me.MousePointer = ccHourglass

If TakeSnapFlag Then                                'Snapshot is being taken
    '11/12/00 If Val(SnapNumSelected.text) > 0 And Val(SnapNumSelected.text) <= 20 Then
    If Val(SnapNumSelected.text) > 0 And Val(SnapNumSelected.text) <= 25 Then
        SnapNum = Val(SnapNumSelected.text)
        SnapType = 1 'default
        'Save new description in snap info file
        Call SaveSnapList
        Call FillSnapList               'show updated list
        
        
    'debug.print " medium_speed = " & medium_speed
    '11/12/00 protect snaps 21 to 25
    If I_AM_LLANELLI Or SnapNum <= 20 Then
        'work out which command to send - depends on configuration
        If medium_speed Then
            If Not split_mode Then
               SnapType = 1
            Else
                If ElecExec Then
                   SnapType = 2
                Else
                   SnapType = 3
                End If
            End If
            
            'debug.print "about to call fortran snapshot"
            
            Call M_SNAP(SnapNum, SnapType)
        Else
            If Not split_mode Then
               SnapType = 1
            Else
                If ElecExec Then
                   SnapType = 2
                Else
                   SnapType = 3
                End If
            End If
            
            'debug.print "about to call fortran snapshot"

            Call LSS_SNAP(SnapNum, SnapType)

        End If
   
        'Send instruction to host
        'If Not test_flag Then
        '        MessageString = SnapCommand & Str$(SnapNumSelected.Text) & "#"
        '        frm_exec.Host_sock_Rx.SendLen = Len(MessageString)
        '        frm_exec.Host_sock_Rx.SendData = MessageString
        'End If
        
        
        '13/12/00 The sleep command used here, hangs sound. No need for the progress bar anyway - just pretty
        'Load SnapProgressForm
        '11/12/00 SnapProgressForm.Show
        'SnapProgressForm.Text1.text = "Saving Snapshot Number " & SnapNumSelected.text
        'SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Min
        'SnapProgressForm.ProgressBar1.Visible = True
        'SnapProgressForm.Text1.Visible = True
        'SnapProgressForm.Text2.Visible = True
        'SnapProgressForm.Text3.Visible = False
        'SnapProgressForm.Show
        'IncrementVal = (SnapProgressForm.ProgressBar1.max - SnapProgressForm.ProgressBar1.Min) / 20
        'For i = 1 To 20
        '    SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Value + IncrementVal
        '    Select Case i
        '        Case 1
        '           SnapProgressForm.Text2.text = "Opening File"
        '        Case 5
        '            SnapProgressForm.Text2.text = "Saving Data"
        '        Case 16
        '            SnapProgressForm.Text2.text = "Closing File"
        '    End Select
        '    DoEvents
        '    Call Sleep(50)
        'Next
        ' Unload SnapProgressForm
        
         Call Logging("INSTRUCTOR SAVED SNAPSHOT # " & SnapNum & " " & SnapNumSelected.text, LogSnaps)
      
      
      Else 'Pdisplay message explaining to non-llanelli instructor that he cannot overwrite this snap
         Load SnapProgressForm
         SnapProgressForm.Text1.Visible = False
         SnapProgressForm.Text2.Visible = False
         SnapProgressForm.Text3.Visible = True
         '13/12/00
         SnapProgressForm.CmdOK.Visible = True
         SnapProgressForm.Show
         SnapProgressForm.Text3.text = "Cannot Overwrite this RESERVED Snapshot"
         SnapProgressForm.ProgressBar1.Visible = False
         '13/12/00For i = 1 To 20
         '13/12/00   Call Sleep(50)
         '13/12/00   DoEvents
         '13/12/00 Next
         '13/12/00 unloading done on progress form OK click
         
         '13/12/00 Unload SnapProgressForm
      End If
         
    End If
    

    'End If
Else
        Debug.Print "reloading snapshot,"

    '11/12/00 If Val(SnapNumSelected.text) > 0 And Val(SnapNumSelected.text) <= 20 Then
    If Val(SnapNumSelected.text) > 0 And Val(SnapNumSelected.text) <= 25 Then
            SnapNum = Val(SnapNumSelected.text)
                Debug.Print "SnapNum = " & SnapNum


            ' 12/10/09 added...TESTED OK....
            frm_Host.StatusBar1.Panels.Item(2).text = "Snapshot " & Format(SnapNum, "#0") & " Start"


        'work out which command to send - depends on configuration
         If medium_speed Then
            If Not split_mode Then
               SnapType = 1
            Else
                If ElecExec Then
                   SnapType = 2
                Else
                   SnapType = 3
                End If
            End If
            Call M_RELSNAP(SnapNum, SnapType)
        Else
            If Not split_mode Then
               SnapType = 1
            Else
                If ElecExec Then
                   SnapType = 2
                Else
                   SnapType = 3
                End If
            End If
            
            '29/11/99 halt before loading
            
            d_run = False   ' HALT  the     Diesel  Simulator
            e_run = False   ' HALT  the Electrical  Simulator
            RunIO = False
            
            Call LSS_RELSNAP(SnapNum, SnapType)
            RunIO = True '6/1/00 Run IO here to prepare O/P in advance of running model
            
            '08/09/00
            ResetAllAlarms = True

        End If
        If Not test_flag Then
               ' MessageString = SnapCommand & Str$(SnapNumSelected.Text) & "#"
                'frm_exec.Host_sock_Rx.SendLen = Len(MessageString)
                'frm_exec.Host_sock_Rx.SendData = MessageString
        End If
        Load SnapProgressForm
        SnapProgressForm.Show
        SnapProgressForm.Text1.text = "Reloading Snapshot Number " & SnapNumSelected.text
        SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Min
        SnapProgressForm.ProgressBar1.Visible = True
        SnapProgressForm.Text1.Visible = True
        SnapProgressForm.Text2.Visible = True
        SnapProgressForm.Text3.Visible = False

        IncrementVal = (SnapProgressForm.ProgressBar1.max - SnapProgressForm.ProgressBar1.Min) / 20
        For i = 1 To 20
            SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Value + IncrementVal
            Select Case i
                Case 1
                    SnapProgressForm.Text2.text = "Opening File"
                Case 5
                    SnapProgressForm.Text2.text = "Loading Data"
                Case 16
                    SnapProgressForm.Text2.text = "Closing File"
            End Select
            DoEvents
            Call Sleep(50)
        Next
        
        
        ' 9/8/00
        Force_Instrp_update = True
        
      
        Unload SnapProgressForm
        
        
            ' === 23/08/02  (Read from the file c:\Steam_Sim\Text\Constants.txt ) ===
            Call SET_Constants

                
        '25/10/99  check if any DIs or AI's unmatched (in on-line only)
        '14/01/00 stsrt immedidtely if PTT
         If PTT Then
         
            d_run = True   ' start  the     Diesel  Simulator
            e_run = True   ' start  the Electrical  Simulator
            
            
         Else
        '14/01/00 If Not PTT Then
            Call SnapCompareForm.DI_Compare
            Call SnapCompareForm.AI_Compare
            
            If DIDifferences Or AIDifferences Then
                SnapCompareForm.Show
            Else
                '29/11/99 start after loading
                d_run = True   ' start  the     Diesel  Simulator
                e_run = True   ' start  the Electrical  Simulator
                RunIO = True
            End If
        End If
        
        

        
        Call Logging("INSTRUCTOR LOADED SNAPSHOT # " & SnapNum & " " & SnapNumSelected.text, LogSnaps)

        HeaderPrinted = False           'Ensure new log header is printed next time
    
    
    End If
    
    
        '  21/7/98 RWJ  update all sliders:
        For i = 1 To 48
           update_VBAI(i) = True
        Next i
    
    
End If

' 15/10/98
If DEMO_flag Then
    DEMO_halt_model = True
End If



' 11/12/98
Me.MousePointer = ccDefault

Unload Me

End Sub


