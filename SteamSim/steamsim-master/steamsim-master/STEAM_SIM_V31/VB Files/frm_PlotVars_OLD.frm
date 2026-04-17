VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form frm_PlotVars_OLD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plot variables "
   ClientHeight    =   9810
   ClientLeft      =   1920
   ClientTop       =   2205
   ClientWidth     =   11475
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9810
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9015
      Left            =   0
      Top             =   60
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   15901
      Begin VB.Frame FaultEntry 
         Height          =   8955
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   10275
         Begin VB.ListBox VarTypes 
            BackColor       =   &H00C0C0C0&
            Columns         =   1
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   4380
            Left            =   120
            TabIndex        =   11
            Top             =   4320
            Width           =   3675
         End
         Begin VB.ListBox Description 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4380
            Left            =   3855
            TabIndex        =   10
            Top             =   4350
            Width           =   6315
         End
         Begin VB.TextBox SelectedVar 
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   5655
         End
         Begin VB.CommandButton EnterCommand 
            Caption         =   "Enter"
            Height          =   375
            Left            =   5880
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
         Begin VB.ListBox ActiveVarList 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1980
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   7815
         End
         Begin VB.CommandButton ClearAll 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   9135
            TabIndex        =   6
            ToolTipText     =   "Clear All Actions that are currently entered"
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton OK 
            Caption         =   "OK"
            Height          =   375
            Left            =   9120
            TabIndex        =   5
            Top             =   3840
            Width           =   975
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   8040
            Top             =   3840
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   5040
            TabIndex        =   4
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   9225
            TabIndex        =   3
            Text            =   "none"
            Top             =   3240
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frm_PlotVars_OLD.frx":0000
            Left            =   8910
            List            =   "frm_PlotVars_OLD.frx":0002
            TabIndex        =   2
            Top             =   2565
            Width           =   1140
         End
         Begin VB.CommandButton ClearCommand 
            Caption         =   "Clear"
            Height          =   375
            Left            =   9075
            TabIndex        =   1
            ToolTipText     =   "Clear the Selected Action Only"
            Top             =   405
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Variable Types"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
            Height          =   255
            Left            =   3840
            TabIndex        =   17
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Selected Variable"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Variables to Plot"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Line Line1 
            BorderWidth     =   5
            Index           =   0
            X1              =   8070
            X2              =   8550
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   5
            Index           =   1
            X1              =   8070
            X2              =   8550
            Y1              =   1890
            Y2              =   1890
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   5
            Index           =   2
            X1              =   8070
            X2              =   8550
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFF00&
            BorderWidth     =   5
            Index           =   3
            X1              =   8070
            X2              =   8550
            Y1              =   2370
            Y2              =   2370
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   5
            Index           =   4
            X1              =   8070
            X2              =   8550
            Y1              =   2610
            Y2              =   2610
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF00FF&
            BorderWidth     =   5
            Index           =   5
            X1              =   8070
            X2              =   8550
            Y1              =   2850
            Y2              =   2850
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   5
            Index           =   6
            X1              =   8070
            X2              =   8550
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderWidth     =   5
            Index           =   7
            X1              =   8070
            X2              =   8550
            Y1              =   3330
            Y2              =   3330
         End
         Begin VB.Label Label5 
            Caption         =   "Thick Line"
            Height          =   255
            Left            =   5400
            TabIndex        =   14
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Grid Style"
            Height          =   255
            Left            =   8835
            TabIndex        =   13
            Top             =   2955
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Sample Interval"
            Height          =   255
            Left            =   8880
            TabIndex        =   12
            Top             =   2310
            Width           =   1260
         End
      End
   End
End
Attribute VB_Name = "frm_PlotVars_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Combo1_Click()
Dim resp As Integer

    If Combo1.ListIndex >= 0 And Combo1.ListIndex <= 3 Then
        GridStyle = Combo1.ListIndex
        Debug.Print "setting GridStyle in click = " & GridStyle
        resp = MsgBox("Changes to the Grid will take effect " & Chr(10) & _
                      "next time you select plotting", vbInformation & vbOKOnly, "Grid Line Style")
                      
        'UnloadPlot = True
        'Sleep (1500)
        'frm_Plot1.Show
        
    End If

End Sub


Sub Fill_Active_Var_List(form_name As Form)
Dim i As Integer
Dim LenText As Integer
Dim VarString As String
Dim Padding As Integer


    If VarChanges Then ' only update list if there has been a change to plottrer configuration
        ActiveVarList.Clear
        NumberOfVars = 0
        For i = 1 To 8
              If PlotVar(i) <> 0 Then ' there is an entry being plotted
                NumberOfVars = NumberOfVars + 1
                LenText = Len(LegendText(i))
                VarString = Trim(Mid(LegendText(i), 5, LenText - 4)) 'drop
                'frm_Plot1.Label2(i - 1) = VarString
                'LegendText(i) = Trim(Mid(VarString, 4, LenText - 3))

                'Padding = 45 - LenText
                Padding = 55 - LenText
                'VarString = VarString + Space(Padding) & Var_time(i)
                If Padding < 1 Then
                  Padding = 1
                End If
                
                form_name.ActiveVarList.AddItem LegendText(i) & Space(Padding) & Var_time(i)

              Else
                VarString = "Not Allocated"
                form_name.ActiveVarList.AddItem LegendText(i)

              End If
              plot_channels(i) = PlotVar(i)
              'frm_Plot1.Label2(i - 1) = VarString 'update legend on graph itself, text only no number or time
              GraphFrm.Label2(i - 1) = VarString 'update legend on graph itself, text only no number or time
              Sets(i) = plot_channels(i)
        Next
        VarChanges = False
 ' 26/3/98   Call setplotchannels(plot_channels(1))
        Call setplotchannels

    End If
  
 
End Sub
Sub FillDescription(FirstAction As Integer, LastAction As Integer)
Dim i As Integer
Dim ActionOK As Integer
Dim ThisVarNumber As Integer
Dim ActionAlreadyEntered As Integer
Dim StrLen As Integer




For i = FirstAction To LastAction                       'check for range of variable numbers defined by group

     ThisVarNumber = Val(Mid$(VarText(i), 1, 4))            'get Variable number
     If ThisVarNumber > 0 Then
          StrLen = Len(VarText(i))
          Description.AddItem Format(str(ThisVarNumber), "000") & Space(1) & Mid(VarText(i), 5, StrLen - 4)
     End If
Next i

End Sub



Sub FillXref()
Dim i As Integer


'Makes a Action cross referenmce array.
'       Element number = actual Action number (1-1000)
'       Value = Text array reference(1-260)


For i = 1 To 260
ThisVarNumber = Val(VarText(i))
    If ThisVarNumber > 0 Then
        VarXRef(ThisVarNumber) = i
    Else
        VarXRef(ThisVarNumber) = -1
    End If
Next

End Sub


Private Sub ActiveVarList_Click()
Dim AllText As String
Dim MyPos As Long

'strip off time data from string and put in Selected Action box

SelectedVar.text = Mid(ActiveVarList.List(ActiveVarList.ListIndex), 1, 50)
SelectedVar.Enabled = True


End Sub

Private Sub ActiveVarList_GotFocus()
'FocusOnActiveVarList = True

End Sub


Private Sub ActiveVarList_LostFocus()
'FocusOnActiveVarList = False

End Sub


Private Sub Cancel_Click()
Unload Me

End Sub

Private Sub ClearAll_Click()
Dim i As Integer


' Clear a ALL Actions...
    For i = 1 To 8
       'PlotVar(i) = 0
        PlotVar(i) = 0
        VarChanges = True
        'Call Logging(SelectedVar.text & " De-Selected from Plot ", LogVars)
        'SelectedVar.text = ClearText
        LegendText(i) = "Not Allocated"
        GraphFrm.Label2(i - 1) = LegendText(i)
    Next
    VarChanges = True
    Call Logging("All Plot Variables Cleared by Student " & Time, LogVars)

    SelectedVar.text = ClearText


'deselect any item from list box so it can be updated
Call Timer1_Timer
End Sub

Private Sub Combo2_Click()
Dim i As Integer


Select Case Combo2.ListIndex
    Case 0
      'fast
      Speed = 500 'milli seconds
    Case 1
      'medium
      Speed = 1000
    Case 2
      'slow
      Speed = 2000
    Case 3
      'very slow
      Speed = 5000
      
    Case 4
      Speed = 10000
    
    Case 5
       Speed = 20000
   
    Case 6
       Speed = 30000
    
    Case 6
       Speed = 60000
   
  End Select
  'StrSpeed = Combo2.List(Combo2.ListIndex) 'Put the text into string to display
  
  GraphFrm.Timer1.Enabled = True
  GraphFrm.Timer1.Interval = Speed
  GraphFrm.Label3.Caption = Format(str(Timer1.Interval / 1000), "##0.0") & Space(1) & "sec"



End Sub

Private Sub Description_Click()
Dim ValueRequiredString As String * 1
Dim length As Integer
Dim i As Integer
Dim SpecialAction As Byte

'---------------------------------------------------------------------------------------------------------------------
'FUNCTION:
'Puts selected Action from Description list into Selected Action Box
'If Action is already entered, show it in black and enable the clear buttons, disable the enter
'Show it in grey and enable Enter button and disable clear buttons
'If it is a special Action, show the spin button, value box and current value
'-------------------------------------------------------------------------------------------------------------------


'Check if selected Action is already entered, if yes then just offer clear functions and put Selected Text in bold

SelectedVar.text = Description.List(Description.ListIndex)
ThisVarNumber = Val(Mid$(SelectedVar.text, 1, 5))
'Debug.Print "SelectedVar.Text = " & SelectedVar.Text
'Debug.Print "ThisVarNumber = " & ThisVarNumber
If ThisVarNumber > 0 And ThisVarNumber < BiggestVarNumber Then
    For i = 1 To 8
        If PlotVar(i) = ThisVarNumber Then        'ie Action already entered, offer the clear button
            EnterCommand.Enabled = False
            ClearCommand.Enabled = True
            SelectedVar.Enabled = True
            Exit For
    
        Else
            'Enable Enter button and disable clear buttons
            EnterCommand.Enabled = True    'enables ENTRY
            ClearCommand.Enabled = False  'disables CLEAR
        End If
    Next
End If

End Sub

Sub Description_DblClick()
Dim DummyText As String

'does the same as Action Entry except action goes in directly

ActionEntryFlag = True

'if Action not already entered, call Action entry routine
DummyText = Description.List(Description.ListIndex)
ThisVarNumber = Val(DummyText)

If Not Action(1, ThisVarNumber) Then
    Call EnterCommand_Click
End If


End Sub

Private Sub ClearCommand_Click()
Dim Offset As Integer
Dim i As Integer
 'Clear selected Action text  to show Action is cleared
 
    ThisVarNumber = Val(SelectedVar.text)
    Debug.Print "in clear SelectedVar.Text = " & SelectedVar.text & "; Val = " & ThisVarNumber
    If ThisVarNumber > 0 And ThisVarNumber < BiggestVarNumber Then
          For i = 1 To 8
            If PlotVar(i) = ThisVarNumber Then
                PlotVar(i) = 0
                VarChanges = True
                Call Logging(SelectedVar.text & " De-Selected from Plot ", LogVars)
                SelectedVar.text = ClearText
                LegendText(i) = "Not Allocated"
                GraphFrm.Label2(i - 1) = LegendText(i)
                Exit For
            End If
          Next
    End If
 
 'update Active Action List immediately
 Call Timer1_Timer


End Sub

Sub EnterCommand_Click()
'---------------------------------------------------------------------------------------------------------------------
'FUNCTION:
'This routine entered EITHER after enter button clicked OR by double clicking a Action in the description list
'-------------------------------------------------------------------------------------------------------------------



'does the same as DblClick and enters Actions

Dim ValueRequiredString As String * 1
Dim length As Integer
Dim i As Integer
Dim msg As String
Dim Title As String
Dim response As Integer
Dim DummyText As String
Dim PlotNo As Integer





'Get the Variable number
SelectedVar.Enabled = True

' === 14/12/02 19.158 ===
'ThisVarNumber = Val(SelectedVar.text)
ThisVarNumber = Val(Mid$(SelectedVar.text, 1, 4))

Debug.Print "in EnterCommand_Click() ThisVarNumber = " & ThisVarNumber
    
If (ThisVarNumber > 0 And ThisVarNumber <= BiggestVarNumber) Then
    For i = 1 To 8
        If PlotVar(i) = 0 Then
            PlotVar(i) = ThisVarNumber
            ThickLine(i) = Check1.Value 'set line thickness
            PlotNo = i
            LegendText(i) = SelectedVar.text
            VarChanges = True
            Exit For
        End If
    Next
    Debug.Print "Variable " & ThisVarNumber & " set "
    Var_time(PlotNo) = Time
    Call Logging(SelectedVar.text & " Selected to Plot ", LogVars)
End If


 
'clear the Selected Action text and make value box and button invisible
SelectedVar.text = ClearText
Check1.Value = 0
Description.ListIndex = -1                                                  'deselect any item from list box

'Variable is entered so diable the Enter command and enable the clear
EnterCommand.Enabled = False
ClearCommand.Enabled = True
 'update Active Action List immediately
Call Timer1_Timer
End Sub

Private Sub VarTypes_Click()
Dim i As Integer

Description.Clear
i = VarTypes.ListIndex
Call FillDescription(VarGroup(i, 1), VarGroup(i, 2))
 

End Sub



Private Sub Form_Load()
Dim i As Integer
Dim j As Long
Dim ii As Integer
Dim Filepathname As String
Dim LastLine As Byte
Dim FG1 As Integer
Dim FG2 As Integer

Dim counter As Integer
Dim ErrorFlag As Byte
Dim NumberofTypes As Integer
Dim MatchPosition As Integer
Dim VTypeStr As String
Dim MyChar As String * 1


On Error GoTo Errh1
ErrorFlag = False
ClearText = Space(50)
 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)
 
 
'read in strings until the end of file for labels & captions on form
Call GetLanguageStrings("PlotVars.txt", ErrorFlag)


frm_PlotVars.Caption = Language_Str(1)
Label1.Caption = Language_Str(2)
Label2.Caption = Language_Str(3)
Label3.Caption = Language_Str(4)
Label4.Caption = Language_Str(5)

EnterCommand.Caption = Language_Str(6)
ClearCommand.Caption = Language_Str(7)

ClearAll.Caption = Language_Str(8)
OK.Caption = Language_Str(9)


'read in strings until the end of file for action descriptions
'11/09/00

' === 30/11/01 ===
'Call GetLanguageStrings("SteamVarsDescription.txt", ErrorFlag)
Call GetLanguageStrings("Steam_AO_text.txt", ErrorFlag)

'12/09/00 clear array containing addresses of start & end points for each group
    For i = LBound(VarGroup, 1) To UBound(VarGroup, 1)
        VarGroup(i, 1) = 0
        VarGroup(i, 2) = 0
    Next
    BiggestVarNumber = 0


If Not ErrorFlag Then
    NumberofTypes = 0
    counter = 1
    FG1 = 0
    FG2 = -1
    LastLine = False
    i = 0
    Do
        Language_Str(i) = Trim(Language_Str(i)) 'get rid of leading & trailing spaces
        VTypeStr = ""
        
        LastLine = Language_Str(i) Like "*END OF TEXT*"    ' Returns True if a match
        If LastLine = False Then
            MatchPosition = InStr(1, Language_Str(i), "{", 1)
            If MatchPosition <> 0 Then 'it is a group title, so put it in type list and save next address
                
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change curly bracket for a space
                MatchPosition = InStr(1, Language_Str(i), "}", 1) 'find other curly bracket
                If MatchPosition <> 0 Then
                  Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change other curly bracket for a space
                End If
                
                Var_Types(FG1) = Language_Str(i)
                
                'Set this start address
                VarGroup(FG1, 1) = counter
                'also set end address of previous group - provided its not the first group
                
                If FG1 > 0 Then
                    VarGroup(FG1 - 1, 2) = counter - 1
                End If

                'Debug.Print "set group for FG1 = " & FG1 & ",  " & Language_Str(i)
                FG1 = FG1 + 1
                NumberofTypes = NumberofTypes + 1
            Else
                    'Debug.Print "found Action and  FG2 = " & FG2
                  If FG1 > 0 Then
                    'Debug.Print "FG2 = " & FG2 & "FOR :  " & Language_Str(i)
                    'VarGroup(FG1 - 1, 2) = counter            'it may be the last entry in this group. If not keep incrementing

                    If Len(Language_Str(i)) > 6 Then 'it must be a genuine entry
                        ThisVarNumber = Val(Mid$(Language_Str(i), 1, 4))   'get the actual Variable number from first 4 characters of the text
                        If ThisVarNumber <= UBound(VarText) Then
                          VarText(counter) = Language_Str(i) 'store the Variable text sequentially in an array
                          counter = counter + 1
                        End If
                    
                        If ThisVarNumber > BiggestVarNumber Then
                            BiggestVarNumber = ThisVarNumber
                        End If
                  
                    End If
                  End If
            End If
        Else
            VarGroup(FG1 - 1, 2) = counter - 1 'set end address of last group
            Exit Do
        End If
        i = i + 1
        If i > 2000 Then ' in case END OF TEXT is not found
            Exit Do
        End If
    Loop Until LastLine = True
End If

'Fill Grid Style Box with options
Combo1.AddItem "none"
Combo1.AddItem "x-axis only"
Combo1.AddItem "y-axis only"
Combo1.AddItem "both axes"
'Combo2.ListIndex = 1


VarTypes.Clear
Description.Clear
VarChanges = True ' to force refresh of list

For i = 0 To NumberofTypes - 1
   VarTypes.AddItem Var_Types(i)
Next i

For i = 0 To 8
  Debug.Print "VarGroup(" & i & ") = " & VarGroup(i, 1) & " and " & VarGroup(i, 2)
  'Debug.Print "Var_Types(" & i & ") = " & Var_Types(i)
Next



  
  With Combo2
    .AddItem "0.5 sec"  'added 11/09/00
    .AddItem "1 sec"
    .AddItem "2 sec"
    .AddItem "5 sec"
    .AddItem "10 sec"
    .AddItem "20 sec"
    .AddItem "30 sec"
    .AddItem "1 min"

    'added 11/09/00.ListIndex = 0
    .ListIndex = 1 'default to 1 sec
    
  End With
  
  GraphFrm.Timer1.Enabled = True
  GraphFrm.Timer1.Interval = 1000 'default
  

Exit Sub

Errh1:


End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 
End Sub


Private Sub OK_Click()
    'frm_Plot1.Show
    Unload Me
End Sub


Private Sub Timer1_Timer()
Dim i As Integer

'calls Fill_active_Var_list to update list. Add new variables, remove old variables
'Enable buttons according to number of Variables entered etc


 'enable/disable clear all button as appropriate
    If NumberOfVars > 0 Then
            ClearAll.Enabled = True
    Else
            ClearAll.Enabled = False
    End If
    If Val(SelectedVar.text) > 0 Then 'there is an action selected
        For i = 1 To 8
            If Val(PlotVar(i)) = Val(SelectedVar.text) Then 'enable the CLEAR if variable is in list
                'Debug.Print "Val(PlotVar(i)) = " & Val(PlotVar(i))
                'Debug.Print "Val(Selectedvar.Text) = " & Val(SelectedVar.Text)
                'Check1.Value = ThickLine(i)
                ClearCommand.Enabled = True
                EnterCommand.Enabled = False
                Exit For
            Else                                        'enable the ENTER if action is not in
                ClearCommand.Enabled = False
                EnterCommand.Enabled = True
                
            End If
        Next
    Else    'no action is selected
        ClearCommand.Enabled = False
        EnterCommand.Enabled = False
    End If
    
    Call Fill_Active_Var_List(frm_PlotVars)
    
Exit Sub

End Sub



