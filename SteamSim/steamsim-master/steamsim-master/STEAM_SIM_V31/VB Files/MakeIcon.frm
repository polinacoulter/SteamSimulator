VERSION 5.00
Begin VB.Form frm_PlotVars 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plot variables "
   ClientHeight    =   9090
   ClientLeft      =   1920
   ClientTop       =   2205
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9090
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ClearCommand 
      Caption         =   "Clear"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      ToolTipText     =   "Clear the Selected Action Only"
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame FaultEntry 
      Height          =   8955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6360
         Top             =   3840
      End
      Begin VB.CommandButton OK 
         Caption         =   "OK"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton ClearAll 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         ToolTipText     =   "Clear All Actions that are currently entered"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox ActiveVarList 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2160
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   6735
      End
      Begin VB.CommandButton EnterCommand 
         Caption         =   "Enter"
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Selectedvar 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5295
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
         Left            =   3240
         TabIndex        =   2
         Top             =   4320
         Width           =   5115
      End
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
         TabIndex        =   1
         Top             =   4320
         Width           =   3075
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   7
         X1              =   6960
         X2              =   7440
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   6
         X1              =   6960
         X2              =   7440
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   5
         X1              =   6960
         X2              =   7440
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   4
         X1              =   6960
         X2              =   7440
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   3
         X1              =   6960
         X2              =   7440
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   2
         X1              =   6960
         X2              =   7440
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   1
         X1              =   6960
         X2              =   7440
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         Index           =   0
         X1              =   6960
         X2              =   7440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Variables to Plot"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Selected Variable"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Variable Types"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frm_PlotVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Fill_Active_Var_List(form_name As Form)
Dim i As Integer
Dim LenText As Integer
Dim VarString As String
Dim Padding As Integer


    If VarChanges Then ' only update list if there has been a change to plottrer configuration
        ActiveVarList.clear
        For i = 1 To 8
              If PlotVar(i) <> 0 Then ' there is an entry being plotted
                'VarString = VarText(VarXRef(i))
                'Debug.Print " found PlotVar(" & i & ") = " & PlotVar(i)
                'Debug.Print "VarText(VarXref(" & PlotVar(i) & ") = " & VarText(VarXRef(PlotVar(i)))
                VarString = Mid(VarText(VarXRef(PlotVar(i))), 1, LenText - 5)  'omit last 5 chars which are Units
                LenText = Len(VarString)
                Padding = 50 - LenText
                VarString = VarString + Space(Padding) & Var_time(i)
              Else
                VarString = "Not Allocated"
              End If
              form_name.ActiveVarList.AddItem VarString
        Next
        VarChanges = False
    End If
  
 
End Sub
Sub FillDescription(FirstAction As Integer, LastAction As Integer)
Dim i As Integer
Dim ActionOK As Integer
Dim ThisVarNumber As Integer
Dim ActionAlreadyEntered As Integer
Dim StrLen As Integer




For i = FirstAction To LastAction                       'check for range of Action numbers defined by group

     ThisVarNumber = Val(VarText(i))              'get Action number
     If ThisVarNumber > 0 Then
                StrLen = Len(VarText(i))
                'Description.AddItem " " & VarText(i)
                Description.AddItem " " & Mid(VarText(i), 1, StrLen - 5) 'omit last 5 chars which are Units
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

Selectedvar.Text = Mid(ActiveVarList.List(ActiveVarList.ListIndex), 1, 50)
Selectedvar.Enabled = True

'enable either the clear or enter command
If Action(1, Val(Selectedvar.Text)) Then
    EnterCommand.Enabled = False
    ClearCommand.Enabled = True
Else
    EnterCommand.Enabled = True
    ClearCommand.Enabled = False
End If

'12/3/98 ClearAll.Enabled = True 'now done in timer
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
       PlotVar(i) = 0
    Next
    VarChanges = True
     Call Logging("All Plot Variables Cleared by Student " & Time, LogVars)

     Selectedvar.Text = ClearText


'deselect any item from list box so it can be updated
Call Timer1_Timer
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

Selectedvar.Text = Description.List(Description.ListIndex)
ThisVarNumber = Val(Selectedvar.Text)
'Debug.Print "SelectedVar.Text = " & SelectedVar.Text
'Debug.Print "ThisVarNumber = " & ThisVarNumber
If ThisVarNumber > 0 And ThisVarNumber < BiggestVarNumber Then
If Action(1, ThisVarNumber) Then       'ie Action already entered, offer the cler button
    EnterCommand.Enabled = False
    ClearCommand.Enabled = True
    'ClearAll.Enabled = True
    Selectedvar.Enabled = True
Else
'Enable Enter button and disable clear buttons
        EnterCommand.Enabled = True    'enables ENTRY
        ClearCommand.Enabled = False  'disables CLEAR
        
        'If Action requires additional value input, show box but keep text grey until accepted
       
        
End If
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
Dim offset As Integer
Dim i As Integer
 'Clear selected Action text  to show Action is cleared
 
    ThisVarNumber = Val(Selectedvar.Text)
    Debug.Print "in clear SelectedVar.Text = " & Selectedvar.Text & "; Val = " & ThisVarNumber
    If ThisVarNumber > 0 And ThisVarNumber < BiggestVarNumber Then
          PlotVar(ThisVarNumber) = 0
          VarChanges = True
          Call Logging(Selectedvar.Text & " Cleared by Student " & Time, LogActions)
          Selectedvar.Text = ClearText
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
Selectedvar.Enabled = True
ThisVarNumber = Val(Selectedvar.Text)

Debug.Print "in EnterCommand_Click() ThisVarNumber = " & ThisVarNumber
    
If (ThisVarNumber > 0 And ThisVarNumber < BiggestVarNumber) Then
    For i = 1 To 8
        If PlotVar(i) = 0 Then
            PlotVar(i) = ThisVarNumber
            PlotNo = i
            VarChanges = True
            Exit For
        End If
    Next
    Debug.Print "Variable " & ThisVarNumber & " set "
    Var_time(PlotNo) = Time
    Call Logging(Selectedvar.Text & " Entered ", LogVars)
End If


 
'clear the Selected Action text and make value box and button invisible
Selectedvar.Text = ClearText
Description.ListIndex = -1                                                  'deselect any item from list box

'Variable is entered so diable the Enter command and enable the clear
EnterCommand.Enabled = False
ClearCommand.Enabled = True
 'update Active Action List immediately
Call Timer1_Timer
End Sub

Private Sub VarTypes_Click()
Dim i As Integer

Description.clear
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

Dim Counter As Integer
Dim ErrorFlag As Byte
Dim NumberofTypes As Integer
Dim MatchPosition As Integer
Dim VTypeStr As String
Dim MyChar As String * 1



ErrorFlag = False
ClearText = Space(50)
 
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
Call GetLanguageStrings("PlotVarsDescription.txt", ErrorFlag)

If Not ErrorFlag Then
    NumberofTypes = 0
    Counter = 1
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
            'If Language_Str(i) Like "*{*}*" Then 'it is a group title, so put it in type list and save next address
            If MatchPosition <> 0 Then 'it is a group title, so put it in type list and save next address
                
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change curly bracket for a space
                MatchPosition = InStr(1, Language_Str(i), "}", 1)
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change other curly bracket for a space

                Var_Types(FG1) = Language_Str(i)
      
                VarGroup(FG1, 1) = Counter

                'debug.print "set group for FG1 = " & FG1 & ",  " & Language_Str(i)
                FG1 = FG1 + 1
                NumberofTypes = NumberofTypes + 1
            Else
                If Language_Str(i) Like "###*" Then           'it is a Variable Text entry
                    'Debug.Print "found Action and  FG2 = " & FG2
                  If FG1 > 0 Then
                    'Debug.Print "FG2 = " & FG2 & "FOR :  " & Language_Str(i)
                    'VarGroup(FG1 - 1, 2) = i                  'it may be the last entry in this group. If not keep incrementing
                    VarGroup(FG1 - 1, 2) = Counter            'it may be the last entry in this group. If not keep incrementing

                    VarText(Counter) = Language_Str(i)        'store the Variable text sequentially in an array
                    ThisVarNumber = Val(Language_Str(i))     'get the actual Variable number from the text
                    If ThisVarNumber <= 2000 Then
                    VarXRef(ThisVarNumber) = Counter       'fill in the x-ref for retrieval later
                    Else
                        'error trap
                    End If
                    If ThisVarNumber > BiggestVarNumber Then
                        If ThisVarNumber <= 1400 Then BiggestVarNumber = ThisVarNumber
                    End If
                    Counter = Counter + 1
                  
                  End If
                End If
            End If
        Else
            Exit Do
        End If
        i = i + 1
        If i > 2000 Then ' in case END OF TEXT is not found
            Exit Do
        End If
    Loop Until LastLine = True
End If


VarTypes.clear
Description.clear

For i = 0 To NumberofTypes - 1
   VarTypes.AddItem Var_Types(i)
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 
End Sub


Private Sub OK_Click()
    Unload Me
End Sub


Private Sub Timer1_Timer()
Dim i As Integer

'calls Fill_active_Var_list to update list. Add new variables, remove old variables
'Enable buttons according to number of Variables entered etc


 'enable/disable clear all button as appropriate
    If NumberOfActions > 0 Then
            ClearAll.Enabled = True
    Else
            ClearAll.Enabled = False
    End If
    If Val(Selectedvar.Text) > 0 Then 'threre is an action selected
        If Action(1, Val(Selectedvar.Text)) Then 'enable the CLEAR if action is in
            ClearCommand.Enabled = True
            EnterCommand.Enabled = False
        Else                                        'enable the ENTER if action is not in
            ClearCommand.Enabled = False
            EnterCommand.Enabled = True
        End If
    Else    'no action is selected
        ClearCommand.Enabled = False
        EnterCommand.Enabled = False
    End If
    
    Call Fill_Active_Var_List(frm_PlotVars)
    
Exit Sub

End Sub



