VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Begin VB.Form frm_Student_Actions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Action Entry"
   ClientHeight    =   12795
   ClientLeft      =   1920
   ClientTop       =   2250
   ClientWidth     =   10305
   FillStyle       =   0  'Solid
   Icon            =   "frm_Student_Actions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12795
   ScaleWidth      =   10305
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   12675
      Left            =   60
      Top             =   90
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   22357
      Begin VB.Frame FaultEntry 
         Height          =   12645
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   9675
         Begin VB.ListBox ActionTypes 
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
            Height          =   8220
            Left            =   120
            TabIndex        =   10
            Top             =   4320
            Width           =   3075
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
            Height          =   8220
            Left            =   3240
            TabIndex        =   9
            Top             =   4320
            Width           =   6375
         End
         Begin VB.TextBox SelectedAction 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   540
            Width           =   8295
         End
         Begin VB.CommandButton EnterCommand 
            Caption         =   "Enter"
            Height          =   375
            Left            =   8580
            TabIndex        =   7
            Top             =   300
            Width           =   975
         End
         Begin VB.ListBox ActiveActionList 
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
            TabIndex        =   6
            Top             =   1560
            Width           =   8295
         End
         Begin VB.CommandButton ClearAll 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   8580
            TabIndex        =   5
            ToolTipText     =   "Clear All Actions that are currently entered"
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton OK 
            Caption         =   "OK"
            Height          =   375
            Left            =   8580
            TabIndex        =   4
            Top             =   3840
            Width           =   975
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   8760
            Top             =   960
         End
         Begin VB.CommandButton RefreshListCommand 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   8580
            TabIndex        =   3
            ToolTipText     =   "Clears Previous Actions from List. Show current actions only"
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "ActionTypes"
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Selected Action"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Time"
            Height          =   255
            Left            =   7500
            TabIndex        =   12
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Actions Entered"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.CommandButton ClearCommand 
         Caption         =   "Clear"
         Height          =   375
         Left            =   8580
         TabIndex        =   1
         ToolTipText     =   "Clear the Selected Action Only"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox List1 
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
         Height          =   7740
         Left            =   9720
         TabIndex        =   0
         Top             =   900
         Visible         =   0   'False
         Width           =   4755
      End
   End
End
Attribute VB_Name = "frm_Student_Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Fill_Active_Action_List(form_name As Form)
Dim i As Integer
Dim j As Integer
Dim ActionInListOK As Byte
Dim ItemToRemove As Integer
Dim DummyText As String
Dim PositionQ As Integer
Dim String1 As String
Dim String2 As String
Dim TimeStr As String
Dim SingleScan As Integer
Dim Offset As Integer
Dim SSFreqAft11, SSFreqFord14, SSVoltsAft71, SSVoltsAft74 As Single
Dim ActionOK As Integer
Dim ActionAlreadyEntered As Integer
Dim LenText As Integer
Dim ActString As String
Dim TText As String
Dim Padding As Integer


' ===24/3/98===
   For i = 1 To BiggestActionNumber
           Action(1, i) = False
           Action(2, i) = False
   Next i
    For i = 1 To 32
       If active_actions(i) > 0 Then
          Action(1, active_actions(i)) = True
       End If
        If active_reset_actions(i) > 0 Then
          Action(2, active_reset_actions(i)) = True
       End If
   Next i
   'Debug.Print "active_actions(1:3) =  " & active_actions(1) & "  " & active_actions(2) & "  " & active_actions(3)
'==============
    
    NumberOfActions = 0
    For i = 1 To BiggestActionNumber
    
        If Action(1, i) Then 'if a Action is really is active......
           
                NumberOfActions = NumberOfActions + 1
                'TimeStr = Time
                TimeStr = Format(Time, "h:m:s")
                ActionInListOK = False
                For j = 0 To form_name.ActiveActionList.ListCount - 1              'check if it is in the Active Action list
                        If Val(form_name.ActiveActionList.List(j)) = i Then
                            'ActionInListOK = True                                   'found it in list so don't add it again unless
                                                                                    'entry contains the word CLEARED
                            'Debug.Print "cleared at " & InStr(1, ActiveActionList.List(j), "CLEARED", 1)
                            If InStr(1, ActiveActionList.List(j), "CLEARED", 1) = 0 Then  'don't add entry
                                ActionInListOK = True
                            Else
                                ActionInListOK = False
                            End If
                        End If
                 Next
                If ActionInListOK = False Then   'there has been an action detected which is not in the list                                                     'if not the add it to the list using the Xref array
                         
                            'Action_time(i) = TimeStr
                            If Val(Action_time(i)) = 0 Then
                                Action_time(i) = TimeStr 'if list is only being refreshed, then orig time will be stored
                            End If
                            ActString = ActionText(ActionXRef(i))
                            
                            Debug.Print "setting action, text = " & ActString

                            LenText = Len(ActString)
                            
                            If LenText <= 70 Then
                                Padding = 70 - LenText
                            Else
                                Padding = 1
                            End If
                            If Padding > 0 Then
                                ActString = ActString + Space(Padding) & Action_time(i)
                            Else
                                ActString = ActString & Action_time(i)
                            End If
                       
                            form_name.ActiveActionList.AddItem ActString
                 End If
         End If
'24/3/98 If Action(2, i) Then
         If Action(1, i) And Action(2, i) Then
                         TimeStr = Format(Time, "h:m:s")
                         Action_time(i) = TimeStr
                         ActString = ActionText(ActionXRef(i)) & " CLEARED"
                         LenText = Len(ActString)
                            
                         If LenText <= 50 Then
                            Padding = 50 - LenText
                         Else
                            Padding = 1
                         End If
                         
                         ActString = ActString + Space(Padding) & Action_time(i)
                         form_name.ActiveActionList.AddItem ActString
             ' 24/3/98   Action(1, i) = False
             ' 24/3/98   Action(2, i) = False
                         
              ' 24/3/98   Send the reset action to the model
              
                         Call setACTION(12, i)    ' 12 For FALSE
                         Call setACTION(11, i)    ' 11 For FALSE
                                               
                         Action_time(i) = Space(8)
                                                                                                      
         End If
          
   Next i
   
  
   
 
End Sub
Sub FillDescription(FirstAction As Integer, LastAction As Integer)
Dim i As Integer
Dim ActionOK As Integer
Dim ThisActionNumber As Integer
Dim ActionAlreadyEntered As Integer



For i = FirstAction To LastAction                                     'check for range of Action numbers defined by group

     ThisActionNumber = Val(ActionText(i))              'get Action number
     If ThisActionNumber > 0 Then
                Description.AddItem " " & ActionText(i)
     End If
 
   
Next i

End Sub



Sub FillXref()
Dim i As Integer


'Makes a Action cross referenmce array.
'       Element number = actual Action number (1-1000)
'       Value = Text array reference(1-260)


For i = 1 To 260
ThisActionNumber = Val(ActionText(i))
    If ThisActionNumber > 0 Then
        ActionXRef(ThisActionNumber) = i
    Else
        ActionXRef(ThisActionNumber) = -1
    End If
Next

End Sub


Private Sub ActiveActionList_Click()
Dim AllText As String
Dim MyPos As Long

'strip off time data from string and put in Selected Action box

SelectedAction.text = Mid(ActiveActionList.List(ActiveActionList.ListIndex), 1, 50)
SelectedAction.Enabled = True

'enable either the clear or enter command
If Action(1, Val(SelectedAction.text)) Then
    EnterCommand.Enabled = False
    ClearCommand.Enabled = True
Else
    EnterCommand.Enabled = True
    ClearCommand.Enabled = False
End If

'12/3/98 ClearAll.Enabled = True 'now done in timer
End Sub

Private Sub ActiveActionList_GotFocus()
'FocusOnActiveActionList = True

End Sub


Private Sub ActiveActionList_LostFocus()
'FocusOnActiveActionList = False

End Sub


Private Sub Cancel_Click()
Unload Me

End Sub

Private Sub ClearAll_Click()
Dim i As Integer


' === 28/03/01 ===
If Not I_AM_INSTRUCTOR Then
     ' I MUST be a STUDENT
          
          If LSS_LOGICALS(64) Then
              ' The instructor has DIS-ABLED me so exit this sub
            Exit Sub
          End If
End If




' Clear a ALL Actions...
' 24/3/98
    For i = 1 To BiggestActionNumber
       If Action(1, i) Then Action(2, i) = True
    Next
    
               Call setACTION(3, 1)    ' 3 to clear ALL Actions
    
     Call Logging("All Actions Cleared by Student ", LogActions)

     SelectedAction.text = ClearText
     

'deselect any item from list box so it can be updated
Call Timer1_Timer

' 24/3/98 ADDED:
    ActiveActionList.clear 'clear list  so it can be updated



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

SelectedAction.text = Description.List(Description.ListIndex)
ThisActionNumber = Val(SelectedAction.text)
'Debug.Print "SelectedAction.Text = " & SelectedAction.Text
'Debug.Print "ThisActionnumber = " & ThisActionnumber
If ThisActionNumber > 0 And ThisActionNumber < BiggestActionNumber Then
If Action(1, ThisActionNumber) Then       'ie Action already entered, offer the cler button
    EnterCommand.Enabled = False
    ClearCommand.Enabled = True
    'ClearAll.Enabled = True
    SelectedAction.Enabled = True
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
ThisActionNumber = Val(DummyText)

If Not Action(1, ThisActionNumber) Then
    Call EnterCommand_Click
End If


End Sub

Private Sub ClearCommand_Click()
Dim Offset As Integer
Dim i As Integer


' === 28/03/01 ===
If Not I_AM_INSTRUCTOR Then
     ' I MUST be a STUDENT

          If LSS_LOGICALS(64) Then
              ' The instructor has DIS-ABLED me so exit this sub
            Exit Sub
          End If
End If


 'Clear selected Action text  to show Action is cleared
 
    ThisActionNumber = Val(SelectedAction.text)
    Debug.Print "in clear SelectedAction.Text = " & SelectedAction.text & "; Val = " & ThisActionNumber
    If ThisActionNumber > 0 And ThisActionNumber < BiggestActionNumber Then
    
    ' 24/3/98   Action(2, ThisActionNumber) = True
               
                  Call setACTION(2, ThisActionNumber)    '  2 For true
     ' 24/3/98    Call setACTION(11, ThisActionNumber)    ' 11 For FALSE
         
            Call Logging("Student CLEARED Action :  " & SelectedAction.text, LogActions)
            SelectedAction.text = ClearText
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
Dim Response As Integer
Dim DummyText As String
Dim SpecialAction As Byte
Dim Freq As String
Dim Volts As String
Dim Offset As Integer




' === 28/03/01 ===
If Not I_AM_INSTRUCTOR Then
     ' I MUST be a STUDENT
          
          If LSS_LOGICALS(64) Then
              ' The instructor has DIS-ABLED me so exit this sub
            Exit Sub
          End If
End If



'if entered from double click, pick up text from description list
'If ActionEntryFlag Then
   ' ActionEntryFlag = False
   ' DummyText = Description.List(Description.ListIndex)
    'length = Len(DummyText)
'End If




'Get the Action number
SelectedAction.Enabled = True
ThisActionNumber = Val(SelectedAction.text)

Debug.Print "in EnterCommand_Click() ThisActionNumber = " & ThisActionNumber
    
       
If (ThisActionNumber > 0 And ThisActionNumber <= BiggestActionNumber) Then
    
    
' 24/3/98  Actions are returned from the model
'    Action(1, ThisActionNumber) = True                  'Set the Action
    
  ' 24/3/98   Send the action to the model
    Call setACTION(1, ThisActionNumber)

    
    Debug.Print "action " & ThisActionNumber & " set "
    Action_time(ThisActionNumber) = Time
    Call Logging("Student ENTERED Action :  " & SelectedAction.text, LogActions)
End If


   

    
'End If


'clear the Selected Action text and make value box and button invisible
SelectedAction.text = ClearText
Description.ListIndex = -1                                                  'deselect any item from list box

'Action is entered so diable the Enter command and enable the clear
EnterCommand.Enabled = False
ClearCommand.Enabled = True
 'update Active Action List immediately
Call Timer1_Timer
End Sub

Private Sub ActionTypes_Click()
Dim i As Integer

Description.clear
i = ActionTypes.ListIndex
Call FillDescription(ActionGroup(i, 1), ActionGroup(i, 2))
 

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
Dim FTypeStr As String
Dim MyChar As String * 1


' 20/1/99  TESTING RWJ
List1.clear

ErrorFlag = False
ClearText = Space(50)
 
 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)
 

 
'read in strings until the end of file for labels & captions on form
Call GetLanguageStrings("Actions.txt", ErrorFlag)


frm_Student_Actions.Caption = Language_Str(1)
Label1.Caption = Language_Str(2)
Label2.Caption = Language_Str(3)
Label3.Caption = Language_Str(4)
Label4.Caption = Language_Str(5)
Label5.Caption = Language_Str(6)

EnterCommand.Caption = Language_Str(7)
ClearCommand.Caption = Language_Str(8)

ClearAll.Caption = Language_Str(9)
OK.Caption = Language_Str(10)


'read in strings for actual actions until the end of file for action descriptions
    Call GetLanguageStrings("SteamActionsText.txt", ErrorFlag)


If Not ErrorFlag Then
    NumberofTypes = 0
    counter = 1
    FG1 = 0
    FG2 = -1
    LastLine = False
    i = 0
    Do
        Language_Str(i) = Trim(Language_Str(i)) 'get rid of leading & trailing spaces
        FTypeStr = ""
        
        LastLine = Language_Str(i) Like "*END OF TEXT*"    ' Returns True if a match
        If LastLine = False Then
            MatchPosition = InStr(1, Language_Str(i), "{", 1)
            'If Language_Str(i) Like "*{*}*" Then 'it is a group title, so put it in type list and save next address
            If MatchPosition <> 0 Then 'it is a group title, so put it in type list and save next address
                
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change curly bracket for a space
                MatchPosition = InStr(1, Language_Str(i), "}", 1)
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change other curly bracket for a space

                Action_Types(FG1) = Language_Str(i)
      
                ActionGroup(FG1, 1) = counter

                'debug.print "set group for FG1 = " & FG1 & ",  " & Language_Str(i)
                FG1 = FG1 + 1
                NumberofTypes = NumberofTypes + 1
            Else
                If Language_Str(i) Like "###*" Then           'it is a Action entry
                  'Find if the string "(PTT)" is in the action text. This marks it as a PTT action only
                  MatchPosition = InStr(1, Language_Str(i), "(PTT)", 1)
                  If MatchPosition <> 0 Then 'Strip text "(PTT)" from action text
                    Mid(Language_Str(i), MatchPosition, 5) = Space(5)
                  End If
                  'Only list If Either
                  '   It is a PartTaskTrainer (PTT)
                  'Or
                  '   It is Full Mission and action is NOT a PTT action (matchposition = 0)
                  
                  If PTT Or (Not PTT And MatchPosition = 0) Then
                    If FG1 > 0 Then
                        ActionGroup(FG1 - 1, 2) = counter            'it may be the last entry in this group. If not keep incrementing
                        ActionText(counter) = Language_Str(i)        'store the Action text sequentially in an array
                        List1.AddItem counter & "  " & ActionText(counter)
                        ThisActionNumber = Val(Language_Str(i))     'get the actual Action number from the text
                        If ThisActionNumber <= 2000 Then
                            ActionXRef(ThisActionNumber) = counter       'fill in the x-ref for retrieval later
                        Else
                            'error trap
                        End If
                        If ThisActionNumber > BiggestActionNumber Then
                            If ThisActionNumber <= 1400 Then BiggestActionNumber = ThisActionNumber
                        End If
                        counter = counter + 1
                  
                    End If
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





ActionTypes.clear
Description.clear

For i = 0 To NumberofTypes - 1
   ActionTypes.AddItem Action_Types(i)
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 
End Sub


Private Sub OK_Click()
    Unload Me
End Sub

Private Sub PTT_Test_Click()
PTT = Not PTT

End Sub

Private Sub RefreshListCommand_Click()
    ActiveActionList.clear 'clear list  so it can be updated
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

' === 28/03/01 ===
Dim Instructor_disablement As Boolean


Instructor_disablement = False

If Not I_AM_INSTRUCTOR Then
     ' I MUST be a STUDENT

          If LSS_LOGICALS(64) Then
          
              ' The instructor has DIS-ABLED Enter AND Clear AND ClearAll so
              Instructor_disablement = True
                            
          End If
End If


'calls Fill_active_Action_list to update list. Add new Actions, remove old Actions
'Enable buttons according to number of Actions entered etc
Call Fill_Active_Action_List(frm_Student_Actions)


'enable/disable clear all button as appropriate
' === 28/03/01 ===
'If NumberOfActions > 0 Then
If NumberOfActions > 0 And Not Instructor_disablement Then
            ClearAll.Enabled = True
Else
            ClearAll.Enabled = False
End If
    
    
    
' === 28/03/01 ===
'If Val(SelectedAction.text) > 0 Then    'there is an action selected
 If Val(SelectedAction.text) > 0 And Not Instructor_disablement Then
        'there is an action selected
    
    
        If Action(1, Val(SelectedAction.text)) Then 'enable the CLEAR if action is in
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
    
    
Exit Sub

End Sub



