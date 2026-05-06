VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form SnapCompareForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snapshot Reload"
   ClientHeight    =   12615
   ClientLeft      =   3675
   ClientTop       =   5430
   ClientWidth     =   10845
   ForeColor       =   &H000000FF&
   Icon            =   "SnapCompare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12615
   ScaleWidth      =   10845
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   195
      Left            =   7740
      Top             =   4260
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   344
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Engine Room  AIR ON"
      Height          =   735
      Left            =   6720
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN I/O"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox AIList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Top             =   4620
      Width           =   10275
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Text            =   "Snapshot Reload"
      Top             =   60
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox DIList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3210
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10275
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1440
      Top             =   120
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "RUN I/O first  ( before EXIT )"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Adjust Switches and Potentiometers until both lists are clear."
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   660
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Potentiometers"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4380
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Switches"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   660
      Width           =   1275
   End
End
Attribute VB_Name = "SnapCompareForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer

' === 27/11/01 ===
Private AI_update_the_list As Boolean


Private Sub Command1_Click()


'  10/5/00  We need the ER control air ON before running the model so...
'   Runnining the I/O first will get the DO( ) status for the air solenoids in the ER
'   ( will have been saved in the snapshot )
'   The I/O program will then send this to the ER
            RunIO = True


End Sub

Private Sub Command2_Click()
Dim msg As Integer

If (DIDiffIndex <> OldDIDiffIndex) Or AIDifferences Then
    response = MsgBox("You have not reset all all inputs to the state they were in " & Chr(10) & _
              "when this snapshot was taken. Are you sure you want to continue ?", vbYesNo, "Warning")
              
  If response = vbNo Then
    Exit Sub
  End If

End If
'29/11/99 start after loading
            RunIO = True
            'Call set_RunIO(VB_RunIO)
                '29/11/99 start after loading
                d_run = True   ' Start  the     Diesel  Simulator
                e_run = True   ' Start  the Electrical  Simulator

Unload Me

End Sub

Private Sub Command3_Click()


'  10/5/00  TESTING

'    If i = 1137 Or i = 1155 And Check2.Value Then 'Fix to allow boiler valves to have air
'        D_OUTPUT(i) = True
'    Else
'        D_OUTPUT(i) = False
'    End If

'  PUT Engine Room AIR ON
D_OUTPUT(1137) = True
D_OUTPUT(1137) = True


End Sub

Private Sub Form_Load()

' === 27/11/01 ===
AI_update_the_list = True


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Timer1_Timer()
Dim i As Integer


'Every few seconds check if there are any DI or AI differences and
' if so, display text in ListBoxes

Call DI_Compare
'If DIDifferences Then
    If DIDiffIndex <> OldDIDiffIndex Then
        OldDIDiffIndex = DIDiffIndex
        DIList.clear
        If DIDiffIndex > 0 And DIDiffIndex < UBound(DIReloadText) Then
          For i = 1 To DIDiffIndex
              DIList.AddItem DIReloadText(i)
          Next
        End If
    End If
    
'End If
'DIList.AddItem "Hello " & Time


Call AI_Compare


If AIDifferences Then

    ' === 27/11/01 ===
    If AI_update_the_list Then
       AI_update_the_list = False
       
        'If AIDiffIndex <> OldAIDiffIndex Then
            OldAIDiffIndex = AIDiffIndex
            AIList.clear
            If AIDiffIndex > 0 And AIDiffIndex < UBound(AIReloadText) Then
                For i = 1 To AIDiffIndex
                    AIList.AddItem AIReloadText(i)
                Next
            End If
        'End If
        
    ' === 27/11/01 ===
    End If

Else
    If OldAIDiffIndex > 0 Then
        AIList.clear
      OldAIDiffIndex = 0
    End If
End If

End Sub
Sub AI_Compare()
Dim max As Integer
Dim Index As Integer
Dim UD As String * 4
Dim PercentDiff As Integer
Dim AITextStr As String * 50

Dim AI_Old As Integer
Dim AI_New As Integer
Dim DiffText As String

'get AI status as stored in snapshot
Call get_AI_Reloads(VB_AI_Reload(0))


AIDiffIndex = 0
AIDifferences = False

' compare stored AIs with current AIs for Pots
' If any different, set flag (AIDifferences)
' store text for any differences in AIReloadText array

  'While index <= max
  For Index = 0 To 255
    'Debug.Print "index = " & index
   
    Select Case Index
    
    Case 0 To 1, 4 To 19, 40 To 41, 44 To 46
    '    MECD  , ERoom  , E Mimic , E S/B  Pots
        AI_Old = VB_AI_Reload(Index)
        AI_New = A_INPUT(Index)
        
        If Abs(AI_Old - AI_New) > 4 Then
                    
            If AI_Old > AI_New Then
              UD = "UP"
            Else
              UD = "DOWN"
            End If
            PercentDiff = (AI_Old - AI_New) * 100 / 255
            AIDiffIndex = AIDiffIndex + 1
            AIDifferences = True
            '16/11/99  AIDataBase.MoveFirst
            'AIDataBase.Find "[i/o ref] =" & index
            'If Not (AIDataBase.EOF) Then
                'AIReloadText(AIDiffIndex) = AIDataBase(1)
            'Call GetHardIOText("AI", AIDiffIndex, AIReloadText(AIDiffIndex))
            '12/12/12
            If Index = 0 Then
                AIReloadText(AIDiffIndex) = "Throttle Local Pointer Position"
                'this is a cheat because direct access file can't hold data in record 0 All others OK.
            Else
                Call GetHardIOText("AI", Index, AIReloadText(AIDiffIndex))
            End If
            '12/12/00 Call GetHardIOText("AI", Index, AIReloadText(AIDiffIndex))
            DiffText = " Old = " & Format(AI_Old * 0.392, "##0") & "%  New = " & Format(AI_New * 0.392, "##0") & "%  Adjust " & UD & " By " & PercentDiff & "%"
            Mid$(AIReloadText(AIDiffIndex), 50, Len(DiffText)) = DiffText
            
            'Else
                'AIReloadText(AIDiffIndex) = "Unknown Potentiometer"
            'End If
        End If
        
        ' === 27/11/01 Added next 4 lines: ===
        ' 12/12/01 Got an overflow crash so REMOVED....
        'If Abs(A_INPUT_old(Index) - A_INPUT(Index)) > 4 Then
        '   A_INPUT_old(Index) = A_INPUT(Index)
           AI_update_the_list = True
        'End If
        
        
   End Select
   
 Next
 
 
 
 


End Sub

Sub DI_Compare()
Dim max As Integer
Dim Index As Integer

'get DI status as stored in snapshot
Call get_DI_Reloads(VB_DI_Reload(0))

max = 425

DIDiffIndex = 0
AIDiffIndex = 0
DIDifferences = False
AIDifferences = False


      '352 to 425
' compare stored DIs with current DIs for hard switches and ERoom
' If any different, set flag (DIDifferences)
' store text for any differences in DIReloadText array

  'While index <= max
  For Index = 323 To 956
    'Debug.Print "index = " & index
   
    Select Case Index
    
    Case 323 To 325, 352 To 425, 837 To 838, 847 To 848, 928 To 939, 945 To 956
    
    
    
        'Logic problem between VB & Fortran If (D_INPUT(index) <> VB_DI_Reload(index)) And VB_DI_Reload(index) Then
        If ((D_INPUT(Index) And Not VB_DI_Reload(Index)) _
            Or (Not D_INPUT(Index) And VB_DI_Reload(Index))) And VB_DI_Reload(Index) Then
            DIDiffIndex = DIDiffIndex + 1
            DIDifferences = True
            '11/11/99 no database
            'DIDataBase.MoveFirst
            'DIDataBase.Find "[i/o ref] =" & index
            'If Not (DIDataBase.EOF) Then
                'DIReloadText(DIDiffIndex) = "Select " & DIDataBase(1)
            'Else
                'DIReloadText(DIDiffIndex) = "Unknown Switch"
            'End If
            Call GetHardIOText("DI", Index, DIReloadText(DIDiffIndex))
            
        End If
   End Select
 Next
 
 
 
 
 
'Wend

  
  
'FaultTypes.clear
'Description.clear

'For i = 0 To NumberofTypes - 1
  ' FaultTypes.AddItem Fault_Types(i)
'Next i

End Sub

