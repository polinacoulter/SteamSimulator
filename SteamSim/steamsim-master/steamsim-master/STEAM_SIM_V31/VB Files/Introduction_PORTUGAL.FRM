VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "Comctl32.ocx"
Object = "{5FC43928-557A-11CF-BF88-0040956003D8}#4.0#0"; "fximg432.ocx"
Begin VB.Form frm_intro_PORTUGAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LSS-3M Marine Diesel Engine Simulator"
   ClientHeight    =   13995
   ClientLeft      =   15
   ClientTop       =   1635
   ClientWidth     =   18690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   13995
   ScaleWidth      =   18690
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Simulator Instructor Information"
      Height          =   2295
      Left            =   12480
      TabIndex        =   1
      Top             =   11160
      Width           =   6255
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5280
         Top             =   960
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   1800
         Width           =   255
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   4440
         Top             =   360
      End
      Begin VB.CommandButton Accept 
         Caption         =   "Exit"
         Height          =   375
         Index           =   1
         Left            =   4980
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Option1"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton Accept 
         Caption         =   "Accept"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   1500
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox CourseTitleBox 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox ClassNameBox 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox InstrNameBox 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         HelpContextID   =   1
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         WhatsThisHelpID =   1
         Width           =   2175
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1920
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   327680
         Appearance      =   1
         MouseIcon       =   "Introduction_PORTUGAL.frx":0000
      End
      Begin VB.Label Label2 
         Caption         =   "Please Wait"
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   1620
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Split Mode"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Normal Mode"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Course Title:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Class Name:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Instructor Name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   13650
      Visible         =   0   'False
      Width           =   18690
      _ExtentX        =   32967
      _ExtentY        =   609
      SimpleText      =   "Hello"
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Unknown"
            TextSave        =   "Unknown"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2884
            Text            =   "Simulator Not Started "
            TextSave        =   "Simulator Not Started "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "Introduction_PORTUGAL.frx":001C
   End
   Begin Fximg32Lib.FXImage FXImage1 
      Height          =   3735
      Left            =   1200
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
      _Version        =   262144
      _ExtentX        =   6588
      _ExtentY        =   6588
      _StockProps     =   64
      BOuterStyle     =   1
      ShowDithered    =   1
      Picture         =   "Introduction_PORTUGAL.frx":0038
      AutoSize        =   2
      MEffect         =   34
      MDelay          =   100
      MGrain          =   10
      TEffect         =   0
      TForeColor      =   0
      TDelay          =   0
      TGrain          =   5
      TBackColor      =   0
      TStyle          =   0
      BOuterColor1    =   8421504
      BOuterWidth     =   3
      Multitask       =   0   'False
      DMode           =   0
      DEffect1        =   5
      DEffect2        =   0
      DEffect3        =   0
      TEnabled        =   0   'False
      BOuterColor2    =   16777215
      Transparent     =   0   'False
      TransparentColor=   0
      BorderWidth     =   3
      BorderColor     =   16711680
      BInnerStyle     =   0
      BInnerWidth     =   1
      BInnerColor1    =   8421504
      BInnerColor2    =   16777215
      Notify          =   0   'False
      Wand            =   0   'False
      WandColor       =   16777215
      WandWidth       =   1
      NotifyDelay     =   0
      CancelMode      =   0
      FXEnabled       =   -1  'True
      TransparentMode =   0
      DecompressMode  =   0
      Palette         =   0
      CancelLoad      =   0   'False
      CancelRemove    =   0   'False
      ShowFlipped     =   0   'False
      ShowReverse     =   0   'False
      ShowRotated     =   0
      HFileName       =   ""
      TGStyle         =   0
      TGColor1        =   16777215
      TGColor2        =   0
      TGSteps         =   16
      Capture         =   0
      TOnly           =   0   'False
      CropX           =   0
      CropY           =   0
      ShowHourglass   =   0   'False
      Pointer         =   0
      PointerMode     =   0
      GIFPassword     =   ""
      FileName        =   ""
      Preview         =   0   'False
      PrinterBanding  =   0   'False
      BackColor       =   16777215
      Update          =   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   13080
      X2              =   13200
      Y1              =   4800
      Y2              =   4680
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Marinha de Comercio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6360
      TabIndex        =   20
      Top             =   4560
      Width           =   8535
   End
   Begin VB.Label Label5 
      Caption         =   "Marine Engine Simulator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   19
      Top             =   11880
      Width           =   8415
   End
   Begin VB.Label Label4 
      Caption         =   "Haven Automation Limited LSS-3M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   18
      Top             =   11040
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   1560
      Top             =   11040
      Width           =   855
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   2280
      Top             =   11640
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   1440
      Top             =   10920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Escola de Pesca e da"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7080
      TabIndex        =   16
      Top             =   2640
      Width           =   6615
   End
End
Attribute VB_Name = "frm_intro_PORTUGAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ReadIniFile()
Dim HOST_string As String
Dim MaxNumberPC_String As String
Dim In_a_group_string As String
Dim Off_Line_string As String


Open "c:\Portugal V1.3\SIM.INI" For Input As #1

'Get the name of THIS PC
Input #1, PC_name
This_PC_number = Val(Mid$(PC_name, 3, 2))

'Is this the host PC ?
Input #1, HOST_string
If HOST_string = "HOST" Then
    I_AM_HOST = True
   Debug.Print host_name & " I_am_host = true"
Else
    I_AM_HOST = False
   Debug.Print host_name & " I_am_host = false"
End If


'Get the Max number of PCs allowed
Input #1, MaxNumberPC_String
Number_of_PCs = Val(MaxNumberPC_String)
Debug.Print "Number of PCs = " & Number_of_PCs

'Is On Line Operation Possible ?
Input #1, In_a_group_string
If InStr(1, In_a_group_string, "yes", 1) Then
    OnLineAllowed = True
    OldOnLineStatus = 1 ' put first PC on-line by default
Else
    OnLineAllowed = False
End If
Debug.Print "OnLineAllowed = " & OnLineAllowed

'Is Off Line Operation Possible ?
Input #1, Off_Line_string
If InStr(1, Off_Line_string, "yes", 1) Then
    OffLineAllowed = True
Else
    OffLineAllowed = False
End If
Debug.Print "OffLineAllowed  = " & OffLineAllowed

' 10/4/98
Close #1

End Sub


Private Sub Accept_Click(Index As Integer)
 split_mode = Option1(1).Value
      
      If split_mode Then
         Debug.Print "Split mode chosen"
      Else
        Debug.Print "Split mode NOT chosen"
      End If
'  1/4/97
'''Timer1.Enabled = False   '   don't need the timer any more

Select Case Index

Case 0                      'Accept Button
'   Set operation mode flag
    If Option1(1).Value = True Then
        'Op_Mode = MODE_SINGLE
        SimConfiguration = "SPLIT"
        HostMessage = "SPLIT"
' 25/3/97
       NORMAL_mode = False
       
    Else
        'Op_Mode = MODE_NORMAL
        SimConfiguration = "NORMAL"
        HostMessage = "NORMAL"
  ' 25/3/97
       NORMAL_mode = True
       
    End If
    
    InstructorName = InstrNameBox.Text
    ClassName = ClassNameBox.Text
    CourseName = CourseTitleBox.Text

    
    'frm_Host.Visible = False
    ''''' RWJ 24/4/98  Load frm_Host
    frm_Host.Visible = True
    frm_Host.WindowState = 2

    
    '24/4/98 next 2 lines worked ok but messy screen
    'frm_Host.Visible = True
    'frm_Host.WindowState = 2
    
' 1/4/97
'     Unload Me
    frm_intro.Visible = False

Case 1
    '   Exit From Program
    Unload Me
    End

Case 2

    ' Show some help

End Select

AcceptButtonPressed = True
End Sub

Private Sub cmd_op_mode_Click(Index As Integer)
Select Case Index
    Case 0
    Op_Mode = MODE_NORMAL
    HostMessage = "NORMAL"
    Case 1
    Op_Mode = MODE_SINGLE
    HostMessage = "SPLIT"
End Select

'frm_exec.Host_sock_Rx.SendLen = Len(HostMessage)
'frm_exec.Host_sock_Rx.SendData = HostMessage

Instructor_Info.Show
Unload Me

End Sub


Private Sub Check1_Click(Index As Integer)


    
End Sub





Private Sub Form_Activate()
 FXImage1.Visible = True
End Sub

Private Sub form_Load()
Dim exec_string As String
Dim ans As Integer

' 21/1/98   SET THE MEDIUM SPEED MODEL ( True for PORTUGAL / False for VIGO&FERROL )
'Portugal_flag = False          ' SETS THE MEDIUM SPEED MODEL FOR VIGO/FERROL
Portugal_flag = True            ' SETS THE MEDIUM SPEED MODEL FOR PORTUGAL
Call set_PORTUGAL(Portugal_flag)

' 26/3/97   Default to HOST_A
HOST = HOST_A


If test_flag Then
    Timer1.Enabled = False
    frm_intro.Accept(0).Visible = True
End If

'read ini file
Debug.Print "in intro, about to read IniFile"

Call ReadIniFile

Debug.Print "Read IniFile"

'28/4/98
PC_IN_CONTROL = This_PC_number




'FXImage1.Picture = LoadPicture("c:\portugal V1.3\images\t42.bmp")

'FXImage1.Picture = LoadPicture("c:\portugal V1.3\images\LSS-2M.jpg")
'If Not I_AM_HOST Then
   ' Timer2.Enabled = False
'End If

Timer2.Enabled = Not I_AM_HOST
 '3/7/97 initialize progress bar
  ProgressBar1.Value = ProgressBar1.Min
  IntroTimerCounter = 0

Debug.Print "End of form_load"


End Sub


Private Sub FXImage1_Click()
Dim i As Integer
i = 2


End Sub

Private Sub Option1_Click(Index As Integer)
Dim ans As String

' 8/4/98
'If PCP_Executive = False Then
If I_AM_HOST = False Then
   ' ans = MsgBox(" Selection of Split/Normal mode MUST be made in the PCP room", 0, "** Warning **")
     ans = MsgBox(" Selection of Split/Normal mode MUST be made by the HOST ", 0, "** Warning **")
Else
      
      
 '    Select Case Index
 '   Case 0
 '       Option1(0).Value = True
 '       Option1(1).Value = False
 '
 '   Case 1
 '       Option1(1).Value = True
 '       Option1(0).Value = False
 '
 '   End Select
 '   If test_flag Then
 '      SPLITMODE = Option1(1).Value = True
 '   End If
    
      split_mode = Option1(1).Value
      
      If split_mode Then
         Debug.Print "Split mode chosen"
      Else
        Debug.Print "Split mode NOT chosen"
      End If
     
End If

  Debug.Print "index = " & Index

    
End Sub


Private Sub Picture1_Change()

End Sub

Private Sub Picture1_Click()

End Sub




Private Sub Timer1_Timer()
Dim IncrementVal As Integer

'Debug.Print "Entered Timer on Intro Form"

' 1/4/97


If I_AM_HOST Then

   frm_intro.Accept(0).Visible = True
   
'   Debug.Print "In timer1 , I_AM_HOST=true"
   
Else


  If Not RxConnected Then
       Debug.Print "In timer1 , RxConnected=false"
            
    '------------------------------------------------------

    'Function is to initialize tcpip Tx and connect to host while Intro Form is displayed
    'Wait 2 seconds after form is loaded before initializing Tx socket

    FirstTcpipInitCounter = FirstTcpipInitCounter + 1
    Select Case FirstTcpipInitCounter
    
    Case 2
        If Not FirstTcpipInit Then
            FirstTcpipInit = True
            'load the exec and tcpip forms but do not show them
            frm_Host.Visible = False
            frm_Host.WindowState = 1
            Load frm_Host
            Debug.Print "loaded host form"

            'initialize tcpip Tx and set up the Receive Sockets
        End If
    
    Case 3
    
       Debug.Print "calling tcpip initialize from timer1,  time = " & Time
       Call frm_Host.tcpip_initialise
       
    Case 5
    
   '     Debug.Print " intro timer 1 a time = " & Time
      
   '
   '    If frm_Host.Client_sock_Tx.Connected Then
    '       MessageString = "CONNECT_PC"
   '        Call frm_Host.SendMessage(MessageString)
    '       Debug.Print "Sending CONNECT_PC Message" & MessageString
   '    Else
    '       Debug.Print " In timer 1 Can't send CONNECT_PC Message,Client_sock_Tx.Connected is false  time = " & Time

   '     End If

      
    
     Case 10, 15, 20
   
          
             Debug.Print "in timer 1,FirstTcpipInitCounter = 3  " & Time

            'If the Tx socket is connected OK, Send a command to connect to my Rx sockets
            ' 16/3/98 If frm_Host.Client_sock_Tx.Connected Then
            If frm_Host.Client_sock_Tx.Connected And Not RxConnected Then

      
                 MessageString = "RETRY_ME"
         
                 Call frm_Host.SendMessage(MessageString)
             
                 Debug.Print "Sending Message" & MessageString
            Else
                Debug.Print "Host_sock_Rx not connected - can't send"
            End If
            
    End Select
        

    'End If

        'Check evey second if host has connected ok, then enable Accept button
        If FirstTcpipInitCounter > 3 Then CheckRxConnect (RxConnected)
        'count time to connection in case host is not running
        CommCounter = CommCounter + 1
        If CommCounter > 40 Then
            CommCounter = 0
            HostAvailable = False
            AcceptButtonPressed = True
            Call Accept_Click(0)
        Else
            ' HostAvailable = True
        End If
    
  Else         ' RxConnected is true

         frm_intro.Accept(0).Visible = True
  End If
End If 'main





End Sub

Sub CheckRxConnect(xxx As Byte)

RxConnected = True

'check if any of the Rx sockets are not connected
For i = 0 To Max_Client_Sock_Rx
      Debug.Print "in checkRxConnect, Client_Sock_Rx(" & i & ").Connected = " & frm_Host.Client_Sock_Rx(i).Connected
      If Not frm_Host.Client_Sock_Rx(i).Connected Then
          RxConnected = False
      End If
Next

End Sub

Private Sub Timer2_Timer()

'3/7/97
Dim IncrementVal As Integer

'GoTo skip_it

' Only if a client

If Not I_AM_HOST Then

  If Not RxConnected Then
    
        Label2.Visible = True
        Label2.Caption = "Please Wait "
        IntroTimerCounter = IntroTimerCounter + 1 'not used for now
        ProgressBar1.Visible = True
        IncrementVal = (ProgressBar1.Max - ProgressBar1.Min) / 25   '25 seconds to fill bar is approx time to connect to host
       
        If ProgressBar1.Value + IncrementVal <= ProgressBar1.Max Then
            ProgressBar1.Value = ProgressBar1.Value + IncrementVal
        End If
  Else
        Label2.Visible = False
        ProgressBar1.Visible = False
  End If
 
 
 'end of 3/7/97 edit
 
 intro_counter = intro_counter + 1
 If intro_counter = 100 Then
  intro_counter = 0
 End If
 
End If

skip_it:
   
End Sub


