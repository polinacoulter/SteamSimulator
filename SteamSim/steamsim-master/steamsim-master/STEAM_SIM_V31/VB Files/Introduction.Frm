VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Object = "{918E6E43-F23A-11D0-901E-0020AF7543C2}#5.0#0"; "fximg50g.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Begin VB.Form frm_intro 
   Caption         =   "LSS-3M Marine Diesel Engine Simulator (Ver    )"
   ClientHeight    =   12765
   ClientLeft      =   30
   ClientTop       =   1650
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12765
   ScaleWidth      =   19080
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   0
      Left            =   0
      TabIndex        =   0
      Top             =   12765
      Visible         =   0   'False
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   0
      SimpleText      =   "Hello"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Unknown"
            TextSave        =   "Unknown"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin FXIMG50GLib.FXImage FXImage1 
      Height          =   12975
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Visible         =   0   'False
      Width           =   18690
      _Version        =   327680
      _ExtentX        =   32967
      _ExtentY        =   22886
      _StockProps     =   65
      BackColor       =   16777215
      Persistence     =   -1  'True
      _StdProps       =   35635726
      ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
      ErrCode         =   1705784705
      ErrInfo         =   615308417
      BOuterStyle     =   1
      Picture         =   "Introduction.frx":0000
      AutoSize        =   2
      MDelay          =   100
      MGrain          =   10
      BOuterWidth     =   3
      DEffect1        =   5
      BorderWidth     =   3
      BorderColor     =   16711680
      FXEnabled       =   -1  'True
      TGColor1        =   16777215
      BackColor       =   16777215
      Update          =   -1  'True
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   960
         TabIndex        =   28
         Top             =   9960
         Visible         =   0   'False
         Width           =   2415
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Visible/Invisible"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   30
            Top             =   300
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            Caption         =   "Unload/Load"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Normal Mode"
            Height          =   255
            Index           =   4
            Left            =   -1980
            TabIndex        =   31
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin SysInfoLib.SysInfo SysInfo1 
         Left            =   660
         Top             =   6180
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Frame Operational_mode_frame 
         Caption         =   "Operational Mode"
         Height          =   975
         Left            =   7860
         TabIndex        =   17
         Top             =   10080
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label1 
            Caption         =   "Normal Mode"
            Height          =   255
            Index           =   1
            Left            =   -1980
            TabIndex        =   18
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Simulator Instructor Information"
         Height          =   2535
         Left            =   9900
         TabIndex        =   6
         Top             =   8580
         Width           =   5595
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Normal Mode"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   540
            TabIndex        =   33
            Top             =   1800
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "      Split Mode"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   32
            Top             =   2100
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Work with Group"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "Medium Speed"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   20
            Top             =   2280
            Width           =   1635
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "Slow Speed"
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   19
            Top             =   2280
            Width           =   1635
         End
         Begin VB.TextBox InstrNameBox 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            HelpContextID   =   1
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   11
            Top             =   360
            WhatsThisHelpID =   1
            Width           =   2175
         End
         Begin VB.TextBox ClassNameBox 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   10
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox CourseTitleBox 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CommandButton Accept 
            Caption         =   "Accept"
            Height          =   375
            Index           =   0
            Left            =   4140
            TabIndex        =   8
            Top             =   1020
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Accept 
            Caption         =   "Exit"
            Height          =   375
            Index           =   1
            Left            =   4140
            TabIndex        =   7
            Top             =   420
            Width           =   975
         End
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   4080
            Top             =   1380
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   4560
            Top             =   1380
         End
         Begin ComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   2280
            TabIndex        =   12
            Top             =   2040
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   450
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Instructor Name:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Class Name:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Course Title:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Please Wait"
            Height          =   195
            Left            =   2280
            TabIndex        =   13
            Top             =   1680
            Width           =   2235
         End
      End
      Begin FXIMG50GLib.FXImage Portugal_logo 
         Height          =   3435
         Left            =   7470
         TabIndex        =   23
         Top             =   12330
         Visible         =   0   'False
         Width           =   3435
         _Version        =   327680
         _ExtentX        =   6059
         _ExtentY        =   6059
         _StockProps     =   65
         BackColor       =   16777215
         Persistence     =   -1  'True
         _StdProps       =   3600619
         ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
         ErrCode         =   1705784705
         ErrInfo         =   615308417
         BOuterStyle     =   1
         Picture         =   "Introduction.frx":2EECF
         AutoSize        =   2
         MEffect         =   34
         MDelay          =   100
         MGrain          =   10
         BOuterWidth     =   3
         DEffect1        =   5
         BorderWidth     =   3
         BorderColor     =   16711680
         FXEnabled       =   -1  'True
         TGColor1        =   16777215
         BackColor       =   16777215
         Update          =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remember to check the ""Work With Group"" box after testing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2685
         Left            =   2700
         TabIndex        =   34
         Top             =   8970
         Visible         =   0   'False
         Width           =   14505
      End
      Begin VB.Image Image2 
         Height          =   1545
         Left            =   60
         Picture         =   "Introduction.frx":4B2D1
         Stretch         =   -1  'True
         Top             =   60
         Width           =   3825
      End
      Begin VB.Image Image1 
         Height          =   1725
         Left            =   1620
         Picture         =   "Introduction.frx":6AC17
         Top             =   12120
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label DEMO_label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DEMONSTRATION VERSION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   4920
         TabIndex        =   27
         Top             =   2100
         Width           =   7455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NB.See Form Load to set Operational Mode.Visible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   2880
         TabIndex        =   26
         Top             =   11100
         Visible         =   0   'False
         Width           =   6915
      End
      Begin VB.Label College_label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Marinha de Comercio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   11820
         TabIndex        =   25
         Top             =   1230
         UseMnemonic     =   0   'False
         Width           =   7635
      End
      Begin VB.Label College_label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escola de Pesca e da"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   11670
         TabIndex        =   24
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   7935
      End
      Begin VB.Label xxxxxx 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NB.See Form Load to alter the above two lines"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   9420
         TabIndex        =   21
         Top             =   1860
         Visible         =   0   'False
         Width           =   7155
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Propulsion Plant Simulator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   735
         Left            =   4290
         TabIndex        =   5
         Top             =   600
         Width           =   7635
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "L3 - MPRI Ship Analytics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Left            =   4320
         TabIndex        =   4
         Top             =   4080
         Visible         =   0   'False
         Width           =   6255
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Visible         =   0   'False
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
      Left            =   3300
      TabIndex        =   3
      Top             =   10560
      Visible         =   0   'False
      Width           =   8535
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
      Left            =   3120
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "frm_intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SetControls()
'    Dim x As Integer
'    Dim y As Integer

 '   x = Screen.TwipsPerPixelX
 '   y = Screen.TwipsPerPixelY
    
 '   Select Case x
    
  '      Case 15, 15
  '          ' Resize and move controls.
  '          txtName.Height = 200
  '          txtName.Width = 500
  '          txtName.Move 200, 200
  '      ' Add code for other resolutions.
   '     …
  '  End Select
    
    
End Sub

Sub ReadIniFile()
Dim HOST_string As String
Dim MaxNumberPC_String As String
Dim In_a_group_string As String
Dim Off_Line_string As String
'11/10/99
Dim InstructorString As String
Dim PTTString As String
Dim IAGString As String


Dim resp As Integer
Dim ErrType As String
Dim FileNo As Integer

On Error GoTo ErrH


FileNo = FreeFile

ErrType = "Opening File"
'11/10/99 change for steam sim  Open "c:\HAVEN\SIM.INI" For Input As #FileNo
Open "c:\STEAM_SIM\SIM.INI" For Input As #FileNo

'Get the name of THIS PC
Input #FileNo, PC_name

' 22/11/99  OK for PCxx but not for HASPxx
'This_PC_number = Val(Mid$(PC_name, 3, 2))
This_PC_number = Val(Mid$(PC_name, Len(PC_name) - 1, 2))


' === 14/03/01 ===

This_PC_Name = PC_name

' === 29/06/01 ONLY FOR TESTING at HA ===
'If This_PC_Name = "PC00" Then
'    This_PC_Name = "HASP01"
'Else
'    This_PC_Name = "HASP00"
'End If




'Is this the host PC ?
ErrType = "Reading HOST String"
Input #FileNo, HOST_string
If HOST_string = "HOST" Then
    I_AM_HOST = True
   Debug.Print host_name & " I_am_host = true"
Else
    I_AM_HOST = False
   Debug.Print host_name & " I_am_host = false"
End If


'Get the Max number of PCs allowed
ErrType = "Reading Number of PCs"

Input #FileNo, MaxNumberPC_String
Number_of_PCs = Val(MaxNumberPC_String)
Debug.Print "Number of PCs = " & Number_of_PCs

'Is On Line Operation Possible ?
ErrType = "Reading On Line Configuration"

Input #FileNo, In_a_group_string
If InStr(1, In_a_group_string, "yes", 1) Then
    OnLineAllowed = True
    OldOnLineStatus = 1 ' put first PC on-line by default
Else
    OnLineAllowed = False
End If
Debug.Print "OnLineAllowed = " & OnLineAllowed

'Is Off Line Operation Possible ?
ErrType = "Reading Off Line Configuration"

Input #FileNo, Off_Line_string
If InStr(1, Off_Line_string, "yes", 1) Then
    OffLineAllowed = True
Else
    OffLineAllowed = False
End If
Debug.Print "OffLineAllowed  = " & OffLineAllowed

Input #FileNo, InstructorString
If InStr(1, InstructorString, "yes", 1) Then
    I_AM_INSTRUCTOR = True
Else
    I_AM_INSTRUCTOR = False
End If
Debug.Print "Instructor ?  = " & I_AM_INSTRUCTOR


Input #FileNo, PTTString
If InStr(1, PTTString, "yes", 1) Then
    PTT = True
    
    ' === 23/01/02 ===
    PTT_mode = True
    
Else
    PTT = False
    
    ' === 23/01/02 ===
    PTT_mode = False
   
End If

Debug.Print "Part Task Trainer ?  = " & PTT

'30/10/99
Input #FileNo, IAGString
If InStr(1, IAGString, "yes", 1) Then
    In_a_group = True
Else
    In_a_group = False
End If
Debug.Print "In a Group  = " & In_a_group

' 10/4/98
Close #FileNo

Exit Sub

ErrH:
resp = MsgBox("There was an error reading C:\Steam_SIm\Sim.ini" & Chr(10) & _
       "You must ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
       Stop
       

End Sub


Private Sub Accept_Click(Index As Integer)


' 16/12/98 ADDED here so that the CORRECT com port will be opened
If Option2(0).Value = True Then
        medium_speed = False
Else
        medium_speed = True
End If
    
' 5/7/99  STEAM_SIM.DLL was copied from SLOW speed model so....


' === 16/05/01 ===
'If TARGET_this_sim = CALIFORNIA Then

        medium_speed = False
        
'End If


' 9/12/98 MOVED from FORM HOST EXEC TIMER
' ( we do not now have a reset  facility and the .EXE version CRASHES on PC00 when calling
' D_MODEL from HOST.FRM FORM LOAD )
' It is VERY important for the COM port to be OPEN BEFORE calling D_MODEL

' 17/12/98 changed to If medium_speed...
If medium_speed Then
'If PORTUGAL Then            ' PORTUGAL IS medium_speed

'        If Not ComOpen Then
'            Debug.Print "Opening com1 , medium speed"

            Call open_com(1)
            
'            ComOpen = True

'        End If
Else
'        If Not ComOpen_LSS Then
'            Debug.Print "Opening com1 , slow speed"

            Call open_com_LSS(1)
            
'            ComOpen_LSS = True
'        End If
End If



If Check1.Value = 1 Then
    In_a_group = True
Else
    In_a_group = False
End If

If I_AM_HOST Or Not In_a_group Then ' Allow user to set Med/Slow speed

    If Option2(0).Value = True Then
        medium_speed = False
    Else
        medium_speed = True
    End If
    
    
    ' 5/7/99  STEAM_SIM.DLL was copied from SLOW speed model so....
    
   ' === 16/05/01 ===
   'If TARGET_this_sim = CALIFORNIA Then
        medium_speed = False
   'End If

    
End If
    
    split_mode = Option1(1).Value
      
    If split_mode Then
         Debug.Print "Split mode chosen"
    Else
        Debug.Print "Split mode NOT chosen"
    End If

    Select Case Index

    Case 0                      'Accept Button
    
         ' 2/12/98
            Me.MousePointer = ccHourglass
    
        '   Set operation mode flag
        If Option1(1).Value = True Then
            SimConfiguration = "SPLIT"
            HostMessage = "SPLIT"
            NORMAL_mode = False
        Else
            SimConfiguration = "NORMAL"
            HostMessage = "NORMAL"
            NORMAL_mode = True
        End If
        
  
        InstructorName = InstrNameBox.text
        ClassName = ClassNameBox.text
        CourseName = CourseTitleBox.text
        

        ' ===== 8/12/98  =====
        'frm_Host.Visible = True
        'frm_Host.WindowState = 2
        ' ===================
         
         
   
        ' 2/12/98
        ' 11/12/98 Moved...
        'Me.MousePointer = ccDefault
        

         ' 8/12/98
         ' 11/12/98 REMOVED....
        'If Option3(0).Value = True Then
        '       frm_intro.Visible = False
        'Else
        '        Unload Me
        'End If
        
        
             ' ===== 8/12/98  =====
             
              frm_Host.Visible = True    ' CAUSES .EXE Version to jump out to WINNT
                                          ' ( ONLY on PC00 )
                                          ' NOW CURED Problem was caused by com_open for fortran , NOW OK
                                        
              ' CAUSED A HANG-UP even when running under VB5
              ' This was because of a "glitched" Tx/Rx board caused by turning the MECD power OFF
              ' Switched EVERYTHING OFF and back on and now VB5 is OK
             'frm_Host.Show               ' ALSO CAUSES .EXE Version to jump out to WINNT
                                          ' NOW CURED Problem was caused by com_open for fortran , NOW OK
                                          ' ( ONLY on PC00 )
           
         
            frm_Host.WindowState = 2
            
        ' 11/12/98
        Me.MousePointer = ccDefault

            
             ' 11/12/98 There is no RESET now so we do not need Intro form...
                Unload Me

        ' ===================

        

    Case 1
        'Exit From Program
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


Private Sub Form_Activate()

 FXImage1.Visible = True
 
End Sub

Private Sub Form_Load()
Dim exec_string As String
Dim ans As Integer



' 3/7/98
Dim width_ratio As Single
Dim height_ratio As Single


' 2/7/98
    Dim X As Integer
    Dim Y As Integer
    
    
'=====  TARGET CONFIGURATION ==========================


' 2/12/98  NOTE see MAIN.BAS for these declarations.....
'Global TARGET_this_sim As Integer
'Global Const NOWHERE = 0      ' USE THIS for a MAXSIM type simulator (NO HARDWARE)
'Global Const THAILAND = 1
'Global Const PORTUGAL = 2
'Global Const VIGO = 3
'Global Const FERROL = 4


                                                                           
' === 10/06/02 ===
'     TARGET_this_sim = CALIFORNIA    ' VERY IMPORTANT TO SET THIS CORRECTLY
               
'     TARGET_this_sim = MARYLAND    ' VERY IMPORTANT TO SET THIS CORRECTLY
                                    
'     TARGET_this_sim = TEXAS    ' VERY IMPORTANT TO SET THIS CORRECTLY
                    


' === 11/12/02 ===
 
' === 03/01/03 ===
    TARGET_this_sim = NOWHERE       ' VERY IMPORTANT TO SET THIS CORRECTLY

' === 14/01/03 ===
'    TARGET_this_sim = CALIFORNIA    ' VERY IMPORTANT TO SET THIS CORRECTLY
     
' === 20/01/03 ===
'    TARGET_this_sim = NOWHERE       ' VERY IMPORTANT TO SET THIS CORRECTLY
     
' === 21/01/03 ===
'    TARGET_this_sim = CALIFORNIA    ' VERY IMPORTANT TO SET THIS CORRECTLY


' === 21/01/03 ===

'    TARGET_this_sim = NOWHERE       ' VERY IMPORTANT TO SET THIS CORRECTLY
' === 28/01/03 ===
    
'    TARGET_this_sim = TEXAS    ' VERY IMPORTANT TO SET THIS CORRECTLY

' === 12/02/03 ===
'    TARGET_this_sim = CALIFORNIA    ' VERY IMPORTANT TO SET THIS CORRECTLY




' === 28/08/03 10:41 for Steam_SimV26_08.exe ===
'     TARGET_this_sim = MARYLAND    ' VERY IMPORTANT TO SET THIS CORRECTLY

' === 28/08/03 10:44 for Steam_SimV26_05.exe ===

' === 11/09/03 ===
'      TARGET_this_sim = CALIFORNIA    ' VERY IMPORTANT TO SET THIS CORRECTLY

' === 11/09/03 ===
TARGET_this_sim = NOWHERE


 ' === 12/12/05 ===
 
 ' ==== 09/06/09 Was TEXAS but now it's CALIFORNIA ===
    'TARGET_this_sim = TEXAS    ' VERY IMPORTANT TO SET THIS CORRECTLY
    TARGET_this_sim = CALIFORNIA    ' VERY IMPORTANT TO SET THIS CORRECTLY

' === 25/02/04 For Ray Gillet SINGLE USER, NO DONGLE , DLL protected(Greece V26.00 G)
' Need to make Super_Pro_Dongle = False to make this for Ray ...

' === 14/12/05 TEXAS is BASBH so use the next line ...===

' === 07/03/07 TEXAS has 5 Pcs with BASBH and 6 with SUPERPRO... ===

' 12/10/09
' So to make Steam_SimV28_BASBH.exe...
Super_Pro_Dongle = False
' So to make Steam_SimV28_SPRO.exe...
'Super_Pro_Dongle = True

' ==================================================================

' === 22/02/06 Here in Swansea... ===
'Super_Pro_Dongle = True



' === 29/01/02 ===
If TARGET_this_sim = CALIFORNIA Then
    CMA_flag = True
Else
    CMA_flag = False
End If

' === 06/03/03 ===
If TARGET_this_sim = TEXAS Then
    Ballast_flag = True
Else
    Ballast_flag = False
End If


' === 25/03/02 ===
' === 10/04/02 Force CMA to the NEW addresses===
'If TARGET_this_sim = CALIFORNIA Then
'    HostAddrPC00 = "1.1.1.0"
'    HostAddrPC01 = "1.1.1.1"
'    HostAddrPC02 = "1.1.1.2"
'    HostAddrPC03 = "1.1.1.3"
'    HostAddrPC04 = "1.1.1.4"
'    HostAddrPC05 = "1.1.1.5"
'Else

    ' === 12/12/05 ===
    ' === 14/12/05 PUT BACK to 192.168 etc for TEXAS ===
    HostAddrPC00 = "192.168.1.100"
    HostAddrPC01 = "192.168.1.1"
    
  ' === 22/02/06 Here in Swansea... ===
  ' HostAddrPC00 = "10.0.0.3"
  ' HostAddrPC01 = "10.0.0.6"
    
' === 07/03/07 SEE BELOW FOR HostAddrPC00 to HostAddrPC15 NOW read from TCPIP_15.INI ===
    HostAddrPC02 = "192.168.1.2"
    HostAddrPC03 = "192.168.1.3"
    HostAddrPC04 = "192.168.1.4"
    HostAddrPC05 = "192.168.1.5"
' === 07/03/07 ==========================================================================
   
    
'End If


'===== 2/12/98 ===================
 

    
'3/9/99 MOVED to frm_HOST.frm
' Dim IO_ret_value As Boolean
'  IO_ret_value = IO_Init(WINCARD_A, Status)
'  IO_ret_value = IO_Init(WINCARD_B, status2)
    
' 3/7/98
' Default to HIGH RES values....
'Me.Width = 17565
'Me.Height = 13410
Me.Width = 17565 * 1.8
Me.Height = 13410 * 1.8
'Me.Width = Screen.Width
'Me.Height = Screen.Height


    X = Screen.TwipsPerPixelX
    Y = Screen.TwipsPerPixelY
    
 ' CAN'T DO THIS ( read only )
 '   Screen.TwipsPerPixelX = 20
 '   Screen.TwipsPerPixelY = 20
    
   Debug.Print "Screen.TwipsPerPixelX=" & X
   Debug.Print "Screen.TwipsPerPixelY=" & Y
    Debug.Print "me.width=" & Me.Width
    Debug.Print "me.height=" & Me.Height
    
   Debug.Print "SysInfo1.WorkAreaWidth = " & SysInfo1.WorkAreaWidth
  Debug.Print "SysInfo1.WorkAreaheight = " & SysInfo1.WorkAreaHeight
   
 '''''  GoTo skip_it:
 
 ' FUDGE FACTORS of 0.6  for now.....
     width_ratio = SysInfo1.WorkAreaWidth * 0.6 * 1.8 / Me.Width
    Debug.Print "width_ratio=" & width_ratio
    
     height_ratio = SysInfo1.WorkAreaHeight * 0.6 * 1.8 / Me.Height
     Debug.Print "height_ratio=" & height_ratio
  
''' 2/7/98
'SysInfo1_DisplayChanged
If Me.Width > SysInfo1.WorkAreaWidth Then

     
     Me.Left = SysInfo1.WorkAreaLeft
     Me.Width = SysInfo1.WorkAreaWidth
     FXImage1.Left = SysInfo1.WorkAreaLeft
     FXImage1.Width = SysInfo1.WorkAreaWidth
     
       Frame1.Left = SysInfo1.WorkAreaLeft + 0.6 * SysInfo1.WorkAreaWidth
       
     ' 2/12/98
     ' Operational_mode_frame.Left = Operational_mode_frame.Left * width_ratio
       Operational_mode_frame.Left = SysInfo1.WorkAreaLeft + 0.4 * SysInfo1.WorkAreaWidth
      
 '     Frame1.Width = Frame1.Width * width_ratio
     
    Debug.Print "NEW me.width=" & Me.Width
    Debug.Print "Frame1.Left=" & Frame1.Left
   
End If
If Me.Height > SysInfo1.WorkAreaHeight Then

     

     Me.Top = SysInfo1.WorkAreaTop
     Me.Height = SysInfo1.WorkAreaHeight
     FXImage1.Top = SysInfo1.WorkAreaTop
     FXImage1.Height = SysInfo1.WorkAreaHeight
     
        ' 1/9/99
      ' Frame1.Top = SysInfo1.WorkAreaTop + 0.8 * SysInfo1.WorkAreaHeight
    
      
      ' === 20/12/01 ===
      
     ' === 31/05/02 ===
     ' If TARGET_this_sim = CALIFORNIA Or TARGET_this_sim = MARYLAND Then
       If TARGET_this_sim = CALIFORNIA Or TARGET_this_sim = MARYLAND Or TARGET_this_sim = TEXAS Then
        Frame1.Top = SysInfo1.WorkAreaTop + 0.7 * SysInfo1.WorkAreaHeight
      Else
        Frame1.Top = SysInfo1.WorkAreaTop + 0.775 * SysInfo1.WorkAreaHeight
      End If
      ' === 10/10/01 end of edir ===
      
      
           ' 2/12/98
      'Operational_mode_frame.Top = (Operational_mode_frame.Top - Operational_mode_frame.Height) * height_ratio
       Operational_mode_frame.Top = SysInfo1.WorkAreaTop + 0.8 * SysInfo1.WorkAreaHeight

 '    Frame1.Height = Frame1.Height * height_ratio

    Debug.Print "NEW me.height=" & Me.Height
    Debug.Print "Frame1.top=" & Frame1.Top

End If
   
skip_it:
   
'=== 27/5/98 For LSS only ======
' These convert nRMA to the approriate PC_CONTROL numbers
PC_control_LSS(1) = 1
PC_control_LSS(2) = 3
PC_control_LSS(3) = 4
PC_control_LSS(4) = 5
PC_control_LSS(5) = 2
PC_control_LSS(6) = 6
PC_control_LSS(7) = 7
PC_control_LSS(8) = 13
PC_control_LSS(9) = 12
PC_control_LSS(10) = 9
PC_control_LSS(11) = 11
PC_control_LSS(12) = 9
PC_control_LSS(13) = 1   ' T/A NOT USED


' 26/3/97   Default to HOST_A
HOST = HOST_A

'22/11/99
HOST = HostAddrPC00

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



' ============== 07/03/07 ADDED ==============================

Call read_TCPIP_INI_file

HostAddrPC00 = HOST_PC_address

HostAddrPC01 = CLIENT_PC_address(1)

HostAddrPC02 = CLIENT_PC_address(2)

HostAddrPC03 = CLIENT_PC_address(3)

HostAddrPC04 = CLIENT_PC_address(4)

HostAddrPC05 = CLIENT_PC_address(5)


HostAddrPC06 = CLIENT_PC_address(6)
HostAddrPC07 = CLIENT_PC_address(7)
HostAddrPC08 = CLIENT_PC_address(8)
HostAddrPC09 = CLIENT_PC_address(9)
HostAddrPC10 = CLIENT_PC_address(10)

HostAddrPC11 = CLIENT_PC_address(11)
HostAddrPC12 = CLIENT_PC_address(12)
HostAddrPC13 = CLIENT_PC_address(13)
HostAddrPC14 = CLIENT_PC_address(14)
HostAddrPC15 = CLIENT_PC_address(15)

' ============== end of edit =================================

' 12/3/99
'DEMO_flag = True
DEMO_flag = False

If DEMO_flag Then

        Me.Caption = "Steam Propulsion Simulator (Ver 12.00 D)"
    
    
    
   DEMO_label.Visible = True
   
Else
   
    ' 18/1/00
    '    Me.Caption = "Steam Propulsion Simulator (Ver 1.00." & Format(TARGET_this_sim, "00") & ")"
    
    ' 28/4/00
    '  Me.Caption = "Steam Propulsion Simulator (Ver 5.00." & Format(TARGET_this_sim, "00") & ")"
    
    ' 12/5/00
    '  Me.Caption = "Steam Propulsion Simulator (Ver 6.00." & Format(TARGET_this_sim, "00") & ")"
       
    ' 7/8/00
    '   Me.Caption = "Steam Propulsion Simulator (Ver 7.00." & Format(TARGET_this_sim, "00") & ")"
   
    ' 18/9/00
    '   Me.Caption = "Steam Propulsion Simulator (Ver 8.00." & Format(TARGET_this_sim, "00") & ")"
               
     ' 16/9/00  VER 9.2    NB THIS VER WAS NOT DONGLE PROTECTED
     ' Me.Caption = "Steam Propulsion Simulator (Ver 9.00." & Format(TARGET_this_sim, "00") & ")"
      
     ' 16/9/00  VER 9.3 (ON)  NOW HAS DONGLE PROTECTION
      
     ' 1/12/00
     ' Me.Caption = "Steam Propulsion Simulator (Ver 9.3." & Format(TARGET_this_sim, "00") & ")"
     ' 8/12/00
     '  Me.Caption = "Steam Propulsion Simulator (Ver 9.4." & Format(TARGET_this_sim, "00") & ")"
    
     ' === 01/03/01 ===
     '  Me.Caption = "Steam Propulsion Simulator (Ver 10.0." & Format(TARGET_this_sim, "00") & ")"
    
    
     ' === 16/03/01 ===
     '  Me.Caption = "Steam Propulsion Simulator (Ver 11.0." & Format(TARGET_this_sim, "00") & ")"
     
     ' === 16/05/01 For Bill Findlay ===
     '  Me.Caption = "Steam Propulsion Simulator (Ver 12.0." & Format(TARGET_this_sim, "00") & ")"
     
     
     ' === 27/06/01 SQUAREGAUGE.OCX has been REMOVED from the project ===
    '  Me.Caption = "Steam Propulsion Simulator (Ver 12.0." & Format(TARGET_this_sim, "00") & ")  21 MAY 2001"
       
    ' === 11/10/01 ===
    '  Me.Caption = "Steam Propulsion Simulator (Ver 14.0." & Format(TARGET_this_sim, "00") & ")  29 SEPT 2001"
    
    ' === 12/10/01 ===
    '  Me.Caption = "Steam Propulsion Simulator (Ver 14.0." & Format(TARGET_this_sim, "00") & ")  11 OCT 2001"
       
       
    ' === 4/12/01 ===
    '  Me.Caption = "Steam Propulsion Simulator (Ver 14.0." & Format(TARGET_this_sim, "00") & ")  12 OCT 2001"
    
    ' === 20/12/01 ===
    '  Me.Caption = "Steam Propulsion Simulator (Ver 15.0." & Format(TARGET_this_sim, "00") & ")  13 DEC 2001"
       
       
    ' === 24/01/02 ===
    '  Me.Caption = "Steam Propulsion Simulator (Ver 15.0." & Format(TARGET_this_sim, "00") & ")  7 JAN 2002"
 
    ' === 02/04/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 16.0." & Format(TARGET_this_sim, "00") & ")  7 MAR 2002"

   ' === 10/04/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 17.0." & Format(TARGET_this_sim, "00") & ")  2 APR 2002"
   
   ' === 31/05/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 17.0." & Format(TARGET_this_sim, "00") & ")  10 APR 2002"
   
   '   Me.Caption = "Steam Propulsion Simulator (Ver 18.0." & Format(TARGET_this_sim, "00") & ")   7 June 2002"
       
   ' === 29/07/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 18.0." & Format(TARGET_this_sim, "00") & ")  10 June 2002"
       
       
    ' === 01/10/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 18.0." & Format(TARGET_this_sim, "00") & ")  29 JULY 2002"
       
   ' === 11/12/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 19.0." & Format(TARGET_this_sim, "00") & ")  1 OCT 2002"
   
   ' === 16/12/02 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 20.0." & Format(TARGET_this_sim, "00") & ")  11 DEC 2002"
       
   ' === 09/01/03 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 20.2." & Format(TARGET_this_sim, "00") & ")  17 DEC 2002"
       
   ' === 20/01/03 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 21.0." & Format(TARGET_this_sim, "00") & ")   9 JAN 2003"
       
   ' === 21/01/03 ===
   '   Me.Caption = "Steam Propulsion Simulator (Ver 21.0." & Format(TARGET_this_sim, "00") & ")  20 JAN 2003"
       
' === 12/02/03 ===
'      Me.Caption = "Steam Propulsion Simulator (Ver 22.0." & Format(TARGET_this_sim, "00") & ")  21 JAN 2003"

' === 19/02/03 ===
'      Me.Caption = "Steam Propulsion Simulator (Ver 22.0." & Format(TARGET_this_sim, "00") & ")  12 FEB 2003"

'      Me.Caption = "Steam Propulsion Simulator (Ver 22.0." & Format(TARGET_this_sim, "00") & ")  19 FEB 2003"
       
' Me.Caption = "Steam Propulsion Simulator (Ver 23.0." & Format(TARGET_this_sim, "00") & ")  05 MAR 2003"

' === 31/03/03 ===
' Me.Caption = "Steam Propulsion Simulator (Ver 24.0." & Format(TARGET_this_sim, "00") & ")  06 MAR 2003"
  
  
' === 28/08/03 ===
' Me.Caption = "Steam Propulsion Simulator (Ver 25.0." & Format(TARGET_this_sim, "00") & ")   1 APR 2003"


' === 11/09/03 ===
'  Me.Caption = "Steam Propulsion Simulator (Ver 26.0." & Format(TARGET_this_sim, "00") & ")  28 Aug 2003"
If Super_Pro_Dongle Then

    ' --- 10/03/04 ---
  'Me.Caption = "Steam Propulsion Simulator Ver 26.0." & Format(TARGET_this_sim, "00") & ")  11 Sept 2003" & "   (S)"
   Me.Caption = "Steam Propulsion Simulator Ver 26.0." & Format(TARGET_this_sim, "00") & ")  10 Mar 2004" & "   (S) (Sng)"
   
   ' === 22/02/06 ===
   ' === 06/03/07 EXTEND from 5 to 15 PCs NETWORKED  ===
   'Me.Caption = "Steam Propulsion Simulator Ver 27.0." & Format(TARGET_this_sim, "00") & "  22 Feb 2006" & "   (SPro)"
    '=== 04/04/07 Enabled sound on all PCs (not just PC00
   'Me.Caption = "Steam Propulsion Simulator Ver 28.0." & Format(TARGET_this_sim, "00") & "  7 March 2007" & "   (SPro)"
   
   
   '05/06/07 fix scrolling problems Texas A&M reported by Bob Majewski
   'Me.Caption = "Steam Propulsion Simulator Ver 28.1." & Format(TARGET_this_sim, "00") & "  4 April 2007" & "   (SPro)"
  
  ' 09/06/09
  ' Me.Caption = "Steam Propulsion Simulator Ver 28.2." & Format(TARGET_this_sim, "00") & "  5 June 2007" & "   (SPro)"
  
  ' Me.Caption = "Steam Propulsion Simulator Ver 29.0." & Format(TARGET_this_sim, "00") & "  4 Sept 2009" & "   (SPro)"
  ' Me.Caption = "Steam Propulsion Simulator Ver 30.0." & Format(TARGET_this_sim, "00") & "  14 Oct 2009" & "   (SPro)"
   Me.Caption = "Steam Propulsion Simulator Ver 31.0." & Format(TARGET_this_sim, "00") & " 22 Feb 2010" & "   (SPro)"

Else

  ' OLD BASBH dongle...
  'Me.Caption = "Steam Propulsion Simulator Ver 26.0." & Format(TARGET_this_sim, "00") & ")  25 Feb 2004" & "   (BASBH)"
    
  ' === 25/02/04 For Ray Gillet SINGLE USER , NO DONGLE , DLL protected(Greece V26.00 G)
  
  ' === 14/12/05 ===
 ' Me.Caption = "Steam Propulsion Simulator Ver 26.0." & Format(TARGET_this_sim, "00") & "  01 Mar 2004" & "   (G)"
   
   ' === 22/02/06 ===
  ' Me.Caption = "Steam Propulsion Simulator Ver 27.0." & Format(TARGET_this_sim, "00") & "  14 Dec 2005" & "   (BASBH)"
    
   ' === 06/03/07 EXTEND from 5 to 15 PCs NETWORKED  ===
    'Me.Caption = "Steam Propulsion Simulator Ver 27.0." & Format(TARGET_this_sim, "00") & "  22 Feb 2006" & "   (BASBH)"
     '=== 04/04/07 Enabled sound on all PCs (not just PC00
    'Me.Caption = "Steam Propulsion Simulator Ver 28.0." & Format(TARGET_this_sim, "00") & "  7 March 2007" & "   (BASBH)"
    '05/06/07 fix scrolling problems Texas A&M reported by Bob Majewski
    'Me.Caption = "Steam Propulsion Simulator Ver 28.1." & Format(TARGET_this_sim, "00") & "  4 April 2007" & "   (BASBH)"
    
    
    ' 09/06/09
    'Me.Caption = "Steam Propulsion Simulator Ver 28.2." & Format(TARGET_this_sim, "00") & "  5 June 2007" & "   (BASBH)"
   
   ' Me.Caption = "Steam Propulsion Simulator Ver 29.0." & Format(TARGET_this_sim, "00") & "  4 Sept 2009" & "   (BASBH)"
  ' Me.Caption = "Steam Propulsion Simulator Ver 30.0." & Format(TARGET_this_sim, "00") & "  14 Oct 2009" & "   (BASBH)"
   Me.Caption = "Steam Propulsion Simulator Ver 31.0." & Format(TARGET_this_sim, "00") & "  22 Feb 2010" & "   (BASBH)"

End If

    If TARGET_this_sim = CALIFORNIA Then
 
 
    Else
                   
       '   FXImage1.Picture = LoadPicture("c:\HAVEN\Images\BlankBackground.bmp")
         
  
       '  FXImage1.Picture = LoadPicture("c:\Steam_Sim\Images\CMA Logo & HAL.gif")
         
         
         
         ' === 04/02/02 ===
     '   FXImage1.Picture = LoadPicture("c:\HAVEN\Images\Turbine_purple.jpg")
     
      ' === 02/04/02 ===
      '  FXImage1.Picture = LoadPicture("c:\Steam_Sim\Images\Turbine_purple.jpg")
         FXImage1.Picture = LoadPicture("c:\Steam_Sim\Images\Turbine_purple.gif")
         
         
        ' FXImage1.Picture = LoadPicture("c:\Steam_Sim\Images\Turbine_purple.bmp")
          
        '  FXImage1.Picture = LoadPicture("c:\Steam_Sim\Images\BlankBackground.bmp")
   
    End If
    
    
    ' When making an .EXE , make for example Executive V1.7.02  (for PORTUGAL)  etc
 
    DEMO_label.Visible = False
    
End If
'================================


' 10/12/98
HARDWARE_flag = False       ' SET true IF this simulator has REAL HARDWARE
LanguagePrefix = "E"        ' For ENGLISH    10/12/98

     
'  NB There are HARDWARE differences between PORTUGAL and VIGO/FERROL
'  The following lines ONLY affect the MEDIUM SPEED MODEL....
If (TARGET_this_sim = VIGO) Or (TARGET_this_sim = FERROL) Then
   ' VIGO and FERROL get this...
        Portugal_flag = False               ' SETS MEDIUM SPEED (in MODEL.DLL) FOR VIGO/FERROL
        Call set_PORTUGAL(Portugal_flag)
        
    HARDWARE_flag = True    ' 10/12/98
    
    ' 15/12/98 ADDED:
ElseIf (TARGET_this_sim = PORTUGAL) Then

        Portugal_flag = True                ' SETS MEDIUM SPEED (in MODEL.DLL) FOR PORTUGAL
        Call set_PORTUGAL(Portugal_flag)
        
    HARDWARE_flag = True        ' 10/12/98
    
    LanguagePrefix = "P"        ' For PORTUGUESE 10/12/98
    

' === 31/05/02 ===
'ElseIf TARGET_this_sim = CALIFORNIA Or TARGET_this_sim = MARYLAND Then
ElseIf TARGET_this_sim = CALIFORNIA Or TARGET_this_sim = MARYLAND Or TARGET_this_sim = TEXAS Then

       
       Portugal_flag = True                ' SETS MEDIUM SPEED (in MODEL.DLL) FOR PORTUGAL
        Call set_PORTUGAL(Portugal_flag)
        
    
        ' 15/11/99 Put back to Original code....
    '29/10/99 HARDWARE_flag set automatically niw if Applicon Boards initialized OK
    
    '  15/11/99 Put back to Original code....
    If PTT Then
    ' 22/10/99  If Part Task Trainer there is NO Hardware
      HARDWARE_flag = False     '   NO HARDWARE
    Else
      HARDWARE_flag = True      '   HARDWARE
    End If
    
    '28/11/99
    
    ' === 12/11/02 5.00 ADDED: ===
    If I_AM_HOST Then 'allow instructor with hardware to have split mode
         Option1(0).Visible = True
         Option1(1).Visible = True
    End If
    ' =======================
    
    LanguagePrefix = "E"        ' For ENGLISH ( or AMERICAN )
    
Else
 
   ' Everyone else gets the PORTUGAL MEDIUM SPEED MODEL...
        Portugal_flag = True                ' SETS MEDIUM SPEED (in MODEL.DLL) FOR PORTUGAL
        Call set_PORTUGAL(Portugal_flag)
        
   ' 15/12/98 Everyone else has NO HARDWARE
    HARDWARE_flag = False       '    NO HARDWARE
    
    
    ' === 12/11/02 5.00 ADDED: ===
    If I_AM_HOST Then 'allow instructor with hardware to have split mode
         Option1(0).Visible = True
         Option1(1).Visible = True
    End If
    ' =======================
    
    
    LanguagePrefix = "E"         ' For ENGLISH ( or AMERICAN )
    
End If


'============ 9/2/99 ========

' 6/7/99

' 4/11/99 MOVE to MDI Form Load

' 15/11/99
'Dongle_Enable_Flag = Not HARDWARE_flag   ' IF there is NO HARWARE then we need the dongle




'Dongle_Enable_Flag = False


'  HA , PC01 won't read the dongle so for now.....
' 5/3/99 Dongle_Enable_Flag = False
' ( SP32W.DLL is installed on PC01 , so why do we need to instal ALL of the SENTINAL software??? )
' N.B. WE DO NOT NEED ALL of the SENTINAL development software
'  See SENTINEL BOX for instructions on installing Sentinel DRIVERS ( ONLY the required Drivers)


'============ 9/2/99 ========

   
   
   
    Operational_mode_frame.Visible = False    ' ( For NORMAL or SPLIT mode )
    
    Check1.Visible = False                    ' ( Work with Group CheckBox )
    
    Option2(0).Visible = False                ' (   SLOW SPEED option )
    Option2(1).Visible = False                ' ( MEDIUM SPEED option )
                 
    Portugal_logo.Visible = False
    
' === 31/05/02 ===
Ballast_IS_included = False

If TARGET_this_sim = NOWHERE Then
    
    College_label1.Caption = ""
    College_label2.Caption = ""
    
    '15/05/01  Check1.Visible = True           ' ( Work with Group CheckBox )
    
    ' === 25/03/02 ===
    'Check1.Visible = False
    PTT = True
    
    ' === 25/02/04 For Ray Gillet SINGLE USER, NO DONGLE , DLL protected(Greece V26.00 G)
    
    ' --- 10/03/04 ----
    ' For SINGLE USER , you need to comment out next line to make this...
    'Check1.Visible = PTT
    
    
    Check1.Value = 0
    Accept(0).Visible = False
    
    
    '   Option2(0).Visible = True   ' (   SLOW SPEED option )
    '   Option2(1).Visible = True   ' ( MEDIUM SPEED option )
     
ElseIf TARGET_this_sim = CALIFORNIA Then

    College_label1.Caption = "California Maritime Academy"
    College_label2.Caption = "Vallejo"
    
    '30/10/99  Check1.Visible = True     ' ( Work with Group CheckBox )
    Check1.Visible = False
    
    ' 12/10/09
    ' This edit below was original made on the laptop in SLC on 15/09/09 and was
    ' e-mailed on 15/09/09 to Scott Green as Steam_SimV29_00_05_BASBH.zip
        Check1.Visible = PTT

    
    '    Option2(0).Visible = True    ' (   SLOW SPEED option )
    '    Option2(1).Visible = True    ' ( MEDIUM SPEED option )

  ' === 20/12/01 ===
ElseIf TARGET_this_sim = MARYLAND Then

    College_label1.Caption = "S.I.U. Maryland"
    College_label2.Caption = ""
    
   ' === 20/12/01 ===
   ' Check1.Visible = False
    PTT = True
    Check1.Visible = PTT
    Check1.Value = 0
    Accept(0).Visible = False
    
  ' === 31/05/02 ===
ElseIf TARGET_this_sim = TEXAS Then

    College_label1.Caption = " Texas A & M "
    College_label2.Caption = ""
    
   ' === 20/12/01 ===
   ' Check1.Visible = False
    PTT = True
    Check1.Visible = PTT
    Check1.Value = 0
    Accept(0).Visible = False
    
    ' === 31/05/02 ===
    Ballast_IS_included = True

ElseIf TARGET_this_sim = THAILAND Then

    College_label1.Caption = "N.A.S.I.C"
    College_label2.Caption = ""
    
    Check1.Visible = True     ' ( Work with Group CheckBox )
    
        Option2(0).Visible = True    ' (   SLOW SPEED option )
        Option2(1).Visible = True    ' ( MEDIUM SPEED option )
        
ElseIf TARGET_this_sim = PORTUGAL Then

    College_label1.Caption = "Escola de Pesca e da"
    College_label2.Caption = "Marinha de Comercio"
    
    
    If I_AM_HOST Then
        Operational_mode_frame.Visible = True    ' ( For Nomal or Split mode )
    End If
    
    Portugal_logo.Visible = True
    
ElseIf TARGET_this_sim = VIGO Then

        ' NB. CHECK the SPANISH SPELLING.....
    College_label1.Caption = "Escola de Pesca"
    College_label2.Caption = "VIGO"
    
    Check1.Visible = True                   ' ( Work with Group CheckBox )
   
    If I_AM_HOST Then
        Operational_mode_frame.Visible = True    ' ( For Nomal or Split mode )
    End If

        
ElseIf TARGET_this_sim = FERROL Then

        ' NB. CHECK the BASQUE SPELLING.....
    College_label1.Caption = "Esquela de Pesca"
    College_label2.Caption = "FERROL"
    
    Check1.Visible = True                    ' ( Work with Group CheckBox )
    
    If I_AM_HOST Then
        Operational_mode_frame.Visible = True    ' ( For Nomal or Split mode )
    End If

    
End If


' 5/3/99
If DEMO_flag Then
    Check1.Value = 0                    ' DO NOT WORK IN A GROUP
    Check1.Visible = False              ' Work with Group CheckBox )
    Dongle_Enable_Flag = False
End If


'=====  TARGET CONFIGURATION end  =====================



' 2/12/98
If I_AM_HOST Then

    Label2.Visible = False
    ProgressBar1.Visible = False
    
    Label1(0).Caption = "Instructor Name"

Else

    Label1(0).Caption = "Student Name"
    
End If


Timer2.Enabled = Not I_AM_HOST
 '3/7/97 initialize progress bar
  ProgressBar1.Value = ProgressBar1.Min
  IntroTimerCounter = 0

Debug.Print "End of form_load"


' 2/12/98
'If I_AM_HOST Then
'   Check1.Visible = True           'Allow instructor to choose group & sim type (slow/medium)
'    Option2(0).Visible = True
'    Option2(1).Visible = True
'   Label1(0).Caption = "Instructor Name"
'
'Else
'  Check1.Visible = True           'Allow student to choose group but not sim type (slow/medium)
'    Option2(0).Visible = False
'    Option2(1).Visible = False
'   Label1(0).Caption = "Student Name"
    
'End If


End Sub


Private Sub Form_Unload(Cancel As Integer)

' 3/9/99 Moved to frm_HOST
'Dim IO_ret_value As Boolean
'
' IO_ret_value = IO_Exit(WINCARD_A, Status)
' IO_ret_value = IO_Exit(WINCARD_B, Status)
  
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


Private Sub SysInfo1_DisplayChanged()

Debug.Print "SysInfo1_DisplayChanged"

' 2/7/98 RWJ
If Screen.ActiveForm.Width > SysInfo1.WorkAreaWidth Then
        Screen.ActiveForm.Left = SysInfo1.WorkAreaLeft
        Screen.ActiveForm.Width = SysInfo1.WorkAreaWidth
End If
If Screen.ActiveForm.Height > SysInfo1.WorkAreaHeight Then
        Screen.ActiveForm.Top = SysInfo1.WorkAreaTop
        Screen.ActiveForm.Height = SysInfo1.WorkAreaHeight
End If


End Sub

Private Sub SysInfo1_SettingChanged(ByVal Item As Integer)

Debug.Print "SysInfo1_SettingChanged"

End Sub

Private Sub SysInfo1_Click()

End Sub
Sub read_TCPIP_INI_file()
Dim resp As Integer
Dim FileNo As Integer
Dim INstring As String
Dim i As Integer
On Error GoTo ErrH_TCPIP


FileNo = FreeFile


' === 07/03/07 ===
'Open "c:\Eng_Sim\TCPIP_15.INI" For Input As #FileNo
Open "c:\Steam_Sim\TCPIP_15.INI" For Input As #FileNo


Input #FileNo, INstring

' === 14/01/05 Having trouble with NEW address 10.0.0.3 ====
'HOST_PC_address = Mid$(INstring, 1, 15)
HOST_PC_address = Trim(Mid$(INstring, 1, 15))


' === 19/09/01 ===
'For i = 1 To 5

' === 11/01/05 ===
'For i = 1 To 10
For i = 1 To 15


  Input #FileNo, INstring
  
    ' === 11/01/05 Having trouble with NEW address 10.0.0.3 ====
    'CLIENT_PC_address(i) = Mid$(INstring, 1, 15)
     CLIENT_PC_address(i) = Trim(Mid$(INstring, 1, 15))

  
  ' === 18/05/01 ===
  If InStr(1, INstring, "yes", 1) Then
      CLIENT_PC_IS_ON_LINE(i) = True
  Else
      CLIENT_PC_IS_ON_LINE(i) = False
  End If

  
Next i


Close #FileNo

Exit Sub

ErrH_TCPIP:
resp = MsgBox("There was an error reading the " & "from the TCPIP_15.INI file" & Chr(10) & _
       "You must ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
       Stop
       


End Sub
Private Sub Timer1_Timer()
Dim IncrementVal As Integer


If Check1.Value = 0 Then  ' Don't Work with group so...
    
    Accept(0).Visible = True
    
End If


'Debug.Print "Entered Timer on Intro Form"

' 1/4/97


If I_AM_HOST Then

   frm_intro.Accept(0).Visible = True
   
'   Debug.Print "In timer1 , I_AM_HOST=true"
   
Else


' === 25/04/02 ===
'If (TARGET_this_sim = CALIFORNIA) Or ((TARGET_this_sim = MARYLAND) And Check1.Value = 1) Then

' === 31/05/02 ===
' If (TARGET_this_sim = CALIFORNIA) Or ((TARGET_this_sim = MARYLAND) And Check1.Value = 1) _
' Or ((TARGET_this_sim = NOWHERE) And Check1.Value = 1) Then
 If (TARGET_this_sim = CALIFORNIA) Or ((TARGET_this_sim = MARYLAND) And Check1.Value = 1) _
 Or ((TARGET_this_sim = TEXAS) And Check1.Value = 1) Or ((TARGET_this_sim = NOWHERE) And Check1.Value = 1) Then
 
    If Not RxConnected Then
    
        ' === 20/12/01 ===
        frm_intro.Accept(0).Visible = False

  
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
  
  
  ' ===20/12/01 ===
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

' === 20/12/01 ===

' === 25/04/02 ===
'If (TARGET_this_sim = CALIFORNIA) Or ((TARGET_this_sim = MARYLAND) And Check1.Value = 1) Then
 
' === 31/05/02 ===
' If (TARGET_this_sim = CALIFORNIA) Or ((TARGET_this_sim = MARYLAND) And Check1.Value = 1) _
' Or ((TARGET_this_sim = NOWHERE) And Check1.Value = 1) Then
 If (TARGET_this_sim = CALIFORNIA) Or ((TARGET_this_sim = MARYLAND) And Check1.Value = 1) _
  Or ((TARGET_this_sim = TEXAS) And Check1.Value = 1) Or ((TARGET_this_sim = NOWHERE) And Check1.Value = 1) Then


    If Not I_AM_HOST Then

        If Not RxConnected Then
    
            Label2.Visible = True
            
            Label2.Caption = "Please Wait "
            
            IntroTimerCounter = IntroTimerCounter + 1 'not used for now
            
            ProgressBar1.Visible = True
            
            IncrementVal = (ProgressBar1.max - ProgressBar1.Min) / 25   '25 seconds to fill bar is approx time to connect to host
       
            If ProgressBar1.Value + IncrementVal <= ProgressBar1.max Then
                ProgressBar1.Value = ProgressBar1.Value + IncrementVal
            End If
        Else
            Label2.Visible = False
            ProgressBar1.Visible = False
        End If
  
        intro_counter = intro_counter + 1
        If intro_counter = 100 Then
            intro_counter = 0
        End If
 
    End If
    
' == 20/12/01 ===
Else
            Label2.Visible = False
            ProgressBar1.Visible = False
End If

skip_it:
   
End Sub


