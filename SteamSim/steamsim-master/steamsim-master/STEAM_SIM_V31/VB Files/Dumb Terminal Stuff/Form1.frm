VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Dumb Terminal for COM2"
   ClientHeight    =   11850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   11850
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   435
      Left            =   540
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   675
      Left            =   4260
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   10800
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      Height          =   675
      Left            =   2460
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   10800
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Left            =   900
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   10800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   10995
      Left            =   420
      TabIndex        =   0
      Top             =   720
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   19394
      _Version        =   393217
      BackColor       =   12648447
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8520
      Top             =   11280
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   120
      Top             =   11040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim com2open As Boolean
Dim com1open As Boolean
Dim TheLastWord As String


Dim OldStr As String





Private Sub Command1_Click()
    OldStr = "MANUAL Reset Dumb Terminal Window"
    RichTextBox1.Text = OldStr
End Sub

Private Sub Form_Load()



On Error GoTo ErrH2
If Not com2open Then
            MSComm2.PortOpen = True
            com2open = True
End If

               
Exit Sub

ErrH2:
 Debug.Print "Cannot open 2, MSComm2.PortOpen = " & MSComm2.PortOpen
 Debug.Print "Com2Open = " & com2open
 Exit Sub
 
 
End Sub

Private Sub Timer1_Timer()
Dim CharStrChar As String * 1
Dim CharStrWord As String

CharStrWord = ""
'Text2.Text = ""

If MSComm2.InBufferCount > 0 Then
    'Text1.Text = "Count = " & MSComm2.InBufferCount
    'Text3.Text = "Len = " & MSComm2.InputLen
    Do While MSComm2.InBufferCount > 0
        CharStrWord = MSComm2.Input
        Text2.Text = CharStrWord
        'CharStrWord = CharStrWord + CharStrChar
        'If CharStrChar = vbCr Then
            'OldStr = CharStrWord & vbCrLf & OldStr
            If InStr(1, CharStrWord, Chr(10), 1) > 0 Or InStr(1, CharStrWord, Chr(12), 1) > 0 Then
                CharStrWord = TheLastWord + CharStrWord
                TheLastWord = ""
                If Len(OldStr) > 10000 Then
                    OldStr = "AUTO Reset Dumb Terminal Window"
                End If
                
                OldStr = CharStrWord & OldStr
                RichTextBox1.Text = OldStr
            Else
                TheLastWord = CharStrWord 'wait till word is completed next time
            End If
            
        'End If
Loop
    
    
    
End If

Do While MSComm2.InBufferCount > 0
    CharStrChar = MSComm2.Input
    CharStrWord = CharStrWord + CharStrChar
    If CharStrChar = vbCr Then
        OldStr = OldStr + CharStrWord
        RichTextBox1.Text = OldStr
    End If
Loop

End Sub


