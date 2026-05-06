VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Winsock_udp 
   Caption         =   "WinSock UDP"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   2460
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   2100
      Width           =   6255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send  ARRAY"
      Height          =   495
      Left            =   2340
      TabIndex        =   6
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set PC Name"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   180
      Width           =   2355
   End
   Begin VB.CommandButton clear 
      Caption         =   "CLEAR"
      Height          =   555
      Left            =   3300
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   3420
      Width           =   6255
   End
   Begin VB.TextBox PC_name 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Text            =   "PC02"
      Top             =   660
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   180
      TabIndex        =   1
      Top             =   1200
      Width           =   3075
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2940
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   900
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "1.1.1.1"
      RemotePort      =   2001
      LocalPort       =   2000
   End
End
Attribute VB_Name = "Winsock_udp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clear_Click()

Text1.text = ""
'Text3.Text = ""

End Sub

Private Sub Command1_Click()

Winsock1.SendData (Text3.text)

End Sub


Private Sub Command2_Click()
 
If PC_name.text = "PC02" Then
   Winsock1.RemoteHost = "1.1.1.1"
    Winsock1.LocalPort = 2000
   Winsock1.RemotePort = 2001
Else
   Winsock1.RemoteHost = "1.1.1.2"
    Winsock1.LocalPort = 2001
   Winsock1.RemotePort = 2002
End If

End Sub

Private Sub Command3_Click()
Dim REAL_ARR(5) As Single
Dim BYTE_ARR(5) As Byte

'Union , REAL_ARR(1), BYTE_ARR(1)

Dim i As Integer
 
For i = 0 To 5
    REAL_ARR(i) = i + 0.567
    BYTE_ARR(i) = 41 + i
Next i

Text2.text = REAL_ARR(1) & "  " & REAL_ARR(2) & "  " & REAL_ARR(3)
Text4.text = BYTE_ARR(1) & "  " & BYTE_ARR(2) & "  " & BYTE_ARR(3)

Winsock1.SendData (BYTE_ARR)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String * 100
'Dim RX_BYTE_ARR(5) As Byte


Debug.Print "Received " & bytesTotal

Winsock1.GetData strData, vbString
Text1.text = Text1.text & strData

' DOESN'T WORK
'Winsock1.GetData RX_BYTE_ARR(0), vbArray + vbByte
'Text1.Text = RX_BYTE_ARR(1) & "  " & RX_BYTE_ARR(2) & "  " & RX_BYTE_ARR(3) & "  " & RX_BYTE_ARR(4) & "  " & RX_BYTE_ARR(5)

'Debug.Print "Received " & RX_BYTE_ARR

End Sub


