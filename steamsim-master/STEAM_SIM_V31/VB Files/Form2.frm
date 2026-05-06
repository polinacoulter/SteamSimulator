VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   3855
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   2700
      Width           =   4515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send to PC02"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1140
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2940
      TabIndex        =   4
      Top             =   1740
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send to PC01"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1140
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1020
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
      RemoteHost      =   "1.1.1.1"
      RemotePort      =   1
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3900
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
      RemoteHost      =   "1.1.1.2"
      LocalPort       =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Tx to PC02"
      Height          =   255
      Left            =   3660
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tx to PC01"
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Winsock1.SendData (Text3.Text)

End Sub


Private Sub Command2_Click()

Winsock2.SendData (Text3.Text)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

    Winsock1.GetData strData, vbString
    Text1.Text = Text1.Text & strData

End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

    Winsock2.GetData strData, vbString
    Text1.Text = Text1.Text & strData

End Sub

