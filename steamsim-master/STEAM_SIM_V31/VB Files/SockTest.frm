VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form SockTest 
   Caption         =   "Socket Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin CSSocketWrench.Socket Socket1 
      Left            =   480
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin CSSocketWrench.Socket Socket2 
      Index           =   0
      Left            =   180
      Top             =   1920
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin CSSocketWrench.Socket Socket2 
      Index           =   1
      Left            =   840
      Top             =   1980
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   2220
      Width           =   1395
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1440
      Width           =   1755
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1755
   End
   Begin VB.Label Label3 
      Caption         =   "Reply"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Send"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Host"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   675
   End
End
Attribute VB_Name = "SockTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Socket1.HostName = Trim$(Text1.Text)
    Socket1.RemotePort = IPPORT_ECHO
    Socket1.Action = SOCKET_CONNECT

End Sub

Private Sub Form_Load()

    Socket1.AddressFamily = AF_INET
    Socket1.Protocol = IPPROTO_IP
    
 '   Socket1.Type = SOCK_STREAM
     Socket1.SocketType = SOCK_STREAM
    
'     Socket1.Type = SOCK_DGRAM
    Socket1.SocketType = SOCK_DGRAM
    
    Socket1.Binary = False
    
 ' WON'T WORK With SOCK_DGRAM...
'     Socket1.BufferSize = 1024
    
    Socket1.Blocking = False
    
 
    
    Socket2(0).AddressFamily = AF_INET
    Socket2(0).Protocol = IPPROTO_IP
    
 '   Socket2(0).Type = SOCK_STREAM
     Socket2(0).SocketType = SOCK_STREAM
     
    Socket2(0).Blocking = False
    Socket2(0).LocalPort = IPPORT_ECHO
    Socket2(0).Action = SOCKET_LISTEN
    lastsocket = 0
   
    
     
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim I As Integer
 
    If Socket1.Connected Then Socket1.Action = SOCKET_CLOSE
    End

    If Socket2(0).Listening Then Socket2(0).Action = SOCKET_CLOSE
    For I = 1 To lastsocket
        If Socket2(I).Connected Then Socket2(I).Action = SOCKET_CLOSE
    Next I
    End



End Sub

Private Sub Socket1_Connect()

    Text2.Enabled = True
    Text3.Enabled = True

End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)

    Socket1.RecvLen = DataLength
    Text3.Text = Socket1.RecvData

End Sub

Private Sub Socket2_Accept(index As Integer, SocketId As Integer)

    Dim I As Integer
    
    For I = 1 To lastsocket
        If Not Socket2(I).Connected Then Exit For
    Next I
    If I > lastsocket Then
        lastsocket = lastsocket + 1: I = lastsocket
        Load Socket2(I)
    End If
    Socket2(I).AddressFamily = AF_INET
    Socket2(I).Protocol = IPPROTO_IP
    Socket2(I).Type = SOCK_STREAM
    Socket2(I).Binary = True
    Socket2(I).BufferSize = 1024
    Socket2(I).Blocking = False
    Socket2(I).Accept = SocketId

End Sub

Private Sub Socket2_Read(index As Integer, DataLength As Integer, IsUrgent As Integer)

    Socket2(index).RecvLen = DataLength
    Socket2(index).SendLen = DataLength
    Socket2(index).SendData = Socket2(index).RecvData

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Socket1.SendLen = Len(Text2.Text)
        Socket1.SendData = Text2.Text
        KeyAscii = 0: Text2.Text = ""
    End If

End Sub
