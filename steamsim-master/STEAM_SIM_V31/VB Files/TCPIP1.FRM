VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frm_tcpip 
   Caption         =   "TCP/IP Test"
   ClientHeight    =   2760
   ClientLeft      =   9330
   ClientTop       =   4845
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2760
   ScaleWidth      =   4065
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   120
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3201
      _Version        =   327680
      Enabled         =   0   'False
      ScrollBars      =   2
      RightMargin     =   3000.811
      TextRTF         =   $"tcpip1.frx":0000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Reply:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frm_tcpip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub connect()
Dim i As Integer

' Connect the sockets up
For i = 0 To Max_Host_sock_Tx
    frm_Host.Host_sock_Tx(i).Action = SOCKET_CONNECT
Next i

End Sub



Sub tcpip_send_array_single(Socket_No As Integer, ByRef REAL_ARRAY() As Single)
Dim i As Integer
Dim ch As String

' WORKED BUT a LOT of CPU TIME used:     ( HOST IDLE time 80% to 89% , PCP IDLE time 72% to 79 %)
'For i = 0 To array_size
'    On Error GoTo ErrorHandler
'    ch = Str$(REAL_ARRAY(i))
'    frm_host.Host_sock_Tx(Socket_No).SendLen = Len(ch + Chr$(10))
'    frm_host.Host_sock_Tx(Socket_No).SendData = ch & Chr$(10)
'    frm_host.Host_sock_Tx(Socket_No).Action = SOCKET_FLUSH
'Next i


' WORKED BUT MUCH less CPU TIME used:   ( HOST IDLE time 97 % , PCP IDLE time 89% )
    On Error GoTo ErrorHandler
    ch = "#"
    
' 17/3/98
For i = 0 To Array_Size
'For i = 0 To Real_Array_Size
    ch = ch & Str$(REAL_ARRAY(i)) & "#"
Next i
    frm_Host.Host_sock_Tx(Socket_No).SendLen = Len(ch + Chr$(10))
    frm_Host.Host_sock_Tx(Socket_No).SendData = ch & Chr$(10)
    frm_Host.Host_sock_Tx(Socket_No).Action = SOCKET_FLUSH


Exit Sub


ErrorHandler:

End Sub

Private Sub Command2_Click()
    Dim va As Variant

'    Call send_array(Val_Array())
    Call tcpip_send_array_single(0, T42DIESELR())
End Sub

Private Sub Form_Load()
Dim i As Integer

    frm_Host.MousePointer = 11

' 20/1/97
    On Error GoTo error_handler:
    

' Make 1..n extra sockets (the first, 0, is always present on the form
'For i = 1 To Max_Host_sock_Tx
    'Load frm_host.Host_sock_Tx(i)
'Next i


' 16/3/98  REMOVED
'For i = 0 To Max_Host_sock_Tx
'  If i < 2 Then
'    frm_Host.Host_sock_Tx(i).HostAddress = PCP_TCPIP_address
'  Else
'    frm_Host.Host_sock_Tx(i).HostAddress = SCP_TCPIP_address
'  End If
'
'    frm_Host.Host_sock_Tx(i).AddressFamily = AF_INET
'    frm_Host.Host_sock_Tx(i).Protocol = IPPROTO_IP
'
'   frm_Host.Host_sock_Tx(i).SocketType = SOCK_STREAM
'
'    frm_Host.Host_sock_Tx(i).Binary = False
'    frm_Host.Host_sock_Tx(i).RemotePort = Port_1 + i
'    frm_Host.Host_sock_Tx(i).BufferSize = 32767
'    frm_Host.Host_sock_Tx(i).Blocking = False
'Next i


'    Counter = 0
'    Label4.Caption = PCP_TCPIP_address

'    Safe_to_continue = False
     
'    connect

' '   ' Wait for a result
'    While Not Safe_to_continue
'        DoEvents
'    Wend
'
'    If Connected Then
'        Timer1.Enabled = True
'    End If
'
'    frm_Host.MousePointer = 0

'Exit Sub

'error_handler:
'MsgBox "Socket 1 error"

End Sub

Private Sub Form_Resize()
Dim i As Integer

i = Me.ScaleWidth - RichTextBox1.Left
If i > 0 Then
    RichTextBox1.Width = i
    RichTextBox1.RightMargin = RichTextBox1.Width - 360
End If

i = Me.ScaleHeight - RichTextBox1.Top
If i > 0 Then RichTextBox1.Height = i


End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer



End Sub









Private Sub Text2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Host_sock_Tx.SendLen = Len(Text2.Text)
'    Host_sock_Tx.SendData = Text2.Text
'    KeyAscii = 0: Text2.Text = ""
'End If
End Sub


Private Sub Timer1_Timer()
    Call tcpip_send_array_single(0, T42DIESELR())
    Call tcpip_send_array_single(1, T42BUSR())
End Sub


