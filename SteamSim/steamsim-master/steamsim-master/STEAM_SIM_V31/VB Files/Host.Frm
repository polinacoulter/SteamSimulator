VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frm_Host 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Marine Engine Simulator"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   17910
   LinkTopic       =   "MDIForm1"
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   0
      Left            =   540
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   1
      Left            =   1020
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   2
      Left            =   1980
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   3
      Left            =   2460
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   4
      Left            =   3180
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   5
      Left            =   3660
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   6
      Left            =   4380
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   7
      Left            =   4860
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   8
      Left            =   5700
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   9
      Left            =   6180
      Top             =   3480
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   0
      Left            =   540
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   1
      Left            =   1140
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   2
      Left            =   1740
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   3
      Left            =   2280
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   4
      Left            =   2820
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   5
      Left            =   3300
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Client_sock_Tx 
      Left            =   8640
      Top             =   7380
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Client_Sock_Rx 
      Index           =   0
      Left            =   9300
      Top             =   7380
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   10
      Left            =   600
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   11
      Left            =   1080
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   12
      Left            =   1980
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   13
      Left            =   2460
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   14
      Left            =   3180
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   15
      Left            =   3660
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   16
      Left            =   4380
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   17
      Left            =   4860
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   18
      Left            =   5700
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   19
      Left            =   6180
      Top             =   4260
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   20
      Left            =   540
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   21
      Left            =   1020
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   22
      Left            =   1920
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   23
      Left            =   2400
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   24
      Left            =   3120
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   25
      Left            =   3600
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   26
      Left            =   4320
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   27
      Left            =   4800
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   28
      Left            =   5640
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket Host_sock_Tx 
      Index           =   29
      Left            =   6120
      Top             =   5040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   6
      Left            =   540
      Top             =   7620
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   7
      Left            =   1140
      Top             =   7620
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   8
      Left            =   1740
      Top             =   7620
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   9
      Left            =   2280
      Top             =   7620
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   10
      Left            =   2820
      Top             =   7620
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   11
      Left            =   3300
      Top             =   7620
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   12
      Left            =   600
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   13
      Left            =   1200
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   14
      Left            =   1800
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   15
      Left            =   2340
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   16
      Left            =   2880
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin SocketWrenchCtrl.Socket HostSockRx 
      Index           =   17
      Left            =   3360
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
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
   Begin VB.Timer Super_Pro_Dongle_Timer 
      Interval        =   1000
      Left            =   4440
      Top             =   2700
   End
   Begin VB.Timer Dongle_Timer 
      Interval        =   5000
      Left            =   4680
      Top             =   2040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer_IO 
      Interval        =   250
      Left            =   5280
      Top             =   1620
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4560
      Top             =   1080
   End
   Begin VB.Timer Exec_Timer 
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4920
      Top             =   600
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   17910
      _ExtentX        =   31591
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Text            =   "Unknown"
            TextSave        =   "12:37"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3176
            MinWidth        =   3176
            Object.Tag             =   ""
            Object.ToolTipText     =   "Simulator Start Mode"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10584
            MinWidth        =   10584
            Picture         =   "Host.frx":0000
            Object.Tag             =   ""
            Object.ToolTipText     =   "Latest Active Alarm"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Rx Sock"
            TextSave        =   "Rx Sock"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Tx Sock"
            TextSave        =   "Tx Sock"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "Host.frx":031A
            Object.Tag             =   ""
            Object.ToolTipText     =   "Engineer Call"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Host.frx":0634
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_save_snapshot 
         Caption         =   "Save Snapshot"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu_run 
      Caption         =   "Run"
      Begin VB.Menu mnu_cold 
         Caption         =   "Cold Start"
      End
      Begin VB.Menu mnu_load_snapshot 
         Caption         =   "Load Snapshot"
      End
      Begin VB.Menu mnuHalt 
         Caption         =   "Halt"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuContinue 
         Caption         =   "Continue"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Electrical 
      Caption         =   "Electrical"
      Begin VB.Menu mnuSY200 
         Caption         =   "Main Switchboard"
      End
      Begin VB.Menu mnuSWBD11 
         Caption         =   "Switchboard #1"
      End
      Begin VB.Menu mnuSWBD13 
         Caption         =   "Switchboard #2"
      End
      Begin VB.Menu mnuEmergSWBD 
         Caption         =   "Emergency Switchboard"
      End
      Begin VB.Menu mnuDA 
         Caption         =   "Diesel Generator"
      End
      Begin VB.Menu mnuTurbo 
         Caption         =   "Turbo Generators"
      End
   End
   Begin VB.Menu mnu_systems 
      Caption         =   "Aux Systems"
      Begin VB.Menu mnuSalt_Water 
         Caption         =   "Salt Water Services"
      End
      Begin VB.Menu mnuDistill 
         Caption         =   "Distiller"
      End
      Begin VB.Menu mnu_main_engine 
         Caption         =   "Main Engine"
         Begin VB.Menu mnuSY010 
            Caption         =   "Main Engine Overview"
         End
         Begin VB.Menu mnuME 
            Caption         =   "Main Engine"
         End
         Begin VB.Menu mnuMEcontrol 
            Caption         =   "Main Engine Control"
         End
         Begin VB.Menu mnuScavAir 
            Caption         =   "Scavenge Air"
         End
         Begin VB.Menu mnuExhaust 
            Caption         =   "Exhaust Gasses"
         End
         Begin VB.Menu mnuRunGear 
            Caption         =   "Running Gear"
         End
         Begin VB.Menu mnuCrank 
            Caption         =   "Crank Case"
         End
         Begin VB.Menu mnuOMD 
            Caption         =   "Oil Mist Detector"
         End
      End
      Begin VB.Menu mnuLubrication 
         Caption         =   "Lubrication"
         Begin VB.Menu mnuLub 
            Caption         =   "M.E.Lub Oil"
         End
         Begin VB.Menu mnuLOSEP 
            Caption         =   "Lub Oil Storage and Separator"
         End
         Begin VB.Menu mnuCyllub 
            Caption         =   "Cylinder Lub Oil"
         End
         Begin VB.Menu mnuGBOX 
            Caption         =   "GearBox LO"
         End
         Begin VB.Menu mnuCAMSHAFT 
            Caption         =   "Camshaft LO"
         End
      End
      Begin VB.Menu mnuCooling 
         Caption         =   "Cooling"
         Begin VB.Menu mnuHTCOOL 
            Caption         =   "High Temperature"
         End
         Begin VB.Menu mnuHTCWME 
            Caption         =   "High Temperature M.E. Monitoring"
         End
         Begin VB.Menu mnuLTCW 
            Caption         =   "Low Temperature"
         End
         Begin VB.Menu mnuSY096 
            Caption         =   "Low Temperature Scavenge Air Coolers"
         End
         Begin VB.Menu mnuSY091 
            Caption         =   "Low Temperature Coolers"
         End
         Begin VB.Menu mnusy101 
            Caption         =   "Salt Water"
         End
         Begin VB.Menu mnusy104 
            Caption         =   "Salt Water Coolers"
         End
         Begin VB.Menu mnuFVCW 
            Caption         =   "Fuel Valve Cooling"
         End
      End
      Begin VB.Menu mnuCPP 
         Caption         =   "CPP"
      End
      Begin VB.Menu mnuFuel 
         Caption         =   "Fuel"
         Begin VB.Menu mnusy071 
            Caption         =   "Fuel Oil Supply"
         End
         Begin VB.Menu mnusy072 
            Caption         =   "Fuel Oil Ciculating"
         End
         Begin VB.Menu mnuFOMainEngine 
            Caption         =   "Fuel Oil M.E."
         End
         Begin VB.Menu mnuDOBunkers 
            Caption         =   "DO Bunkers"
         End
         Begin VB.Menu mnuDOSEP 
            Caption         =   "DO Separator"
         End
         Begin VB.Menu mnuHOBunkers 
            Caption         =   "HO Bunkers"
         End
         Begin VB.Menu mnuHOSEP 
            Caption         =   "HO Separator"
         End
         Begin VB.Menu mnuHOSEP2 
            Caption         =   "HO Separator No 2"
         End
      End
      Begin VB.Menu mnuCOMPAIR 
         Caption         =   "Compressed Air"
         Begin VB.Menu mnuAIRCOMP 
            Caption         =   "Air Compressors"
         End
         Begin VB.Menu mnuAIRbottles 
            Caption         =   "Air Bottles"
         End
         Begin VB.Menu mnuDAcompair 
            Caption         =   "Diesel Alternators Air Start"
         End
      End
      Begin VB.Menu mnuBilge 
         Caption         =   "Bilge"
         Begin VB.Menu mnuBilgesys 
            Caption         =   "Bilge System"
         End
         Begin VB.Menu mnuBilgeTanks 
            Caption         =   "Bilge Tanks"
         End
      End
      Begin VB.Menu mnuFire 
         Caption         =   "Fire Main"
      End
      Begin VB.Menu MNUFWG 
         Caption         =   "Fresh Water Generator"
      End
      Begin VB.Menu mnuBoiler 
         Caption         =   "Boiler"
         Begin VB.Menu mnuBOILERctrl 
            Caption         =   "Boiler Control"
         End
         Begin VB.Menu mnuSY150 
            Caption         =   "Burner FO Supply"
         End
         Begin VB.Menu mnuSY151 
            Caption         =   "Boiler Combustion"
         End
         Begin VB.Menu mnuSY155 
            Caption         =   "Boiler Drum"
         End
         Begin VB.Menu mnuSY156 
            Caption         =   "Combustion Logic"
         End
         Begin VB.Menu mnuSY161 
            Caption         =   "Feed Water"
         End
         Begin VB.Menu mnuSY164 
            Caption         =   "Condenser and Steam Returns"
         End
         Begin VB.Menu mnuSY171 
            Caption         =   "Exhaust Gas Boiler"
         End
         Begin VB.Menu mnuSY172 
            Caption         =   "Sootblowers"
         End
      End
      Begin VB.Menu mnu_Bilge_OWS 
         Caption         =   "Bilge and OWS"
      End
      Begin VB.Menu mnuBunkers 
         Caption         =   "Bunkers"
      End
      Begin VB.Menu mnuLube_oil 
         Caption         =   "Lube Oil"
      End
      Begin VB.Menu mnu_LO_Purifier 
         Caption         =   "Lube Oil Purifier"
      End
      Begin VB.Menu mnu_Control_Air 
         Caption         =   "Control Air"
      End
      Begin VB.Menu mnu_Ballast 
         Caption         =   "Ballast"
      End
   End
   Begin VB.Menu mnuSteam_propulsion 
      Caption         =   "Steam Systems"
      Begin VB.Menu mnuCondensate 
         Caption         =   "Condensate"
      End
      Begin VB.Menu mnuMain_Turbines 
         Caption         =   "Main Turbines"
      End
      Begin VB.Menu mnuMain_Condenser 
         Caption         =   "Main Condenser"
      End
      Begin VB.Menu mnuMain_Bearings 
         Caption         =   "Main Bearings"
      End
      Begin VB.Menu mnu_Feed_alone 
         Caption         =   "Feed Water"
      End
      Begin VB.Menu mnu_Aux_Steam_Alone 
         Caption         =   "Aux Steam"
      End
      Begin VB.Menu mnuFeed 
         Caption         =   "Feed and Aux Steam"
      End
      Begin VB.Menu mnu_BoilStbd 
         Caption         =   "Starboard Boiler"
      End
      Begin VB.Menu mnu_boilport 
         Caption         =   "Port Boiler"
      End
      Begin VB.Menu mnuGland_steam 
         Caption         =   "Gland Seal"
      End
      Begin VB.Menu mnuFOserv 
         Caption         =   "Fuel Oil Service"
      End
      Begin VB.Menu mnuFuel_system 
         Caption         =   "Fuel Oil to Burners"
      End
      Begin VB.Menu mnuForced_Draft 
         Caption         =   "Combustion Air"
      End
      Begin VB.Menu mnuBurner_Management 
         Caption         =   "Burner Management"
         Begin VB.Menu mnu_BNR_Port 
            Caption         =   "Port"
         End
         Begin VB.Menu mnu_Bnr_stbd 
            Caption         =   "Starboard"
         End
      End
      Begin VB.Menu mnu_Controllers 
         Caption         =   "Controllers"
         Begin VB.Menu mnu_Drum_STBD 
            Caption         =   "Drum Level STBD"
         End
         Begin VB.Menu mnu_Drum_PORT 
            Caption         =   "Drum Level PORT"
         End
         Begin VB.Menu mnu_feed_pump_STBD 
            Caption         =   "Feed Pump STBD"
         End
         Begin VB.Menu mnu_feed_pump_PORT 
            Caption         =   "Feed Pump PORT"
         End
         Begin VB.Menu mnu_Steam_Temp_STBD 
            Caption         =   "Steam Temperature STBD"
         End
         Begin VB.Menu mnu_Steam_Temp_PORT 
            Caption         =   "Steam Temperature PORT"
         End
         Begin VB.Menu mnu_Fuel_controller_STBD 
            Caption         =   "Fuel Oil STBD"
         End
         Begin VB.Menu mnu_Fuel_controller_PORT 
            Caption         =   "Fuel Oil PORT"
         End
         Begin VB.Menu mnu_Air_controller_STBD 
            Caption         =   "Air STBD"
         End
         Begin VB.Menu mnu_Air_controller_PORT 
            Caption         =   "Air PORT"
         End
         Begin VB.Menu mnu_Fuel_Air_STBD 
            Caption         =   "Fuel/Air Ratio STBD"
         End
         Begin VB.Menu mnu_Fuel_Air_PORT 
            Caption         =   "Fuel/Air Ratio PORT"
         End
         Begin VB.Menu mnu_MASTER_STBD 
            Caption         =   "MASTER STBD"
         End
         Begin VB.Menu mnu_MASTER_PORT 
            Caption         =   "MASTER PORT"
         End
         Begin VB.Menu mnu_PLANT_MASTER 
            Caption         =   "PLANT MASTER"
         End
         Begin VB.Menu mnu_Controller_status 
            Caption         =   "CONTROLLER STATUS"
         End
         Begin VB.Menu mnu_COMBUST_CONTROL 
            Caption         =   "COMBUSTION CONTROL"
         End
      End
      Begin VB.Menu mnuSteamSteam 
         Caption         =   "Contaminated Steam Evaporator"
      End
      Begin VB.Menu mnu_VIAX 
         Caption         =   "VIAX"
      End
   End
   Begin VB.Menu mnu_student 
      Caption         =   "Student"
      Begin VB.Menu mnu_actions 
         Caption         =   "Actions"
      End
      Begin VB.Menu mnu_plotvars 
         Caption         =   "Plot Variables"
      End
      Begin VB.Menu mnu_dynamics 
         Caption         =   "Engine Dynamics"
      End
      Begin VB.Menu mnu_EOT_MECD 
         Caption         =   "EOT"
      End
      Begin VB.Menu mnuBoilTest 
         Caption         =   "Boiler Test"
      End
      Begin VB.Menu mnuSteam_Calc 
         Caption         =   "Steam Calculator"
      End
   End
   Begin VB.Menu mnuAlarms 
      Caption         =   "Alarms"
   End
   Begin VB.Menu mnu_instructor 
      Caption         =   "Instructor"
      Begin VB.Menu mnu_configure 
         Caption         =   "Configure Network"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_faults 
         Caption         =   "Faults"
      End
      Begin VB.Menu mnueot 
         Caption         =   "EOT"
      End
      Begin VB.Menu mnu_logging 
         Caption         =   "Logging"
         Begin VB.Menu mnu_startlog 
            Caption         =   "Start Log"
         End
         Begin VB.Menu mnu_endlog 
            Caption         =   "End Log"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_viewlog 
            Caption         =   "View Log"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_printlog 
            Caption         =   "Print Log"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnu_parameters 
         Caption         =   "Parameters"
      End
      Begin VB.Menu mnu_real_PID 
         Caption         =   "Real PID"
      End
      Begin VB.Menu mnuAdvanced 
         Caption         =   "Advanced Mode"
      End
      Begin VB.Menu mnuINaGROUP 
         Caption         =   "In a Group"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUseEngineRoom 
         Caption         =   "Use Engine Room"
      End
      Begin VB.Menu mnuControl 
         Caption         =   "Control"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuPCINCONTROL 
            Caption         =   "PC in Control"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuStud 
         Caption         =   "Student Actions"
         Begin VB.Menu mnu_DISABLE_ACT_STUDENT_Actions 
            Caption         =   "Disable ACT Student Actions"
         End
         Begin VB.Menu mnu_DISABLE_ACT_STUDENT_Action_DISPLAY 
            Caption         =   "Disable ACT Student Action Display"
         End
      End
      Begin VB.Menu mnuModel_Constants 
         Caption         =   "Model Constants"
      End
   End
   Begin VB.Menu mnuHardware 
      Caption         =   "&Hardware"
      Begin VB.Menu mnuHWtest 
         Caption         =   "&HardWare Tests ( DI/DO/AI/AO)"
      End
      Begin VB.Menu mnuTest_Scope 
         Caption         =   "&Hardware Test Synchro-Scope"
      End
   End
   Begin VB.Menu mnuSounds 
      Caption         =   "Sounds"
      Begin VB.Menu mnuEnableSounds 
         Caption         =   "Enable Sounds"
      End
      Begin VB.Menu mnuSoundMix 
         Caption         =   "Sound Mixer"
      End
   End
   Begin VB.Menu mnu_Time 
      Caption         =   "Time"
      Begin VB.Menu mnu_FAST 
         Caption         =   "Fast Time"
      End
      Begin VB.Menu mnu_Normal 
         Caption         =   "Normal Time"
      End
      Begin VB.Menu mnu_SLOW 
         Caption         =   "Slow Time"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Begin VB.Menu mnu_Portrait 
         Caption         =   "Portrait"
      End
      Begin VB.Menu mnu_Landscape 
         Caption         =   "Landscape"
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "Language"
      Begin VB.Menu mnuOther 
         Caption         =   "Other"
      End
      Begin VB.Menu mnuEnglish 
         Caption         =   "English"
      End
   End
   Begin VB.Menu mnuUnits 
      Caption         =   "Units"
      Begin VB.Menu mnuANSI 
         Caption         =   "ANSI"
      End
      Begin VB.Menu mnuSI 
         Caption         =   "SI"
      End
   End
   Begin VB.Menu mnu_Screen 
      Caption         =   "Screen"
      Begin VB.Menu mnu_Back_Colour 
         Caption         =   "Background Colour"
      End
      Begin VB.Menu mnu_Back_Grey 
         Caption         =   "Background Grey"
      End
      Begin VB.Menu mnu_Back_White 
         Caption         =   "Background White"
      End
   End
End
Attribute VB_Name = "frm_Host"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Client_StartUp()

Me.StatusBar1.Panels.Item(2).text = "Start"

d_run = True
e_run = True

End Sub
Sub SendMessage(TextString As String)

' 25/6/98
'If Not test_flag Then
If In_a_group Then

    If ClientSockTxConnected Then
        frm_Host.Client_sock_Tx.SendLen = Len(TextString)
        frm_Host.Client_sock_Tx.SendData = TextString
        Debug.Print "Sent " & TextString
    Else
        Debug.Print "Can't send " & TextString & " - ClientTxSocket not connected "
    End If
    
End If

End Sub

Sub tcpip_initialise()

'executed by clients only

Dim i As Integer

Debug.Print "Entered tcpip_initialise"
' Make the new sockets, as socket(0) already exists
For i = 1 To Max_Client_Sock_Rx
    Load Client_Sock_Rx(i)
    Debug.Print "loaded client_sock_Rx(" & i & ")"

Next i

' Initialise incoming sockets
For i = 0 To Max_Client_Sock_Rx 'default is 2 sockets ( 0 to 1 )
    '17/6/98 Client_Sock_Rx(i).LocalPort = Host_Tx_Base + i 'should be base + 2*PC number + i
    Client_Sock_Rx(i).LocalPort = Host_Tx_Base + (2 * (This_PC_number - 1)) + i 'should be base + 2*PC number + i
    Client_Sock_Rx(i).AddressFamily = AF_INET
    Client_Sock_Rx(i).Protocol = IPPROTO_IP
    Client_Sock_Rx(i).SocketType = SOCK_STREAM
    Client_Sock_Rx(i).BufferSize = 32767
    Client_Sock_Rx(i).Blocking = False
    Client_Sock_Rx(i).Action = SOCKET_LISTEN
    client_counter(i) = 0
Next i
    Debug.Print "initialized client sockets 1"

' Initialise outgoing socket

'16/6/98Client_sock_Tx.HostAddress = HostAddrPC02
Client_sock_Tx.HostAddress = HostAddrPC00
Client_sock_Tx.AddressFamily = AF_INET
Client_sock_Tx.Protocol = IPPROTO_IP
Client_sock_Tx.SocketType = SOCK_STREAM
Client_sock_Tx.Blocking = False
Client_sock_Tx.Binary = False
Client_sock_Tx.BufferSize = 32767
'17/6/98 Client_sock_Tx.RemotePort = Host_Tx_Base 'should refer to particular host
Client_sock_Tx.RemotePort = Host_Rx_Base + This_PC_number 'HostRx -> Client Tx is 1:1
'this is very important (25/3/98) if, successful it causes a CONNECT event on the host
'which kicks everything off
Client_sock_Tx.Action = SOCKET_CONNECT

    Debug.Print "initialized client sockets 2"
    



Exit Sub ' and avoid error handler


End Sub

Sub retry_tcpip_initialise()
Dim MessageString As String

Debug.Print "Entered retry_tcpip_initialize "



'if at first you don't succeed then try and try again ....

'17/6/98 Client_sock_Tx.HostAddress = HostAddrPC02
Client_sock_Tx.HostAddress = HostAddrPC00
Client_sock_Tx.AddressFamily = AF_INET
Client_sock_Tx.Protocol = IPPROTO_IP
Client_sock_Tx.SocketType = SOCK_STREAM
Client_sock_Tx.Blocking = False
Client_sock_Tx.Binary = False
Client_sock_Tx.BufferSize = 32767
'18/6/98 Client_sock_Tx.RemotePort = Host_Tx_Base 'should refer to particular host
Client_sock_Tx.RemotePort = Host_Rx_Base + This_PC_number 'HostRx -> Client Tx is 1:1

Client_sock_Tx.Action = SOCKET_CONNECT


'MessageString = "RETRY_ME"
'Call frm_Host.SendMessage(MessageString)

End Sub

Sub CheckScheduledFaults()
'Check if any scheduled faults need to be entrered
Dim i As Integer
Dim j As Integer
Dim LastFaultIndex As Integer

If FaultScheduleIndex > 0 Then ' there are some faults scheduled, so check if their time is up
    'debug.print "found a scheduled fault, time left is
    'first decrement times (in secs)
    For i = 1 To FaultScheduleIndex
        FaultSchedule(i, 3) = FaultSchedule(i, 3) - 1
        'Debug.Print "found a scheduled fault, time left is" & FaultSchedule(i, 3)
        If FaultSchedule(i, 3) <= 0 Then ' time is up
        
            LastFaultIndex = i
            
            ' 25/3/98 Fault(FaultSchedule(i, 1)) = True
                      ' 25/3/98   Send the action to the model
                  Call setFAULT(1, FaultSchedule(i, 1))

            fault_time(FaultSchedule(i, 1)) = Time
            
            Call Logging("INSTRUCTOR SCHEDULED FAULT ENTERED: " & Faultext(FaultXRef(FaultSchedule(i, 1))), LogFaults)
        
           ' move other entries above down 1 position
            
            Call ClearScheduledFault(LastFaultIndex)

            Exit For ' don't check any more this pass
        End If
    Next
End If

End Sub
Sub ClearScheduledFault(FIndex As Integer)

Dim j As Integer
            
            
 ' move other entries above down 1 position

 For j = FIndex To FaultScheduleIndex
            FaultSchedule(j, 1) = FaultSchedule(j + 1, 1)
            FaultSchedule(j, 2) = FaultSchedule(j + 1, 2)
            FaultSchedule(j, 3) = FaultSchedule(j + 1, 3)
Next
                
FaultScheduleIndex = FaultScheduleIndex - 1
End Sub
Sub connect_PC(Index As Integer)

' Connect the sockets up


'17/6/98
Debug.Print "In Connect_PC, Index = " & Index
Debug.Print "Trying to connect on Host_Sock_Tx( " & ((Index - 1) * 2) & " and " & ((Index - 1) * 2) + 1

' 25/6/98
If In_a_group Then
  
  For i = 0 To 1
         '22/11/99 frm_Host.Host_sock_Tx(((index - 1) * 2) + i).Action = SOCKET_CONNECT
         'frm_Host.Host_sock_Tx(((index) * 2) + i).Action = SOCKET_CONNECT
         frm_Host.Host_sock_Tx(((Index - 1) * 2) + i).Action = SOCKET_CONNECT
  Next
  
End If

GoTo Skip


For i = 0 To Max_Host_sock_Tx
 
     Debug.Print "entered connect_PC(" & i & ")"

   '25/3/98  If Not frm_Host.Host_sock_Tx(i).Connected Then
   If HostSockTxConnected(Index) = False Then
     Debug.Print "In connect_PC(" & i & ") trying to connect"
     frm_Host.Host_sock_Tx(i).Action = SOCKET_CONNECT
        
    End If
    
Next i

Skip:

End Sub



Sub tcpip_send_array_integer(Socket_No As Integer, ByRef l_array() As Integer)
Dim i As Integer
Dim ch As String
Dim StrArraySize As String * 5

    
     On Error GoTo errorhandler




' Convert integers to string


' 27/5/98
If medium_speed Then
   ' 23/3/98   ( get from model.dll )
    Call get_I_string(I_string)
Else
    Call get_I_string_LSS(I_string)
End If

'Debug.Print "tcpip_send_array_integer LEN(I_string) = " & Len(I_string)


' 16/6/98
'Mid(I_string, 1, 4) = "STRT"
If medium_speed Then
   Mid(I_string, 1, 4) = "SMED"
Else
   Mid(I_string, 1, 4) = "SSLO"
End If

Mid(I_string, 3999, 4) = "_END"


frm_Host.Host_sock_Tx(Socket_No).Write I_string, Len(I_string)

frm_Host.Host_sock_Tx(Socket_No).Action = SOCKET_FLUSH

Exit Sub


errorhandler:
    


End Sub



Sub tcpip_send_array_single(Socket_No As Integer, ByRef REAL_ARRAY() As Single)
Dim i As Integer
Dim ch As String
Dim StrArraySize As String * 5


    On Error GoTo errorhandler


'Debug.Print "tcpip_send_array_single , REAL_ARRAY(1) = " & REAL_ARRAY(1)
'Debug.Print "Len(REAL_ARRAY(1)) = " & Len(REAL_ARRAY(1))



' Convert reals to string

If medium_speed Then
    ' 19/3/98   ( get from model.dll )
    Call get_R_string(R_string)
Else
    ' 27/5/98   ( get from SIM_LSS.dll )
    Call get_R_string_LSS(R_string)
End If

' 16/6/98
'Mid(R_string, 1, 4) = "STRT"
If medium_speed Then
   Mid(R_string, 1, 4) = "SMED"
Else
   Mid(R_string, 1, 4) = "SSLO"
End If

'mid(R_string, Len(R_string) - 3, Len(R_string)) = "_END"
Mid(R_string, 8001, 4) = "_END"


frm_Host.Host_sock_Tx(Socket_No).Write R_string, Len(R_string)
'Debug.Print "number of bytes actually sent = " & frm_Host.Host_sock_Tx(Socket_No).SendLen

frm_Host.Host_sock_Tx(Socket_No).Action = SOCKET_FLUSH

Exit Sub


errorhandler:

End Sub

Sub Retry_PC()


End Sub
Private Sub Client_Sock_Rx_Accept(Index As Integer, SocketId As Integer)
Dim i(0 To 127) As Integer

Debug.Print "Entered client_sock_Rx(i) Accept,   Index = " & Index
Client_Sock_Rx(Index).Action = SOCKET_ACCEPT
' client_sock_Rx(Index).Accept = SocketId

End Sub

Private Sub Client_Sock_Rx_Connect(Index As Integer)
'RichTextBox1.Enabled = True
'Command2.Enabled = True

Debug.Print "client_sock Connected, Index = " & Index
Connection_Status = "Rx Connected"

' 22/4/97
'frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

End Sub

Private Sub Client_Sock_Rx_Disconnect(Index As Integer)
Dim istatus(0 To 127)  As Integer

' Entered when any client_sock is disconnected. This would happen as a result of the executive client
' sending a RETRY command to the host causing the host's Tx sockets to be closed then re-opened OR
' the host shutting down.
' This routine effectively re-initialises the listening sockets so the host can re-connect.
' Re-initialization is by CLOSING, waiting 3 secs and then OPENING again.
    
   
'Close the listening (Rx) socket which has just been disconnected
    Client_Sock_Rx(Index).Action = SOCKET_CLOSE
    'istatus(Index) = Client_Sock_Rx(Index).State
    Debug.Print "In Client_Sock_Rx( " & Index & ") - Disconnect, Just CLOSED SOCKET "
    
'Wait then re-open so host can re-connect at some time to fresh socket(s)
    Call Sleep(3000)
    Client_Sock_Rx(Index).Action = SOCKET_OPEN
    Client_Sock_Rx(Index).Action = SOCKET_LISTEN
    'istatus(Index) = Client_Sock_Rx(Index).State
    Debug.Print "In Client_Sock_Rx( " & Index & ") - Disconnect, Just RE-OPENED & LISTENING AGAIN "
    Connection_Status = "Rx Waiting"
    
'  22/4/97
'     frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status
End Sub

Private Sub Client_Sock_Rx_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
'Dim EStr As String
'Dim Ecode As Integer

Debug.Print "client_sock_Rx(" & Index & ") - LastError : " & ErrorString & " Error Code: " & ErrorCode

End Sub

Private Sub Client_Sock_Rx_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
Dim recv As Variant
Dim n
Dim i As Integer
Dim StartPosition As Integer, EndPosition As Integer
Dim msg As String
Dim RejectIntegerString As Integer
Dim RejectRealString As Integer
Dim RxLength As Integer
Dim TxLength As Integer

Dim real_var As String

Dim recv_string As String
Dim DummyLength As Integer
Dim resp As Integer
' 23/3/98
Dim recv_integer_string As String



'Debug.Print "In Client_sock_Rx_read A, Datalength = " & DataLength & "   Index = " & Index

 Select Case Index
   Case 0                   ' REALS

             If DataLength = 8004 Then
                    
                Client_Sock_Rx(Index).Read recv_string, 8004
                'Debug.Print "LENb(recv_string) = " & LenB(recv_string)
                
                'Debug.Print "REAL recv_string(1:4) = " & Mid(recv_string, 1, 4)
                'Debug.Print "recv_string(8001:8004) = " & Mid(recv_string, 8001, 4)
               
             ElseIf DataLength > 8004 Then
             
                resp = Client_Sock_Rx(Index).Flush
                Debug.Print "DataLenghth TOO LONG  ,Flushing Buffers, Resp = " & resp
                
             Else
             
                Client_Sock_Rx(Index).Read recv_string, DataLength
                
' 16/6/98       If Mid(recv_string, 1, 4) = "STRT" Then
                If Mid(recv_string, 1, 4) = "SMED" Or Mid(recv_string, 1, 4) = "SSLO" Then
                    TempStr1 = recv_string
                    'Debug.Print "found START length = " & LenB(recv_string)
                ElseIf Mid(recv_string, DataLength - 3, 4) = "_END" Then
                    'Debug.Print "found END - making full string,  length = " & LenB(recv_string)

                    TempStr2 = TempStr1 & recv_string
                    'Debug.Print "LENb(TempStr2) = " & LenB(TempStr2)

                    recv_string = TempStr2

                Else
                    resp = Client_Sock_Rx(Index).Flush
 ' 16/6/98           Debug.Print "STRT NOT Received , Flushing Buffers, Resp = " & resp
                     Debug.Print "SMED or SSLO NOT Received , Flushing Buffers, Resp = " & resp
               End If
                 
             End If
             'Debug.Print "LENb(recv_string) = " & LenB(recv_string)

             If LenB(recv_string) = 16008 Then
             
                ' 27/5/98
                If medium_speed Then
                    Call set_R_string(ByVal recv_string)
                Else
                    Call set_R_string_LSS(ByVal recv_string)
                End If
                
 
              End If

        
  
   Case 1                 ' INTEGERS
    
             If DataLength = 4002 Then
                    
                Client_Sock_Rx(Index).Read recv_integer_string, 4002
                'Debug.Print "LENb(recv_integer_string) = " & LenB(recv_integer_string)
                
                'Debug.Print "recv_integer_string(1:4) = " & Mid(recv_integer_string, 1, 4)
                'Debug.Print "recv_integer_string(3999:4002) = " & Mid(recv_integer_string, 3999, 4)
               
             ElseIf DataLength > 4002 Then
             
                resp = Client_Sock_Rx(Index).Flush
                Debug.Print "INTEGER string DataLenghth TOO LONG  ,Flushing Buffers, Resp = " & resp
                
             Else
             
                Client_Sock_Rx(Index).Read recv_integer_string, DataLength
                
 ' 16/6/98     If Mid(recv_integer_string, 1, 4) = "STRT" Then
               
                If Mid(recv_integer_string, 1, 4) = "SMED" Or Mid(recv_integer_string, 1, 4) = "SSLO" Then
                
                
                   TempStr3 = recv_integer_string
                    'Debug.Print "found INTEGER string START length = " & LenB(recv_integer_string)
                ElseIf Mid(recv_string, DataLength - 3, 4) = "_END" Then
                    'Debug.Print "found INTEGER string END - making full string,  length = " & LenB(recv_string)

                    TempStr4 = TempStr3 & recv_integer_string
                    'Debug.Print "INTEGER string LENb(TempStr4) = " & LenB(TempStr4)

                    recv_integer_string = TempStr4

                Else
                    resp = Client_Sock_Rx(Index).Flush
' 16/6/98           Debug.Print "INTEGER string STRT NOT Received , Flushing Buffers, Resp = " & resp
                    Debug.Print "INTEGER string SMED or SSLO NOT Received , Flushing Buffers, Resp = " & resp
                End If
                 
             End If
             'Debug.Print "INTEGER string  LENb(recv_integer_string) = " & LenB(recv_integer_string)

               If LenB(recv_integer_string) = 8004 Then
             
                    If Mid(recv_integer_string, 1, 4) = "SMED" Then
                        medium_speed = True
                    Else
                        medium_speed = False
                    End If

             
                 ' 27/5/98
                 If medium_speed Then
                    Call set_I_string(ByVal recv_integer_string)
                 Else
                    Call set_I_string_LSS(ByVal recv_integer_string)
                 End If
                 
             End If
    

    
  End Select
    

Exit Sub

errorhandler:
msg = MsgBox("Error Reading TCP/IP Socket 1:  Index =  " & Index & "Error = " & error, vbOK + vbCritical, "Error Report")
Debug.Print " "
Debug.Print "client_sock Read, last data length = " & DataLength
Debug.Print "client_sock Read, last recv = " & Mid(recv, 1, 50)
Exit Sub

End Sub

Private Sub client_sock_Tx_Accept(SocketId As Integer)
Debug.Print "client_sock_Tx Accept "
End Sub

Private Sub client_sock_Tx_Blocking(Status As Integer, Cancel As Integer)
Debug.Print "client_sock_Tx Blocking "

End Sub


Private Sub client_sock_Tx_Cancel(Status As Integer, Response As Integer)
Debug.Print "client_sock_Tx Cancel "

End Sub


Private Sub client_sock_Tx_Connect()

'When Tx connection established with host, put message on Status Bar

ClientSockTxConnected = True


Connection_Status = "Tx Connected"
Debug.Print " client_sock_Tx.Connect - Connected OK time = " & Time

'   22/4/97
'    frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

End Sub

Private Sub client_sock_Tx_Disconnect()

'If host is stopped and the Tx socket (2) is disconnected, then try to reconnect every 5 seconds
'first put message on Status Bar

Connection_Status = "Tx Disconnected"

Debug.Print "client_sock_Tx CLOSED in client_sock_Tx.Disconnect calling Retry_tcpip_initialize"
ClientSockTxConnected = False

'Close the Tx socket and wait before re-initializing
Client_sock_Tx.Action = SOCKET_CLOSE
Call Sleep(5000)
Call retry_tcpip_initialise

End Sub


Private Sub client_sock_Tx_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

'If host is stopped and the Tx socket (2) is disconnected, then try to reconnect every 5 seconds

'first put message on Status Bar
Connection_Status = "Tx Disconnected"

'Close the Tx socket and wait before re-initializing
Client_sock_Tx.Action = SOCKET_CLOSE
Debug.Print "client_sock_Tx CLOSED in client_sock_Tx.Last Error calling Retry_tcpip_initialize"
Call Sleep(5000)

Call retry_tcpip_initialise
End Sub


Private Sub client_sock_Tx_Write()
Debug.Print " sent message written from client_sock_Tx.Read"

End Sub


Private Sub Dongle_Timer_Timer()

' === 12/12/02 ===
'Exit Sub

' === 11/09/03 ===
If Super_Pro_Dongle Then

    ' === NEW Super Pro Dongle... ===
     
    ' === Use the NEW Super-Pro Dongle (see Super_Pro_Dongle_Timer) ===

    For i = 81 To 100
     
        If Len(QueryTable(i).Response) <= 8 Then   'added because getting overflow errors for big queries
            UseDongleResponseStart = 1
            UseDongleResponseLength = Len(QueryTable(i).Response)
        Else
            UseDongleResponseStart = Int((Len(QueryTable(i).Response) - 8) * Rnd + 1)   'picks random start point in string
            UseDongleResponseLength = 8
        End If

        Crash_it = Sqr(Val("&H" + Mid(QueryTable(i).Response, UseDongleResponseStart, UseDongleResponseLength)) - Val("&H" + Mid(Dongle_Reading(i), UseDongleResponseStart, UseDongleResponseLength)))
    Next i


    For i = 21 To 40
     
        If Len(QueryTable(i).Response) <= 8 Then     'added because getting overflow errors for big queries
            UseDongleResponseStart = 1
            UseDongleResponseLength = Len(QueryTable(i).Response)
        Else
            UseDongleResponseStart = Int((Len(QueryTable(i).Response) - 8) * Rnd + 1)   'picks random start point in string
            UseDongleResponseLength = 8
        End If

        Crash_it = Sqr(Val("&H" + Mid(QueryTable(i).Response, UseDongleResponseStart, UseDongleResponseLength)) - Val("&H" + Mid(Dongle_Reading(i), UseDongleResponseStart, UseDongleResponseLength)))
    Next i
    
    For i = 1 To 20
         
        If Len(QueryTable(i).Response) <= 8 Then   'added because getting overflow errors for big queries
            UseDongleResponseStart = 1
            UseDongleResponseLength = Len(QueryTable(i).Response)
        Else
            UseDongleResponseStart = Int((Len(QueryTable(i).Response) - 8) * Rnd + 1)   'picks random start point in string
            UseDongleResponseLength = 8
        End If

        Crash_it = Sqr(Val("&H" + Mid(QueryTable(i).Response, UseDongleResponseStart, UseDongleResponseLength)) - Val("&H" + Mid(Dongle_Reading(i), UseDongleResponseStart, UseDongleResponseLength)))
    Next i
 
Else

    ' === OLD BASBH Dongle... ===
    
    If Dongle_Enable_Flag = True Then
    
        OPEN_SENT ' Open the dongle port
        
        ' === 11/09/03 OLD BASBH Dongle... ===
        'Dongle_Response = Format$(SSQUERY(ByVal Dongle_Password, ByVal Len(Dongle_Password)))
         Dongle_Response_Integer = Format$(SSQUERY(ByVal Dongle_Password, ByVal Len(Dongle_Password)))
        
        CLOSE_SENT ' Close the dongle port
        
        
        
    End If
    
End If

End Sub

Private Sub Exec_Timer_Timer()


Dim ETCounter  As Integer



' === 21/12/01 ===

' === 11/03/09 ===
'Dongle_Integer = Dongle_Response
Dongle_Integer = Dongle_Response_Integer




'    === 28/03/01 15.034 =======

'    LSS_LOGICALS(97) = Reserved ( See in VB it is used as Silence_BLEEP )

If LSS_LOGICALS(97) Then
    If I_AM_HOST Then
    
        Silence_BLEEP_TIME = Silence_BLEEP_TIME + 1
        
        If Silence_BLEEP_TIME > 4 Then
            Silence_BLEEP_TIME = 0
            status_flag = False
            Call set_LSS_logicals(97, status_flag)
        End If
        
    End If
Else
            Silence_BLEEP_TIME = 0
End If

'    === 28/03/01 15.034 end ===



' === 28/03/01 =============

If LSS_LOGICALS(64) Then
  mnu_DISABLE_ACT_STUDENT_Actions.Checked = True
Else
  mnu_DISABLE_ACT_STUDENT_Actions.Checked = False
End If

If LSS_LOGICALS(63) Then

      mnu_DISABLE_ACT_STUDENT_Action_DISPLAY.Checked = True
  
    ' If Not I_AM_INSTRUCTOR And Not PTT Then
    ' I MUST be an ACT Terminal (Full Mission)
    
      If Not I_AM_INSTRUCTOR Then
          ' I MUST be a STUDENT
            mnu_actions.Enabled = False
            Unload frm_Student_Actions

      End If
    
Else

      mnu_DISABLE_ACT_STUDENT_Action_DISPLAY.Checked = False
  
    ' If Not I_AM_INSTRUCTOR And Not PTT Then
    ' I MUST be an ACT Terminal (Full Mission)
    
      If Not I_AM_INSTRUCTOR Then
          ' I MUST be a STUDENT
            mnu_actions.Enabled = True
      End If
    
End If

' === 28/03/01 End of Edit ===




' 8/12/98
'If MDIform_loaded = False Then
'    Exit Sub
'End If



If medium_speed <> old_medium_speed Then

  If medium_speed Then
     Call Fill_DIText
  Else
     Call Fill_LSS_DIText
  End If

    
 '       ' 9/12/98 MOVE to INTRODUCTION.FRM
 '       ' ( we do not now have a reset  facility and the .EXE version CRASHES on PC00 when calling
 '       ' D_MODEL from HOST.FRM FORM LOAD )
 '       ' It is VERY important for the COM port to be OPEN BEFORE calling D_MODEL
 
'   If medium_speed Then
'       If Not ComOpen Then
'            Debug.Print "Opening com1 , medium speed"
'            Call open_com(1)
'            ComOpen = True
'        End If
'    Else
'        If Not ComOpen_LSS Then
'            Debug.Print "Opening com1 , slow speed"
'            Call open_com_LSS(1)
'            ComOpen_LSS = True
'        End If
'    End If



    diesel_init = True
    elec_init = True
    diesel_cold = True
    elec_cold = True

    ' === 23/01/02 ===
    'If medium_speed Then
    '    Call D_MODEL(split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    'Else
    
        Call INIT_LSS    ' ( INITIALI.F90 )
        
      ' === 06/03/03 ===
      ' Call D_MODEL_LSS(CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
        Call D_MODEL_LSS(Ballast_flag, CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    
    'End If


    d_run = True
    e_run = True
    diesel_init = False
    elec_init = False
    diesel_cold = False
    elec_cold = False

    ' === 23/01/02 ===
    'If medium_speed Then
    '    Call D_MODEL(split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    'Else
    
      ' === 06/03/03 ===
      ' Call D_MODEL_LSS(CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
        Call D_MODEL_LSS(Ballast_flag, CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    
    'End If


    'If medium_speed Then
    
        ' 14/12/98
        'For i = 1 To 450
        '    Call get_alarm_text(i, alarm_text)
        '    alarmtxt(i) = alarm_text
        'Next
            '' ????? 11/10/99 Call get_alarm_text
        
    'Else
        'For i = 1 To 420
            'Call get_alarm_text_LSS(i, alarm_text)
            'alarmtxt(i) = alarm_text
       ' Next
    'End If


    If medium_speed Then
        Call get_elec_addresses(ESTARTSW(1), _
        ESTOPSW(1), EAUTOSW(1), EMANSW(1), ESTBY1SW(1), ESTBY2SW(1), _
        AVRAUTOSW(1), AVRMANSW(1), MBKSW(1), OPENMBKSW(1), AVR_pot(1), _
        SYNCHSELSW(1), CHSYNCHSW, SYNCHOFFSW, SELLAMPSSW, SELSYNCHSW, _
        REVRESET(1))

        Call get_bkr_addresses(BKR(1, 1))
        Call get_V_addresses(DI_VIO(1), DO_VIO(1))
        Call get_VV_addresses(DI_ON_VVIO(1), DI_OFF_VVIO(1))
        Call get_P_addresses(PumpIO(1, 1))
        Call get_RMA_addresses(RMAIO(1, 1))
  
    Else
        Call get_VV_addr_LSS(DI_ON_VVIO_LSS(1), DI_OFF_VVIO_LSS(1))
        Call get_BP_addr(PumpIO_LSS(1, 1))
        Call get_RMA_addr_LSS(RMAIO_LSS(1, 1))
  
        Call get_elec_addr_LSS(ESTARTSW(1), _
        ESTOPSW(1), EAUTOSW(1), EMANSW(1), ESTBY1SW(1), ESTBY2SW(1), _
        AVRAUTOSW(1), AVRMANSW(1), MBKSW(1), OPENMBKSW(1), AVR_pot(1), _
        SYNCHSELSW(1), CHSYNCHSW, SYNCHOFFSW, SELLAMPSSW, SELSYNCHSW, _
        REVRESET(1))
        Call get_bkr_addr_LSS(BKR(1, 1))
 
    End If

        ss_supply = False
      diesel_init = False
        elec_init = False
      diesel_cold = False
        elec_cold = False
            d_run = False
            e_run = False
        snap_shot = False
          hw_test = False



'------------------------
'    If medium_speed Then
'        mnuSLOWSPEED.Checked = False
'            mnuMEDIUM.Checked = True
'    Else
'        mnuSLOWSPEED.Checked = True
'        mnuMEDIUM.Checked = False
'    End If

    If medium_speed Then
    
        ' 11/12/98
       mnuTurbo.Enabled = False
       mnuTurbo.Visible = False


        mnuCAMSHAFT.Enabled = False
        mnuCAMSHAFT.Visible = False
     
        mnuHOSEP2.Visible = False
         mnuHOSEP.Caption = "HO Separator"
     
        mnuAIRCOMP.Visible = False

        mnuBoiler.Visible = False
    
        ' 11/6/98
        mnuGBOX.Visible = True
        mnuFVCW.Visible = True
    
    Else

        ' 11/12/98
        mnuTurbo.Enabled = True
        mnuTurbo.Visible = True
       
        mnuCAMSHAFT.Enabled = True
        mnuCAMSHAFT.Visible = True
     
        mnuHOSEP2.Visible = True
        mnuHOSEP.Caption = "HO Separator No1"
  
        mnuAIRCOMP.Visible = True
    
      
      ' === 17/05/01 ===
      ' If TARGET_this_sim <> CALIFORNIA Then
      '      mnuBoiler.Visible = True
      ' End If
    
        ' 11/6/98
        mnuGBOX.Visible = False
        mnuFVCW.Visible = False

    End If
    old_medium_speed = medium_speed
    
End If

Call CheckScheduledFaults


' === 21/12/01 ===
If Not DISABLE_PROFIBUS Then
 If Dongle_Integer <> 0 Then
  If Dongle_Integer <> Dongle_Password_Response Then
         Exit_sharply = True
  End If
 Else
 End If
End If



End Sub


Private Sub HostSockRx_Accept(Index As Integer, SocketId As Integer)
Dim i As Integer


HostSockRx(Index).Action = SOCKET_ACCEPT

Debug.Print "HostSockRX Accept ID = " & SocketId
istatus = HostSockRx(Index).State
Debug.Print "HostSockRX status (in accept) = " & istatus

End Sub

Private Sub HostSockRx_Connect(Index As Integer)
Debug.Print "HostSockRx_Connect  event Index=" & Index
'16/6/96 Call connect_PC
'1/12/99
If Index > 0 Then
Call connect_PC(Index)
End If
End Sub

Private Sub HostSockRx_Disconnect(Index As Integer)
Dim istatus As Integer


istatus = HostSockRx(Index).State
'Debug.Print "HostSockRx status = " & istatus
frm_Host.HostSockRx(Index).Action = SOCKET_CLOSE

Debug.Print "passing through HostSockRx Disconnect, just closed socket"

'istatus = Host_sock_Rx(Index).State
istatus = HostSockRx(Index).State

Debug.Print "HostSockRx status = " & istatus & "  Waiting 3 secs before opening again"

Call Sleep(3000)

frm_Host.HostSockRx(Index).Action = SOCKET_OPEN
frm_Host.HostSockRx(Index).Action = SOCKET_LISTEN
 

istatus = HostSockRx(Index).State
Debug.Print "In HostSockRx Disconnect, reopened & listening, istatus =  " & istatus
 

End Sub

Private Sub HostSockRx_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
    Debug.Print "In HostSockRx, LastError = " & ErrorString
End Sub

Private Sub HostSockRx_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
' This is to Receive from PC(index)
Dim TempStr As String

Dim n
Dim Isnap_number As Integer
Dim Iresponse As Integer, i As Integer

istatus = HostSockRx(Index).State
Debug.Print "HostSockRx status (in read) = " & istatus

HostSockRx(Index).RecvLen = DataLength
Search_String = HostSockRx(Index).RecvData


' 23/3/98 =======
'               HostSockRx(Index).Read recv_string, 8004
'                Debug.Print "LENb(recv_string) = " & LenB(recv_string)
'================


Debug.Print "Rx message = " & Search_String
Debug.Print " "

If Search_String = "CONNECT_PC" Then
    Debug.Print " Calling connect_PC from HostSockRx READ "
'    Call connect_PC
End If

'If Search_String = "RETRY_ME" Then
'    Debug.Print " 16/3/98 Calling Retry_PC from HostSockRx READ "
'    Call Retry_PC
'End If

     
Call strip_Rx_strings
   
'   For i = 1 To 5
'     Debug.Print "Rx_strings(" & i & ") = " & Rx_strings(i) & "len = " & Len(Rx_strings(i))
'   Next i
   
' 23/3/98
'MessageString = "#SET_PCDI#" & Format(nPCDI, "00000") & "#"
If Rx_strings(1) = "SET_PCDI" Then


   nPCDI = Val(Rx_strings(2))
   
    If medium_speed Then
      Call set_PCDI(nPCDI)                ' In the LOCAL model.dll
    Else
      Call set_PCDI_LSS(nPCDI)            ' In the LOCAL SIM_LSS.DLL
    End If
       
End If

'=== 11/11/02 ==========================
If Rx_strings(1) = "SET_CV" Then

    nCHANNEL = Val(Rx_strings(2))
    nValue = Val(Rx_strings(3))
    
    Call set_CONTROLLER_VARS(nCHANNEL, nValue)

End If
'=======================================

'MessageString = "#SET_ACTION#" & Format(nset, "00000") & "#" & Format(naction, "00000") & "#"
' 24/3/98
If Rx_strings(1) = "SET_ACTION" Then
   nset = Val(Rx_strings(2))
   naction = Val(Rx_strings(3))
   
       If medium_speed Then
           Call set_ACTION(nset, naction)          ' In the LOCAL model.dll
       Else
           Call set_ACTION_LSS(nset, naction)      ' In the LOCAL SIM_LSS.DLL
       End If
End If

' 24/3/98
'MessageString = "#SET_FAULT#" & Format(nset, "00000") & "#" & Format(nFAULT, "00000") & "#"
If Rx_strings(1) = "SET_FAULT" Then
   nset = Val(Rx_strings(2))
   nFault = Val(Rx_strings(3))
   
       If medium_speed Then
           Call set_FAULT(nset, nFault)          ' In the LOCAL model.dll
       Else
            Call set_FAULT_LSS(nset, nFault)     ' In the LOCAL SIM_LSS.DLL
      End If
       
End If

' 25/3/98
'MessageString = "#SET_HILO#" & Format(nalarm, "00000") & "#" & Format(HILO_Value, "#####0.000") & "#"
If Rx_strings(1) = "SET_HILO" Then
   nalarm = Val(Rx_strings(2))
   HILO_Value = Val(Rx_strings(3))
   
        If medium_speed Then
            Call set_HILO(nalarm, HILO_Value)         ' In the LOCAL model.dll
        Else
            Call set_HILO_LSS(nalarm, HILO_Value)     ' In the LOCAL SIM_LSS.DLL
        End If
         
End If

' 25/3/98
'MessageString = "#setPCcontrol#" & Format(nCHANNEL, "00000") & "#" & Format(Value, "#####0.000") & "#"
If Rx_strings(1) = "setPCcontrol" Then
   nCHANNEL = Val(Rx_strings(2))
   Value = Val(Rx_strings(3))
   
        If medium_speed Then
               Call set_PC_CONTROL(nCHANNEL, Value)         ' In the LOCAL model.dll
        Else
                Call set_PC_CONTROL_LSS(nCHANNEL, Value)    ' In the LOCAL SIM_LSS.DLL
       End If
        
End If

' 25/3/98
'MessageString = "#setAI#" & Format(nCHANNEL, "00000") & "#" & Format(nVALUE, "00000") & "#"
If Rx_strings(1) = "setAI" Then
   nCHANNEL = Val(Rx_strings(2))
   nValue = Val(Rx_strings(3))
   
        If medium_speed Then
             Call set_AI(nCHANNEL, nValue)       ' In the LOCAL model.dll
        Else
             Call set_AI_LSS(nCHANNEL, nValue)   ' In the LOCAL SIM_LSS.DLL
        End If
        
End If



' === 7/12/00 ===============
'    MessageString = "#setVSCROLL#" & Format(nCHANNEL, "00000") & "#" & Format(nValue, "00000") & "#"
If Rx_strings(1) = "setVSCROLL" Then
   nCHANNEL = Val(Rx_strings(2))
   nValue = Val(Rx_strings(3))
   
             Call set_VSCROLL(nCHANNEL, nValue)   ' In the LOCAL STEAM_SIM.DLL
        
End If
' === 7/12/00 end of edit ===


'MessageString = "#setDI#" & Format(nDI, "00000") & "#" & "TRUE" & "#"
If Rx_strings(1) = "setDI" Then

   nDI = Val(Rx_strings(2))
   
    If Rx_strings(3) = "TRUE" Then
       status_flag = True
    Else
       status_flag = False
    End If
                Call set_DI_LSS(nDI, status_flag)     ' In the LOCAL SIM_LSS.DLL
End If

' === 22/11/02 ===
' MessageString = "#setSTEAMlogicals#" & Format(nlog, "00000") & "#" & "TRUE" & "#"
If Rx_strings(1) = "setSTEAMlogicals" Then

   nlog = Val(Rx_strings(2))
   
    If Rx_strings(3) = "TRUE" Then
       status_flag = True
    Else
       status_flag = False
    End If
                Call set_STEAM_logicals(nlog, status_flag)     ' In the LOCAL dll
End If
' === 22/11/02 end of edit===


'MessageString = "#SET_PLOT_CHANNELS#"
'For i = 1 To 8
'     MessageString = MessageString & Str(plot_channels(i)) & "#"
'Next i
If Rx_strings(1) = "SET_PLOT_CHANNELS" Then
    For i = 1 To 8
      plot_channels(i) = Val(Rx_strings(i + 1))
    Next i
    
        If medium_speed Then
            Call set_plot_channels(plot_channels(1))         ' In the LOCAL model.dll
        Else
            Call set_plot_channels_LSS(plot_channels(1))     ' In the LOCAL SIM_LSS.DLL
        End If
        
End If


'10/6/98
If Rx_strings(1) = "SET_LOG" Then


  ' === 06/06/02 ===
   ' Debug.Print "In HostSockRx_Read, Received LogString = " & Rx_strings(2)
     Debug.Print "In HostSockRx_Read, SET_LOG " & Rx_strings(2) & "+" & Rx_strings(3) & "+" & Rx_strings(4)
    
    
    '04/04/00 modify to use regular logging subroutine
    'Rx_strings(2) = Rx_strings(2) + Chr(10)
    
    
    ' === 06/06/02 ===
    Rx_strings(2) = Replace(Rx_strings(2), "$", "#")

    Call Logging(Rx_strings(2), LogActions) 'these are received from remote PC, not all are actions
                                            'but we need something valid here for logging to work
    
    
    
    'THIS BIT IN 'TILL 4/4/00
    'If LogFileOpen Then
    '    Print #LogFileNumber, Rx_strings(2)
    '    CurrentLogString = CurrentLogString & Rx_strings(2)  'maintain mirror of log
    '    If LogViewOn Then 'update log being viewed
    '        frm_LogView.RichTextBox1.text = frm_LogView.RichTextBox1.text & Rx_strings(2)
    '    End If
    'End If
End If

End Sub

Private Sub MDIForm_Load()

Dim i As Integer

' === 07/06/02 ===
Dim IN_string As String

Dim PC_name As String
Dim HOST_string As String
Dim MaxNumberPC_String As String
Dim In_a_group_string As String
Dim Off_Line_string As String

Dim msg, Style, Title, Help, Ctxt, Response, MyString
Dim ErrorFlag As Byte


'3/9/99
Dim IO_ret_value As Boolean
Dim Status As Integer
Dim HW_flagA As Boolean
Dim HW_flagB As Boolean

'11/10/99 Only configure H/W if host
Dim MyStr As String

' === 07/03/07 ===
Dim SOCKET_DO_IT As Boolean


' ===============================================================================
' === 11/09/03 ===

Dim MaximKey As Integer
Dim Compare As Boolean
Dim CompIndex As Integer


If Super_Pro_Dongle Then

     ' === NEW Super Pro Dongle... ===

    ' === 11/09/03 Super-Pro Code copied from Diesel Sim ...

    Result% = RNBOsproFormatPacket(MyAPIPacket, Len(MyAPIPacket))
    Result% = RNBOsproInitialize(MyAPIPacket)




    ' === 11/09/03 ===
    'Call InitQueryTable_Max700
    Call InitQueryTable_Max600



    Result% = RNBOsproFindFirstUnit(MyAPIPacket, &H8434)


    '===== NOTE this following is to format the apipacket to make sure it's talking to the correct
    '===== maxim key ....i.e. if more than one maxim key is attached it finds the right one
    If Result% <> 0 Then
        ' Key not found...
        DongleConnected = False
    Else
    
        Do While Not DongleConnected
        
        Result% = RNBOsproRead(MyAPIPacket, 8, MaximKey)  'key found so check if it's a maxim key
        If MaximKey <> &H5432 Then
            Result% = RNBOsproFindNextUnit(MyAPIPacket)     'not a maxim key so look for next key
            
            If Result% <> 0 Then                            'no other keys so jump out of loop
                DongleConnected = False
                Exit Do
            End If
                       
        Else
        'maxim key found so query the maxim100 algorithm to check it's active....i.e. maxim100 key
            i = Int(Rnd(1) * 200) + 1     'select one random code from the list
            
            ' Get the query value
            QueryStr = GetQueryData(i)
            QueryLength = (Len(QueryTable(i).Query)) / 2
                
                
           ' === 03/02/03 ===
           'Result% = RNBOsproQuery(MyAPIPacket, maxim100_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
           
            ' === 11/09/03 ===
            'Result% = RNBOsproQuery(MyAPIPacket, maxim700_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
             Result% = RNBOsproQuery(MyAPIPacket, maxim600_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
                
            Compare = True
            For CompIndex = 0 To QueryLength - 1
                Compare = Compare And (ResponseStr.data(CompIndex) = QueryStr.data(CompIndex))
            Next CompIndex
            
            If Compare = True Then          'QueryStr same as ResponseStr so algorithm inactive...i.e. not maxim100 key
                DongleConnected = False
                Result% = RNBOsproFindNextUnit(MyAPIPacket)     'not a maxim key so look for next key
                If Result% <> 0 Then                            'no other keys so jump out of loop
                    Exit Do
                End If
            
            Else
                DongleConnected = True         'maxim100 key
            End If
        
        End If
        
        Loop
                        
    End If
    'NOTE .....after this the app talks exclusively to the key you found (MyAPIPacket contains info for that key...set by rainbow)
    '================================================================================

    Result% = RNBOsproRead(MyAPIPacket, Timeout_period_address, Timeout_period)

    System_current_date_LONG = CDate(Date)
    
    If System_current_date_LONG > 30000 Then
       System_current_date_LONG = System_current_date_LONG - 30000
    End If
       System_current_date_INT = System_current_date_LONG


    If Timeout_period <> 0 Then
    
        'System_current_date_LONG = CDate(Date)
        'If System_current_date_LONG > 30000 Then
        '   System_current_date_LONG = System_current_date_LONG - 30000
        'End If
        '   System_current_date_INT = System_current_date_LONG
    
    
        SSPRO_start_date_address = maxim100_MemAddress - 2
        SSPRO_access_code = 0

        Result% = RNBOsproRead(MyAPIPacket, SSPRO_start_date_address, SSPRO_start_date)

        If SSPRO_start_date = 0 Then
    
    
            Result% = RNBOsproWrite(MyAPIPacket, SSPRO_Write_Password, SSPRO_start_date_address, System_current_date_INT, SSPRO_access_code)
            Result% = RNBOsproWrite(MyAPIPacket, SSPRO_Write_Password, SSPRO_current_date_address, System_current_date_INT, SSPRO_access_code)

            SSPRO_start_date = System_current_date_INT
            SSPRO_current_date = System_current_date_INT
        
        Else
    
            Result% = RNBOsproRead(MyAPIPacket, SSPRO_start_date_address, SSPRO_start_date)
            Result% = RNBOsproRead(MyAPIPacket, SSPRO_current_date_address, SSPRO_current_date)
    
        End If
    
        ' === 03/02/03 ===
        Trial_frm.Show
    
    End If



    For i = 1 To 200
        ' === 17/05/02 ===
        ' Dongle_Reading(i) = DongleResponse(i)
        ' === 5/8/02 ===
        '  Dongle_Reading(I) = Dongle_2_Response(I)
            Dongle_Reading(i) = QueryTable(i).Response
    Next i
    
End If

' ==== End of Super-Pro edit=================================================



' === 05/03/02 ===
mnu_COMBUST_CONTROL.Visible = False

' === 31/05/02 ===
If Ballast_IS_included Then
    mnu_Ballast.Visible = True
Else
    mnu_Ballast.Visible = False
End If


' === 21/12/01 ===
Dongle_Integer = 0

' === 11/09/03 ===
'Dongle_Response = 0
Dongle_Response_Integer = 0

Exit_sharply = False


' === 31/01/02 ===
Bcolour = &HC0C0C0                 ' GREY
B_colour_default = &HC0C0C0      ' GREY

' ----- 10/03/04 ----------------------------
If PTT Then
    mnu_real_PID.Visible = False
End If

' ===16/07/09 ==========
If I_AM_INSTRUCTOR Or PTT Then
    mnu_Controllers.Visible = True
Else
    mnu_Controllers.Visible = False
End If
' =======================
' 21/07/09 OLD Combined Feed and Aux Steam screen...
mnuFeed.Visible = False



If TARGET_this_sim = MARYLAND Then
    ' On 7/01/02 SIU Maryland ONLY wanted the EOT_MECD on PC01....
    If I_AM_HOST Then
            mnu_EOT_MECD.Enabled = True
    Else
        If This_PC_number = 1 Then
            mnu_EOT_MECD.Enabled = True
        Else
            mnu_EOT_MECD.Enabled = False
        End If
    End If
Else
            mnu_EOT_MECD.Enabled = True
End If
' -----------------------------------------


'' === 7/01/02 ===============
'
'' === 12/11/02 ===
            mnu_EOT_MECD.Enabled = True
'If PTT Then
'
'     mnu_real_PID.Visible = False
'
'    If I_AM_HOST Then
'            mnu_EOT_MECD.Enabled = True
'    Else
'        If This_PC_number = 1 Then
'            mnu_EOT_MECD.Enabled = True
'        Else
'            mnu_EOT_MECD.Enabled = False
'        End If
'    End If
'
'Else
'            mnu_EOT_MECD.Enabled = True
'End If
'
' === 7/01/02 end of edit ===


' === 20/12/01 ===
If PTT Then
    If I_AM_INSTRUCTOR Then
        mnu_cold.Caption = "Cold Start"
    Else
        mnu_cold.Caption = "Initialise"
    End If
Else
        ' Full Mission...
        mnu_cold.Caption = "Cold Start"
End If
' === 20/12/01 ===




' === 13/03/01 ===
Call FILL_VIRTUAL_PUMP_TEXT




'14/12/00
old_d_run = Not d_run
If d_run Then
    status_flag = True
Else
    status_flag = False
End If
Call set_STEAM_logicals(195, status_flag)


HARDWARE_flag = False



' === 15/03/01 REMEMBER TO REMOVE THIS CODE BEFORE SENDING TO CMA =============

If This_PC_Name = "PC00" Then

        ' We are at HA on RWJ's PC and NOT at CMA where  This_PC_Name = "HASP00"
  
        ' So when testing at HA we need
        ' to GoTo test_label because there are NO Applicom boards in the HA PC
  
  
        ' We are simulating HASP00 here which would normally set   HARDWARE_flag  True
        ' and thus dongle_enable_flag false so let's do just that...
      
    ' Dongle_Enable_Flag = False

   ' === 23/03/01 ===
   ' GoTo test_label
    
End If

' === 15/03/01 end of edit =====================================================


'29/10/99
'If Not PTT Then
'22/11/99 stop ACTs looking for H/W

' xxxx 10/04/02 FOR TESTING FM we need GoTo Skipxx===
'GoTo Skipxx

DISABLE_PROFIBUS = True

If DISABLE_PROFIBUS Then
    HW_flagA = True
    HW_flagB = True
    HARDWARE_flag = False
ElseIf Not PTT And I_AM_HOST Then

    IO_ret_value = IO_Init(WINCARD_A, Status)
    
    If IO_ret_value = False Then
        If I_AM_HOST Then
            MsgBox "IO Card A failed to initialise", vbCritical, "Error"
        End If
        HW_flagA = True
        'Call ErrorMessage(Status)
    Else
        HW_flagA = False
    End If

    IO_ret_value = IO_Init(WINCARD_B, Status)
    
    If IO_ret_value = False Then
        If I_AM_HOST Then
            MsgBox "IO Card B failed to initialise", vbCritical, "Error"
        End If
        HW_flagB = True
        'Call ErrorMessage(Status)
    Else
        HW_flagB = False
    End If

  If HW_flagA Or HW_flagB Then
    HARDWARE_flag = False
  Else
    HARDWARE_flag = True
  End If
    
  'Dongle_Enable_Flag = Not HARDWARE_flag   ' IF there is NO HARWARE then we need the dongle
    
    'open databases for ai and di
    'MsgBox "Opening DI_AI_DATABASE", vbCritical, "Error"
    
    ' 9/11/99    .EXE DOES NOT WORK...
    '   OpenDI_AI_DataBase
    
    'MsgBox "Opened DI_AI_DATABASE", vbCritical, "Error"

    '1/12/99  Tell the model to use the Engine Room...
    status_flag = True
    Call set_STEAM_logicals(71, status_flag)

End If

' xxx
Skipxx:

' === 01/03/01 ===
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    Dongle_Enable_Flag = False
    
Else

    ' === 25/02/04 ===
    ' For BOTH types of dongle you need the NEXT LINE...
      If DISABLE_PROFIBUS Then
          Dongle_Enable_Flag = False
      Else
          Dongle_Enable_Flag = Not HARDWARE_flag   ' IF there is NO HARWARE then we need the dongle
      End If
    
    ' === 25/02/04 For Ray Gillet SINGLE USER, NO DONGLE , DLL protected(Greece V26.00 G)...
    ' Dongle_Enable_Flag = False
    
    
   ' === 15/03/01  =============
   'If This_PC_Name = "PC02" Then
   'If This_PC_Name = "HASP01" Then
    If This_PC_Name = "HASP01" And Not PTT Then
        Dongle_Enable_Flag = False
    End If

End If


' 15/03/01
test_label:

' 15/12/00 Dongle_Enable_Flag = False


    'Get IO Data from text files
    ErrorFlag = False

    
'    MsgBox "Calling GetIOData", vbCritical, "Error"

' We MUST GetIOData  ( This defines ALL the Valve and Breaker addresses needed for VB to work )
    Call GetIOData(ErrorFlag)
    
'    MsgBox "Calling GotIOData", vbCritical, "Error"
    
'  ==================

' 14/7/99
SI_units = False
mnuANSI.Checked = True
SI_units = False
Temp_char = "F"
Press_chars = " psi  "
Flow_chars = " lbs/hr"

Temp_mult = 1.8
Temp_const = 32#

Press_mult = 14.502     '     yyy psi = xxx Bar * 14.502 psi/Bar
Flow_mult = 7938#       '  yyy lbs/hr = xxx kg/sec * 2.205 lbs/kg  * 3600.0 secs/hour



' 14/5/99  FOR NOW...
mnu_dynamics.Visible = False

' 8/12/98
MDIform_loaded = False


'30/12/97 for now,
useEnglish = True

' 4/12/98
If TARGET_this_sim = PORTUGAL Then

   ' 2/9/99 mnuENGINE.Visible = False           ' Slow/Medium speed
    
    mnu_dynamics.Visible = False        ' Engine Dynamics
    mnu_configure.Visible = False       ' PC configuration
    mnuINaGROUP.Visible = False         '  Done in Introduction.frm
    mnu_ChangePassword.Visible = False  ' Needs more work?
    
    ' 10/12/98
    mnuOther.Caption = "Portuguese"
    mnuOther.Checked = True           ' For Portuguese
    mnuEnglish.Checked = False        ' For English
    useEnglish = False
    
End If

old_medium_speed = -123

'28/4/98
'''''Timer1.Enabled = False   'inhibit timer until all ints are complete



'24/4/98 make I/O run automatically

Me.Left = 0
Me.Top = 15000


' 20/10/98
               Sounds_Enabled = True
      mnuEnableSounds.Checked = True
      mnuSounds.Enabled = True
      
' 9/5/00
mnuUseEngineRoom.Visible = False


'11/10/99   If I_AM_HOST Then
If I_AM_INSTRUCTOR Then


'11/12/00
        MyStr = InstructorName
        MyStr = UCase(MyStr)
        
        
        ' === 7/01/02 ===
        If TARGET_this_sim = CALIFORNIA Then
            If MyStr <> "LLANELLI" Then
                mnuSoundMix.Enabled = False
            End If
        End If
        
        
      mnu_run.Enabled = True
      
    Call set_DI_board_default
    
 ''''' 27/4/98 Now done by Init_tasks.F90   Call INITIO
    
    mnu_instructor.Enabled = True
    
    ' 16/12/98
    If HARDWARE_flag Then
        mnuHardware.Enabled = True
        mnuHardware.Visible = True
    Else
        mnuHardware.Enabled = False
        mnuHardware.Visible = False
    End If
    
    
        ' === 19/11/02 ADDED for HA TEST then REMOVED again ...===
        'mnuHardware.Enabled = True
        'mnuHardware.Visible = True

    
        mnu_load_snapshot.Enabled = True
        mnu_save_snapshot.Enabled = True
    
        mnuHalt.Enabled = True
        mnuSounds.Enabled = True
        
    
    ' 24/4/98
     If split_mode Then
        Me.StatusBar1.Panels.Item(4).text = "SPLIT MODE"
     Else
        Me.StatusBar1.Panels.Item(4).text = "NORMAL MODE"
     End If
    
    
Else

            mnu_run.Enabled = True
      
    mnu_instructor.Enabled = False
    
        mnuHardware.Enabled = False
        mnuHardware.Visible = False
        
      ' === 23/01/02 15.036 Need snapshot facilty on client PC's when they are running "not in a group" ===
      '     mnu_load_snapshot.Enabled = False
      '     mnu_save_snapshot.Enabled = False
        If In_a_group Then      ' I am a STUDENT and IN A GROUP so I do NOT need snapshots....
            mnu_load_snapshot.Enabled = False
            mnu_save_snapshot.Enabled = False
        Else                    ' I am a STUDENT but I am NOT IN A GROUP and I need snapshots....
            mnu_load_snapshot.Enabled = True
            mnu_save_snapshot.Enabled = True
        End If
         
    
        mnuHalt.Enabled = False
        
    '27/11/99 just for now mnuSounds.Enabled = False

    
    ' 20/10/98
    '  mnuSounds.Enabled = False
      
End If

' 28/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    mnu_plotvars.Visible = False
    
    ' 29/4/00
    mnuSWBD11.Visible = False
    mnuSWBD13.Visible = False
    mnuEmergSWBD.Visible = False

    mnu_EOT_MECD.Visible = False
    
    mnu_Time.Visible = False
    
    mnuBurner_Management.Visible = False
    
Else

    mnu_plotvars.Visible = True
    
    ' 29/4/00
    mnuSWBD11.Visible = True
    mnuSWBD13.Visible = True
    mnuEmergSWBD.Visible = True
    
    mnu_EOT_MECD.Visible = True
    
    mnu_Time.Visible = True
    
    mnuBurner_Management.Visible = True

End If
    
    
''If medium_speed Then
''    Call Fill_DIText
''Else
''    Call Fill_LSS_DIText
''End If


'16/3/98    default host to 2 Tx and 1 Rx sockets
'17/6/98    default host to 9 Tx and 5 Rx sockets

'           default client to 1 Tx and 2 Rx sockets

If I_AM_HOST Then


    ' === 06/03/07 ===
    'Max_Host_sock_Tx = 9 'counting from 0
    'Max_Host_sock_Rx = 5 'counting from 0
    Max_Host_sock_Tx = 29 'counting from 0
    Max_Host_sock_Rx = 17 'counting from 0
    

    Max_Client_Sock_Tx = 0 'counting from 0
    Max_Client_Sock_Rx = 1 'counting from 0

Else
    Max_Host_sock_Tx = 1 'counting from 0
    Max_Host_sock_Rx = 0 'counting from 0

    Max_Client_Sock_Tx = 0 'counting from 0
    Max_Client_Sock_Rx = 1 'counting from 0

End If



' 6/3/98 PC_IN_CONTROL panel...
' 2/9/99 frm_Host.StatusBar1.Panels.Item(5).Text = ""



' 13/1/98
'Debug.Print "App.path = " & App.Path


' 29/4/98 PC_IN_CONTROL = -1       ' There is NO PC in control

'  15/1/98
'   WHEN running this program from HAL15 (via network)
'  App.Path points back to HAL_0 , which we don't want SO:
'Open App.Path & "\SIM.INI" For Input As #1

' 6/3/98
'On Error GoTo Errorhandler1





'Debug.Print "PC_name = " & PC_name & "     This_PC_number =" & This_PC_number


' 7/4/98


'        MsgBox "Initailaizing Sounds", vbCritical, "Error"


        frmHiddenSounds.Visible = False
'28/11/99 only initialize for PCs 0 & 1
        If This_PC_number = 0 Or This_PC_number = 1 Then
            frmHiddenSounds.InitialiseSounds
        End If

'================= medium_speed ( or slow speed ) =============================

'   Default to MEDIUM SPEED
'   medium_speed = True
'   mnuSLOWSPEED.Checked = False
'   mnuMEDIUM.Checked = True

'   Default to SLOW SPEED (for testing)
'   medium_speed = False

'If medium_speed Then
'    mnuSLOWSPEED.Checked = False
'       mnuMEDIUM.Checked = True
'Else
'    mnuSLOWSPEED.Checked = True
'       mnuMEDIUM.Checked = False
'End If

If medium_speed Then

  ' 18/8/98
   mnuTurbo.Enabled = False
   
   mnuCAMSHAFT.Enabled = False
   mnuCAMSHAFT.Visible = False
     
     mnuHOSEP2.Visible = False
      mnuHOSEP.Caption = "HO Separator"
     
    mnuAIRCOMP.Visible = False

    mnuBoiler.Visible = False
    
' 11/6/98
    mnuGBOX.Visible = True
    mnuFVCW.Visible = True
    
Else

  ' 18/8/98
   mnuTurbo.Enabled = True
   
    mnuCAMSHAFT.Enabled = True
   mnuCAMSHAFT.Visible = True
     
     mnuHOSEP2.Visible = True
      mnuHOSEP.Caption = "HO Separator No1"
  
    mnuAIRCOMP.Visible = True
    
    mnuBoiler.Visible = True
    
' 11/6/98
    mnuGBOX.Visible = False
    mnuFVCW.Visible = False

End If




' === 1/9/99 For CALIFORNIA ===

' === 17/05/01 ===
'  JBF PHONE CALL 15:49
'If TARGET_this_sim = CALIFORNIA Then

    mnu_main_engine.Visible = False
    mnuLubrication.Visible = False
    mnuCooling.Visible = False
    mnuCPP.Visible = False
    mnuFuel.Visible = False
    mnuCOMPAIR.Visible = False
    
    ' 11/7/00
 '   mnuFire.Visible = False
     mnuFire.Visible = True
    
    mnuBilge.Visible = False
    MNUFWG.Visible = False
    mnuBoiler.Visible = False
    
    mnu_student.Visible = True
    mnu_instructor.Visible = True
   ' 2/9/99 mnuENGINE.Visible = False
    mnuLanguage.Visible = False
    
'End If



'==============================================================================

''If medium_speed Then
' 22/12/97 OPEN a com port for FORTRAN
   '' Call open_com(1)
''Else
' 22/5/98
    ''Call open_com_LSS(1)
''End If

' 23/12/97
'alarm_height = 10700
'alarm_width = 6700
' 10/4/98
'alarm_height = 12030
'alarm_width = 6800
'12/10/99


'11/8/00
'alarm_height = 7000
'alarm_width = 8500
alarm_height = 9750
alarm_width = 8670

' 11/3/98
elec_ctrl_height = 4875
elec_ctrl_width = 4470

' 17/8/98 JYA
'InitialiseScreens
    
' =========== IMPORTANT PATCH ================================================
' 10/4/98   A "PATCH" to make EXECUTIVE.EXE run OK
' It was found that the .exe version would "crash" with an Automation error
' if SY030 or SY200 is "opened" before opening either SY080 or SY205 or SY206
' The reason for this is unknown at present ( 10/4/98 )
' So try this "PATCH"   ( NB TEST EXECUTIVE.EXE )
'
'   TESTED EXECUTIVE.EXE and it now works OK
'   It is now possible to open SY030 BEFORE opening SY080
'
'Load sy080
'sy080.Visible = False
' =========== IMPORTANT PATCH END ============================================


' 10/4/98   NB.  See MDIform QueryUnload

    
    
' Fill up the alarmtxt array
'alarm_no = 2


' 23/4/98    split_mode = True   ' SPLIT MODE



diesel_init = True
elec_init = True
diesel_cold = True
elec_cold = True

' 8/12/98  18:45
'Open "c:\RWJ_TEST.txt" For Append As #1
'Write #1, " Point frm_HOST  A7a "
'Close #1    ' Close file.

' === 23/01/02 ===
'If medium_speed Then
'    Call D_MODEL(split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
'Else

        ' 8/12/98  18:45
   '     Open "c:\RWJ_TEST.txt" For Append As #1
   '     Write #1, " Point frm_HOST  A7d "
   '     Close #1    ' Close file.

    Call INIT_LSS    ' ( INITIALI.F90 )
    
    
   ' === 06/03/03 ===
   ' Call D_MODEL_LSS(CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
     Call D_MODEL_LSS(Ballast_flag, CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    
    ' 8/12/98 18:10
   ' Open "c:\RWJ_TEST.txt" For Append As #1
   ' Write #1, " Point frm_HOST  A7e "
   ' Close #1    ' Close file.
    
'End If



' 8/12/98 18:10
'Open "c:\RWJ_TEST.txt" For Append As #1
'Write #1, " Point frm_HOST  A7f "
'Close #1    ' Close file.

' 23/4/98   split_mode = True

d_run = True
e_run = True
diesel_init = False
elec_init = False
diesel_cold = False
elec_cold = False

' === 23/01/02 ===
'If medium_speed Then
'   Call D_MODEL(split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
'Else
    
   ' === 06/03/03 ===
   'Call D_MODEL_LSS(CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    Call D_MODEL_LSS(Ballast_flag, CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)

'End If

'=========================================

'MsgBox "Calling Getting alarm text", vbCritical, "Error"

Call get_alarm_text

If medium_speed Then
  ' 10/3/98 added SYNCHSELSW etc
  Call get_elec_addresses(ESTARTSW(1), _
  ESTOPSW(1), EAUTOSW(1), EMANSW(1), ESTBY1SW(1), ESTBY2SW(1), _
  AVRAUTOSW(1), AVRMANSW(1), MBKSW(1), OPENMBKSW(1), AVR_pot(1), _
  SYNCHSELSW(1), CHSYNCHSW, SYNCHOFFSW, SELLAMPSSW, SELSYNCHSW, _
  REVRESET(1))

  Call get_bkr_addresses(BKR(1, 1))
  Call get_V_addresses(DI_VIO(1), DO_VIO(1))
  Call get_VV_addresses(DI_ON_VVIO(1), DI_OFF_VVIO(1))
  Call get_P_addresses(PumpIO(1, 1))
  Call get_RMA_addresses(RMAIO(1, 1))
  
Else
  Call get_VV_addr_LSS(DI_ON_VVIO_LSS(1), DI_OFF_VVIO_LSS(1))
  ' 26/5/98...
  Call get_BP_addr(PumpIO_LSS(1, 1))
  Call get_RMA_addr_LSS(RMAIO_LSS(1, 1))
  
  ' 5/6/98 ADDED:
    Call get_elec_addr_LSS(ESTARTSW(1), _
  ESTOPSW(1), EAUTOSW(1), EMANSW(1), ESTBY1SW(1), ESTBY2SW(1), _
  AVRAUTOSW(1), AVRMANSW(1), MBKSW(1), OPENMBKSW(1), AVR_pot(1), _
  SYNCHSELSW(1), CHSYNCHSW, SYNCHOFFSW, SELLAMPSSW, SELSYNCHSW, _
  REVRESET(1))
  ' 5/6/98 ADDED:
  Call get_bkr_addr_LSS(BKR(1, 1))
 
End If


        ss_supply = False
      diesel_init = False
        elec_init = False
      diesel_cold = False
        elec_cold = False
            d_run = False
            e_run = False
        snap_shot = False
          hw_test = False

' 19/1/98
 '  test_flag = True
   test_flag = False

Me.StatusBar1.Panels.Item(1).text = Op_Mode_Name
'8/5/97
'StatusBar1.Panels.Item(6).Text = "V1.0"

' 13/3/98
' 12/12/99  ( Can we use this ? ....)

' === 13/03/01 Remove the "And Not PTT" bit because I can't run netwoked PTT at HA ===
'If I_AM_HOST And Not PTT Then
If I_AM_HOST Then

'MsgBox "Setting up Rx Sockets", vbCritical, "Error"

' 25/6/98
   If In_a_group Then
   
     ' Receive on HostSockRx
     For i = 0 To Max_Host_sock_Rx
        frm_Host.HostSockRx(i).AddressFamily = AF_INET
        frm_Host.HostSockRx(i).Protocol = IPPROTO_IP
    
        frm_Host.HostSockRx(i).SocketType = SOCK_STREAM
    
        frm_Host.HostSockRx(i).Blocking = False
        frm_Host.HostSockRx(i).Binary = False
        frm_Host.HostSockRx(i).BufferSize = 32767
        'frm_Host.HostSockRx(i).LocalPort = Host_Tx_Base + i ' should be base + 2*PC number + i

        frm_Host.HostSockRx(i).LocalPort = Host_Rx_Base + i ' HostRx -> ClientTx is 1:1
        frm_Host.HostSockRx(i).Action = SOCKET_LISTEN
     Next
     
    End If
    


' 25/6/98
'   If Not test_flag Then
    If In_a_group Then
    
'MsgBox "Setting up Tx Sockets", vbCritical, "Error"

        ' 22/1/97
        ' Transmit on Host_sock_Tx (0 To Max_Host_sock_Tx)
        For i = 0 To Max_Host_sock_Tx

            ' === 07/03/07 ===
            SOCKET_DO_IT = False

            Select Case i
            
            Case 0, 1 'PC01
            
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC01
                If CLIENT_PC_IS_ON_LINE(1) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC01
                    SOCKET_DO_IT = True
                End If
                
                
            Case 2, 3 'PC02
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC02
                If CLIENT_PC_IS_ON_LINE(2) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC02
                    SOCKET_DO_IT = True
                End If
            Case 4, 5 'PC03
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC03
                If CLIENT_PC_IS_ON_LINE(3) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC03
                    SOCKET_DO_IT = True
                End If
            Case 6, 7 'PC04
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC04
                If CLIENT_PC_IS_ON_LINE(4) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC04
                    SOCKET_DO_IT = True
                End If
            Case 8, 9 'PC05
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC05
                If CLIENT_PC_IS_ON_LINE(5) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC05
                    SOCKET_DO_IT = True
                End If
                
                
            ' === 07/03/07 ===
            Case 10, 11 'PC06
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC02
                If CLIENT_PC_IS_ON_LINE(6) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC06
                    SOCKET_DO_IT = True
                End If
            Case 12, 13 'PC07
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC07
                If CLIENT_PC_IS_ON_LINE(7) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC07
                    SOCKET_DO_IT = True
                End If
            Case 14, 15 'PC08
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC08
                If CLIENT_PC_IS_ON_LINE(8) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC08
                    SOCKET_DO_IT = True
                End If
            Case 16, 17 'PC09
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC09
                If CLIENT_PC_IS_ON_LINE(9) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC09
                    SOCKET_DO_IT = True
                End If
            Case 18, 19 'PC10
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC10
                If CLIENT_PC_IS_ON_LINE(10) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC10
                    SOCKET_DO_IT = True
                End If
            Case 20, 21 'PC11
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC11
                If CLIENT_PC_IS_ON_LINE(11) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC11
                    SOCKET_DO_IT = True
                End If
            Case 22, 23 'PC12
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC12
                If CLIENT_PC_IS_ON_LINE(12) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC12
                    SOCKET_DO_IT = True
                End If
            Case 24, 25 'PC13
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC13
                If CLIENT_PC_IS_ON_LINE(13) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC13
                    SOCKET_DO_IT = True
                End If
            Case 26, 27 'PC14
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC14
                If CLIENT_PC_IS_ON_LINE(14) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC14
                    SOCKET_DO_IT = True
                End If
            Case 28, 29 'PC15
                ' === 07/03/07 ===
                'frm_Host.Host_sock_Tx(i).HostAddress = HostAddrPC15
                If CLIENT_PC_IS_ON_LINE(15) Then
                    Me.Host_sock_Tx(i).HostAddress = HostAddrPC15
                    SOCKET_DO_IT = True
                End If
                
                
                
     
           End Select
            

              ' === 07/03/07 ===
              If SOCKET_DO_IT Then
              
                frm_Host.Host_sock_Tx(i).AddressFamily = AF_INET
                frm_Host.Host_sock_Tx(i).Protocol = IPPROTO_IP
                frm_Host.Host_sock_Tx(i).SocketType = SOCK_STREAM
                frm_Host.Host_sock_Tx(i).RemotePort = Host_Tx_Base + i
                frm_Host.Host_sock_Tx(i).BufferSize = 32767
                frm_Host.Host_sock_Tx(i).Blocking = False
                
            End If
            
        Next i
   
    End If


' 13/3/98
End If


'tasks running =======
'    If (Not tasks_running) Then
         task_array(1) = True
         task_array(2) = False
         
'''''' 10/12/97
' 27/4/98 Put back in:
'11/5/98 don't run io or model if a client
    If I_AM_HOST Then
    
    ' 7/7/98
      If medium_speed Then  ' Init_tasks.F90 is in MODEL.DLL
      
         Call Init_Tasks(task_array(1))
         tasks_running = True
       
      End If
      
    Else
       Call Client_StartUp
    End If

'    Else
'        Debug.Print "PCP tried to INIT_TASKS with the tasks already RUNNING"
'    End If
' =======================================================


'IOtest.Visible = False
'IOtest.zorder
'IOtest.Visible = False


' === 07/06/02 =========================

FileNo = FreeFile
Open "c:\STEAM_SIM\text\MECD_EOT.txt" For Input As #FileNo

For i = 1 To 10
    Input #FileNo, IN_string
    EOT_instructor(i) = Val(IN_string)
Next i

Close #FileNo

Exit Sub

' ======================================




Exit Sub

' 8/12/98
MDIform_loaded = True

Errorhandler1:

msg = "There was an error reading the Sim.ini file"   ' Define message.
Style = vbOK + vbCritical ' Define buttons.
Title = "Initialization Error"  ' Define title.
'Help = "DEMO.HLP"   ' Define Help file.
'Ctxt = 1000 ' Define topic
    
Response = MsgBox(msg, Style, Title)
Close #1
Exit Sub


End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dim Pcolour As Long



'' 28/1/97
'mouse_down = True
''Debug.Print "HOST form mouse_down = true "
'Debug.Print "MDIform MouseDown x,y =   " & X & "     " & Y
'Xval = X
'Yval = Y

'Pcolour = sy030.Picture1.Point(Xval, Yval)
'sy030.Text1.Text = "Pcolour = " & Hex(Pcolour) & "       " & Xval & "    " & Yval


End Sub


Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' 28/1/97
mouse_down = False
'Debug.Print "HOST form mouse_down = false "


End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'prevent accidental exit
Dim Rsp As Integer
          
          Debug.Print "MDIform QueryUnload event"
  
'20/11/97
If reset_flag Then

        reset_flag = False

        ' 1/12/98
          frm_intro.ZOrder
      '   frm_intro_low_res.zorder
        
        Cancel = -1
          
        Debug.Print "MDIform QueryUnload , reset_flag = true "
        
Else
  '19/11/97
   Rsp = MsgBox("Do you really want to stop and return to Windows ?", 20, "WARNING")
  ' This is ICELAND code:
  'Rsp = MsgBox(Main_strings(60), 20, Main_strings(59))

'      Debug.Print "response = " & Rsp

      If Rsp <> vbYes Then
          Cancel = -1
          
          Debug.Print "MDIform cancel QueryUnload"

            
          Exit Sub
      End If
End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim IO_ret_value As Boolean



Debug.Print "MDIform Unload event"
Debug.Print " Shuting down : MDIform and frm_intro"


' 3/9/99


' === 23/11/00 ===

If Not PTT And Not DISABLE_PROFIBUS Then

 IO_ret_value = IO_Exit(WINCARD_A, Status)
 IO_ret_value = IO_Exit(WINCARD_B, Status)
 
End If



' 6/9/99 apm
'close database
  CloseDI_AI_Database
' 25/6/98
If In_a_group Then

  For i = 0 To Max_Host_sock_Tx
      If frm_Host.Host_sock_Tx(i).Connected Then frm_Host.Host_sock_Tx(i).Action = SOCKET_CLOSE
  Next i

  For i = 0 To Max_Host_sock_Rx
      If frm_Host.HostSockRx(i).Listening Then frm_Host.HostSockRx(i).Action = SOCKET_CLOSE
  Next i

End If

' 1/12/98
Unload frm_intro
'Unload frm_intro_low_res
  
' 28/4/98  HALT the IO task....
'Call setPCDI(2568)
VB_RunIO = False
Call set_RunIO(VB_RunIO)


Sleep (1000)
  
' 27/4/98 halt the task
task_array(1) = False
task_array(2) = True
Call Init_Tasks(task_array(1))
  
RunIO = False

Sleep (1000)
  
  
'24/4/98
RunIO = False
If I_AM_HOST Then
    Call free_io_board ' release the board handle
End If

If LogFileOpen Then
    Close #LogFileNumber
    LogFileOpen = False
End If


End Sub


Private Sub MSComm1_OnComm()

End Sub



Private Sub mnu_about_Click()

Winsock_udp.ZOrder

'    frmAbout.zorder
    
    
End Sub


Private Sub mnu_actions_Click()

frm_Student_Actions.ZOrder

End Sub

Private Sub mnu_AnalogTest_Click()
HWTestAnalog.ZOrder
End Sub

Private Sub mnu_Air_controller_PORT_Click()

Controller_AIR_PORT.ZOrder

End Sub

Private Sub mnu_Air_controller_STBD_Click()

Controller_AIR_STBD.ZOrder

End Sub

Private Sub mnu_Aux_Steam_Alone_Click()

' 09/06/09
Aux_Steam_Alone.ZOrder

End Sub

Private Sub mnu_Back_Colour_Click()

' === 31/01/02 ===

   ' Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo Errhandler
   
   ' Set the Flags property.
   CommonDialog1.Flags = cdlCCRGBInit
   ' Display the Color dialog box.
   CommonDialog1.ShowColor
   
   ' Set the form's background color to the selected
   ' color.
   'Form1.BackColor = CommonDialog1.Color
   
   
   Bcolour = CommonDialog1.Color
   
   
    ' NB Looks daft but this NEEDS TO BE DONE....
    'HavenStretch1.Visible = False
   
  ' HavenStretch1.Back_color = Bcolour
   
  ' Call copy_havenstretch_colours(Me, Bcolour)

    ' NB Looks daft but this NEEDS TO BE DONE....
   ' HavenStretch1.Visible = True
  
   
   Exit Sub

Errhandler:
   ' User pressed Cancel button.
   Exit Sub



End Sub

Private Sub mnu_Back_Grey_Click()

' === 31/01/02 ===
   Bcolour = B_colour_default
     
End Sub

Private Sub mnu_Back_White_Click()

' === 31/01/02 ===
   Bcolour = vbWhite    ' WHITE
   
   
End Sub

Private Sub mnu_Ballast_Click()

' === 31/05/02 ===
Ballast_BULK.ZOrder
' === 20/11/02 17.011 ===
Ballast_BULK.WindowState = 0 ' Normal

End Sub

Private Sub mnu_Bilge_OWS_Click()


Bilge_OWS.ZOrder
' === 20/11/02 17.011 ===
Bilge_OWS.WindowState = 0 ' Normal

End Sub

Private Sub mnu_BNR_Port_Click()

BRNR_MANAGEMENT_PORT.ZOrder
' === 20/11/02 17.011 ===
BRNR_MANAGEMENT_PORT.WindowState = 0 ' Normal

End Sub

Private Sub mnu_Bnr_stbd_Click()

BRNR_MANAGEMENT_STBD.ZOrder
' === 20/11/02 17.011 ===
BRNR_MANAGEMENT_STBD.WindowState = 0 ' Normal


End Sub

Private Sub mnu_boilportV2_Click()

Port_Boiler.ZOrder

End Sub

Private Sub mnu_BoilStbdV2_Click()


End Sub

Private Sub mnu_boilport_Click()

' 11/11/99
Port_Boiler.ZOrder
' === 20/11/02 17.011 ===
Port_Boiler.WindowState = 0 ' Normal

End Sub

Private Sub mnu_BoilStbd_Click()

' 11/11/99
Stbd_Boiler.ZOrder
' === 20/11/02 17.011 ===
Stbd_Boiler.WindowState = 0 ' Normal

End Sub

Private Sub mnu_ChangePassword_Click()
    frm_passwwd_change.ZOrder
End Sub

Private Sub mnu_cold_Click()
Dim i As Integer
Dim IN_string As String


' =========================
' === 06/01/03 ===

Dim ErrorFlag As Byte

ErrorFlag = False


' === 05/04/03 Move from below ===
Me.MousePointer = ccHourglass

  
Call GetIOData(ErrorFlag)

' =========================




' === 23/08/02 ===
Call SET_Constants

'' === 04/03/02 =======
'On Error GoTo ErrH
'
'FileNo = FreeFile
'Open "c:\STEAM_SIM\text\CONSTANTS.txt" For Input As #FileNo
'
'    Input #FileNo, IN_string
'    Shaft_Inertia_multiplier = Val(IN_string)
'    i = 1
'    Call set_CONSTANT(i, Shaft_Inertia_multiplier)
'
'    Input #FileNo, IN_string
'    Ship_Inertia_multiplier = Val(IN_string)
'    i = 2
'    Call set_CONSTANT(i, Ship_Inertia_multiplier)
'
'    Input #FileNo, IN_string
'    Pitch_multiplier = Val(IN_string)
'    i = 3
'    Call set_CONSTANT(i, Pitch_multiplier)
'
'Close #FileNo
'
' === 04/03/02 END ===


'21/07/00
Call setFAULT(3, 1)
FaultScheduleIndex = 0


' 21/7/00  To ensure frm_faults updates correcly...
Unload frm_faults



runIO_counter = 0

' === 05/04/03 Move UP ===
'Me.MousePointer = ccHourglass



' 11/12/98
cold_start_hourglass = True

' 9/8/00 ADDED to set EXECI( )
Call INIT_LSS    ' ( INITIALI.F90 )

diesel_init = True
elec_init = True

diesel_cold = True
elec_cold = True


Sim_Running_Status = "Running"

Me.StatusBar1.Panels.Item(2).text = "Cold Start"

' 13/10/09
SnapNum = 0

RunIO = True
mnuHalt.Enabled = True
mnuContinue.Enabled = False


d_run = True
e_run = True


VB_RunIO = True
Call set_RunIO(VB_RunIO)




'  21/7/98 RWJ  update all sliders:
update_VBAIs = True


'  9/5/00  ==== See 5.102 and 5.100 S Green  See MODEL_IN.F90 ====
'    nValue = 100   ' Finished with engines
'   ' Call set_AO_LSS(131, EOTctl1.TelegraphValue)
'    Call set_AO_LSS(131, nValue)
'
'    nValue = 127   ' middle position (STOP)
'   ' Call set_AO_LSS(130, EOTctl1.ThrottleVAlue)
'     Call set_AO_LSS(130, nValue)


'08/09/00
ResetAllAlarms = True


Exit Sub

' === 23/08/02 ===
'ErrH:
'resp = MsgBox("There was an error reading c:\STEAM_SIM\text\CONSTANTS.txt" & Chr(10) & _
'       "You must ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
'       Stop


End Sub

Private Sub mnu_COMBUST_CONTROL_Click()


' === 5/03/02 ===
Controllers_combustion.ZOrder


End Sub

Private Sub mnu_configure_Click()
    frm_config.ZOrder
End Sub

Private Sub mnu_digitalTest_Click()
    HWTestDigital.ZOrder
End Sub

Private Sub mnu_Control_Air_Click()

Control_Air_System.ZOrder
' === 20/11/02 17.011 ===
Control_Air_System.WindowState = 0 ' Normal

End Sub

Private Sub mnu_Controller_status_Click()


' === 23/01/02 ===
Controller_Status.ZOrder


End Sub

Private Sub mnu_DISABLE_ACT_STUDENT_Action_DISPLAY_Click()

'c  === 28/03/01 POint 10.29 ===
'c  LSS_LOGICALS(63) = Reserved ( See in VB it is used as DISABLE_ACT_student_action_display )
'c  LSS_LOGICALS(64) = Reserved ( See in VB it is used as DISABLE_ACT_student_actions )

If LSS_LOGICALS(63) Then
  status_flag = False
Else
  status_flag = True
End If
Call set_LSS_logicals(63, status_flag)


End Sub

Private Sub mnu_DISABLE_ACT_STUDENT_Actions_Click()


'c  === 28/03/01 POint 10.29 ===
'c  LSS_LOGICALS(63) = Reserved ( See in VB it is used as DISABLE_ACT_student_action_display )
'c  LSS_LOGICALS(64) = Reserved ( See in VB it is used as DISABLE_ACT_student_actions )

If LSS_LOGICALS(64) Then
  status_flag = False
Else
  status_flag = True
End If
Call set_LSS_logicals(64, status_flag)


End Sub

Private Sub mnu_Drum_PORT_Click()

' 26/9/99
Controller_Drum_level_PORT.ZOrder

End Sub

Private Sub mnu_Drum_STBD_Click()

' 26/9/99
Controller_Drum_level_STBD.ZOrder

End Sub

Private Sub mnu_dynamics_Click()
    frm_Plot2.ZOrder
End Sub

Private Sub mnu_endlog_Click()
Dim resp As Integer

 If LogFileOpen Then
    LogFileOpen = False
    
    Close #LogFileNumber
        mnu_startlog.Enabled = True
        mnu_endlog.Enabled = False
        'resp = MsgBox("The log file named " & LogFileName & " has been closed " & Chr(20) _
               '& "Do you want to view it now ?", vbYesNo, "Logging")
        resp = MsgBox("The log file named " & LogFileName & " has been closed " & Chr(10) _
         & "You can view it now by selecting the Log View option", vbOKOnly, "Logging")

    If resp = vbYes Then
    
    
    End If
    
 End If



End Sub

Private Sub mnu_EOT_MECD_Click()


    Frm_EOT_MECD.ZOrder


End Sub

Private Sub mnu_exit_Click()

'24/4/98
RunIO = False



' 10/4/98


Unload Me

' 1/12/98
Unload frm_intro
'Unload frm_intro_low_res


End Sub

Private Sub mnu_FAST_Click()

    status_flag = True

' Tell the model to go to fast time...
Call set_STEAM_logicals(5, status_flag)

End Sub

Private Sub mnu_faults_Click()
    frm_faults.ZOrder
End Sub

Private Sub mnu_fortran_conv_Click()

'  test_model.zorder
   comp_air.ZOrder    ' TEST FORM

End Sub

Private Sub mnu_Feed_alone_Click()

' 09/06/09
Feed_Water_Alone.ZOrder

End Sub

Private Sub mnu_feed_pump_PORT_Click()

Controller_FEED_PUMP_PORT.ZOrder

End Sub

Private Sub mnu_feed_pump_STBD_Click()

Controller_FEED_PUMP_STBD.ZOrder

End Sub

Private Sub mnu_Fuel_Air_PORT_Click()

Controller_FUEL_AIR_PORT.ZOrder

End Sub

Private Sub mnu_Fuel_Air_STBD_Click()

Controller_FUEL_AIR_STBD.ZOrder

End Sub

Private Sub mnu_Fuel_controller_PORT_Click()

Controller_FUEL_PORT.ZOrder

End Sub

Private Sub mnu_Fuel_controller_STBD_Click()

Controller_FUEL_STBD.ZOrder

End Sub

Private Sub mnu_instructor_Click()
    'If PasswordSet Then
       ' frm_passwd.zorder
   'Else
      '  frm_password_change.zorder
   ' End If
    
End Sub

Private Sub mnu_Landscape_Click()



Dim original_height As Long
Dim original_width As Long
Dim original_window_state As Long
Dim original_Stretchable_screens As Boolean


On Error GoTo error_handler



' Set the Printer to LANDSCAPE to make BEST use of the paper:
Printer.Orientation = vbPRORLandscape


' Record the Screen dimensions
original_height = Screen.ActiveForm.Height
original_width = Screen.ActiveForm.Width


' === Not needed for CMA
'original_Stretchable_screens = STRETCHABLE_SCREENS
'STRETCHABLE_SCREENS = True


' Record the Screen window state
original_window_state = Screen.ActiveForm.WindowState
Screen.ActiveForm.WindowState = 0   ' Normal


' Make BEST use of the paper :
' FORCE the Screen to be the SAME size as the Printer ( so it FILLS the printed page up completely )
  Screen.ActiveForm.Height = Printer.Height
  Screen.ActiveForm.Width = Printer.Width

' === 31/01/02 ===
Screen.ActiveForm.HavenStretch1.Back_color = vbWhite
Bcolour = vbWhite
Call copy_havenstretch_colours(Screen.ActiveForm, Bcolour)


' PRINT the Screen
Screen.ActiveForm.PrintForm

' === 31/01/02 ===
Screen.ActiveForm.HavenStretch1.Back_color = &HC0C0C0       ' GREY
Bcolour = &HC0C0C0                  ' GREY
Call copy_havenstretch_colours(Screen.ActiveForm, Bcolour)

' To ensure the printer can print MORE than once:
Printer.EndDoc


' === Not needed for CMA
'STRETCHABLE_SCREENS = original_Stretchable_screens



' Restore the Screen dimensions
Screen.ActiveForm.Height = original_height
Screen.ActiveForm.Width = original_width

' Restore the Screen Window state
Screen.ActiveForm.WindowState = original_window_state


Exit Sub

error_handler:
Response = MsgBox("Cannot Print , Ensure a display is selected , " & Err.Description, vbOK, "Printer Error")


End Sub

Private Sub mnu_LO_Purifier_Click()


' === 4/12/01 14.072 ===
CMA_purifier.ZOrder
' === 20/11/02 17.011 ===
CMA_purifier.WindowState = 0 ' Normal


End Sub

Private Sub mnu_load_snapshot_Click()



'23/4/98
RunIO = False

TakeSnapFlag = False
Load SnapForm

Me.StatusBar1.Panels.Item(2).text = "Snapshot Start"

'23/4/98
RunIO = True

' 15/10/98 MOVE to below....
'mnuHalt.Enabled = True
'mnuContinue.Enabled = False

' 28/4/98  RUN the IO ....
'nPCDI = 2567
'Call setPCDI(nPCDI)


VB_RunIO = True
Call set_RunIO(VB_RunIO)

'28/4/98
'''''Timer1.Enabled = True

' 29/4/98

' 15/10/98
If DEMO_flag Then
    mnuHalt.Enabled = False
    mnuContinue.Enabled = False
Else
    mnuHalt.Enabled = True
    mnuContinue.Enabled = False
End If
    '29/11/99 only start after snapshot loaded. So do this on Snap Form
    'd_run = True
    'e_run = True

End Sub

Private Sub mnu_MASTER_PORT_Click()

Controller_MASTER_PORT.ZOrder

End Sub

Private Sub mnu_MASTER_STBD_Click()

Controller_MASTER_STBD.ZOrder

End Sub

Private Sub mnu_MECD_Click()

' === 27/06/01 ==
'MECD_control.zorder

End Sub

Private Sub mnu_Normal_Click()

    status_flag = False

' Tell the model to go to NORMAL time...
Call set_STEAM_logicals(5, status_flag)

End Sub

Private Sub mnu_parameters_Click()

' 16/7/98
If medium_speed Then
   Instrp.ZOrder
Else
   Instrp_LSS.ZOrder
End If

End Sub

Private Sub mnu_PLANT_MASTER_Click()


Controller_PLANT_MASTER.ZOrder

End Sub

Private Sub mnu_plotvars_Click()

   ' frm_Plot1.zorder
    '11/09/00
    GraphFrm.ZOrder
End Sub

Private Sub mnu_Portrait_Click()

Dim ans As Integer


' === 10/04/01 ===
Dim Screen_H_W_RATIO As Single


Dim original_height As Long
Dim original_width As Long
Dim original_window_state As Long
Dim original_Stretchable_screens As Boolean



' === 30/11/01 ======
'
' The code WAS in mnuPrint and was this....
'' Print the form that has the focus
'On Error GoTo ErrH:
'
'ActiveForm.PrintForm
'
'Exit Sub
'
'ErrH:
'ans = MsgBox("There is nothing to print", vbOKCancel, "Print Error", 0, 0)
'
' ===================



On Error GoTo error_handler

' Set the Printer to Portrait
Printer.Orientation = vbPRORPortrait

' Record the Screen dimensions
original_height = Screen.ActiveForm.Height
original_width = Screen.ActiveForm.Width

If original_width > 1 Then
  Screen_H_W_RATIO = original_height / original_width
Else
  Screen_H_W_RATIO = 1#
End If


' 16/11/00 NOTE: See d:\VIGO to use the flag STRETCHABLE_SCREENS ( It's an Instructor Option )
' It is included here for future use
' If you want Stretcahble screens / BIG Screens to be selectable see d:\VIGO
'
' Don't need for CMA...
'original_Stretchable_screens = STRETCHABLE_SCREENS
'STRETCHABLE_SCREENS = True


' Record the Screen window state
original_window_state = Screen.ActiveForm.WindowState
Screen.ActiveForm.WindowState = 0   ' Normal


' Make BEST use of the paper :

' === 10/04/01 ===
' This is PORTRAIT so do this...
'  Screen.ActiveForm.Height = Printer.Height
'  Screen.ActiveForm.Width = Printer.Width

  Screen.ActiveForm.Width = Printer.Width
  Screen.ActiveForm.Height = Printer.Width * Screen_H_W_RATIO


' === 28/01/02 ====
' Wanted a white system colour to save ink....
' Next to lines did sod all so took out...
'Printer.FillColor = vbWhite
'Printer.DrawMode = vbWhiteness

' === 28/01/02 ===
' Also didn't work ( Havenstretch is not transparent )
'Screen.ActiveForm.BackColor = vbWhite
'Screen.ActiveForm.FillColor = vbWhite
'Screen.ActiveForm.ForeColor = vbWhite

' === 31/01/02 ===
Screen.ActiveForm.HavenStretch1.Back_color = vbWhite
Bcolour = vbWhite
Call copy_havenstretch_colours(Screen.ActiveForm, Bcolour)


' PRINT the Screen
Screen.ActiveForm.PrintForm


' To ensure the printer can print MORE than once:
'Printer.EndDoc

' === 31/01/02 ===
Screen.ActiveForm.HavenStretch1.Back_color = &HC0C0C0       ' GREY
Bcolour = &HC0C0C0                  ' GREY
Call copy_havenstretch_colours(Screen.ActiveForm, Bcolour)


' Don't need for CMA...
'STRETCHABLE_SCREENS = original_Stretchable_screens


' Restore the Screen dimensions
Screen.ActiveForm.Height = original_height
Screen.ActiveForm.Width = original_width

' Restore the Screen Window state
Screen.ActiveForm.WindowState = original_window_state


Exit Sub


error_handler:
Response = MsgBox("Cannot Print , Ensure a display is selected : " & Err.Description, vbOK, "Printer Error")


End Sub

Private Sub mnu_real_PID_Click()

' 7/4/98
Real_PID.ZOrder

End Sub

Private Sub mnu_save_snapshot_Click()
TakeSnapFlag = True
Load SnapForm

End Sub

Private Sub mnu_SLOW_Click()

'  10/8/00
    status_flag = True

' Tell the model to go to SLOW time...
Call set_STEAM_logicals(121, status_flag)


End Sub

Private Sub mnu_startlog_Click()
Dim TempStr As String

LogString = " "

LogFileNumber = FreeFile   ' Get unused file

' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler


    LogFileName = "L" & Format(Time, "hhmm") & "_" & Format(Date, "ddd_d_mm_yyyy") & ".log"
    CommonDialog1.DialogTitle = "Log Name Selection"
    CommonDialog1.Filename = LogFileName

    Debug.Print "LogFileName = " & LogFileName
    CommonDialog1.InitDir = "c:\Steam_Sim" & "\logs\"
    CommonDialog1.Filter = "Log Files (*.log)|*.log|"


' Display the Open dialog box
    CommonDialog1.ShowOpen

' Display name of selected file
    Open CommonDialog1.FileTitle For Output Shared As #LogFileNumber

    'Print #LogFileNumber, "This is a test of the log file"
    'LogString = LogString & "This is a test of the log file saved in LogString"

    mnu_startlog.Enabled = False
    mnu_endlog.Enabled = True
    mnu_viewlog.Enabled = True
    LogFileOpen = True
    
'Print Main Title & Info

    '23/11/99
    GoTo Skip
    
    'Call Logging("C.M.A. STEAM PLANT SIMULATOR LOG FILE", LogTitle)
     Call Logging("STEAM PLANT SIMULATOR LOG FILE", LogTitle)
     
    Call Logging(" ", LogTitle)
    TempStr = "DATE: " & Format(Date, "dd:mm:yyyy")
    Call Logging(TempStr, LogTitle)
    TempStr = "START TIME: " & Format(Time, "hh:mm:ss")
    Call Logging(TempStr, LogTitle)
    Call Logging(" ", LogTitle)
    TempStr = "INSTRUCTOR: " & InstructorName
    Call Logging(TempStr, LogTitle)
    TempStr = "CLASS: " & ClassName
    Call Logging(TempStr, LogTitle)
    TempStr = "COURSE: " & CourseName
    Call Logging(TempStr, LogTitle)
    TempStr = "STUDENT: " & StudentName
    Call Logging(TempStr, LogTitle)
    Call Logging(" ", LogTitle)
    Call Logging(" ", LogTitle)
Skip:

  ' === 7/01/02 ===
  '  LogTitleStr = "C.M.A. STEAM PLANT SIMULATOR LOG FILE" & Chr(10)
     LogTitleStr = "STEAM PLANT SIMULATOR LOG FILE" & Chr(10)
    
    
    LogTitleStr = LogTitleStr & "DATE:          " & Format(Date, "mm:dd:yyyy") & Chr(10)
    LogTitleStr = LogTitleStr & "START TIME:    " & Format(Time, "hh:mm:ss") & Chr(10)
    LogTitleStr = LogTitleStr & "INSTRUCTOR:    " & InstructorName & Chr(10)
    LogTitleStr = LogTitleStr & "CLASS:         " & ClassName & Chr(10)
    LogTitleStr = LogTitleStr & "COURSE:        " & CourseName & Chr(10) & Chr(10) & Chr(10)
    'Call Logging(LogTitleStr, LogTitle)



Exit Sub
    
Errhandler:
    'User pressed the Cancel button
Exit Sub

End Sub

Private Sub mnu_Steam_Temp_PORT_Click()

Controller_Temp_PORT.ZOrder

End Sub

Private Sub mnu_Steam_Temp_STBD_Click()

Controller_Temp_STBD.ZOrder

End Sub

Private Sub mnu_VIAX_Click()


' 27/11/00
VIAX.ZOrder

' === 20/11/02 17.011 ===
VIAX.WindowState = 0 ' Normal

End Sub

Private Sub mnu_viewlog_Click()

ViewLogFileNumber = FreeFile
Dim MyChar As String * 1
Dim Mytext As String
Dim StrLen As Integer
Dim TmpFileName As String


CommonDialog1.CancelError = True
On Error GoTo Errhandler


CommonDialog1.DialogTitle = "View Log"

'Display name of current log  file as a suggestion
If LogFileOpen Then
    CommonDialog1.Filename = LogFileName
End If


CommonDialog1.InitDir = "c:\Steam_Sim" & "\logs\"

CommonDialog1.Filter = "Log Files (*.log)|*.log|"
' Display the Open dialog box
CommonDialog1.ShowOpen

' Open the selected file

'If selected file is the current log, read the LogString instead
If CommonDialog1.FileTitle = LogFileName Then
    frm_LogView.ZOrder
    '23/11/99 frm_LogView.RichTextBox1.text = CurrentLogString
    frm_LogView.RichTextBox1.text = LogTitleStr & CurrentLogString
    LogViewOn = True
    
    'MyText = CurrentLogString
Else
    Debug.Print "opening " & CommonDialog1.FileTitle
    
    Open CommonDialog1.FileTitle For Input Shared As #ViewLogFileNumber
    Mytext = " "
    Mytext = " "
    frm_LogView.ZOrder

    'read all text from selected log file
    Do While Not EOF(ViewLogFileNumber)
        MyChar = Input(1, #ViewLogFileNumber)   ' Get one character.
        Mytext = Mytext & MyChar
        'Debug.Print "MyText = " & MyText
    Loop
    frm_LogView.RichTextBox1.text = Mytext

    Close #ViewLogFileNumber

End If

'put text in rich text box
'Debug.Print "about to show form with rtb"
'frm_LogView.zorder
'frm_LogView.RichTextBox1.Text = MyText


Exit Sub

Errhandler:
Exit Sub


End Sub

Private Sub mnuAdvanced_Click()


' 15/2/00  Moved here

mnuAdvanced.Checked = Not mnuAdvanced.Checked

If mnuAdvanced.Checked Then
  Advanced_mode = True
  
    ' === 05/03/02 ===
    
    ' === 03/04/02 ===
    'mnu_COMBUST_CONTROL.Visible = True

Else
  Advanced_mode = False
  
    ' === 05/03/02 ===
    mnu_COMBUST_CONTROL.Visible = False
    
End If



End Sub

Private Sub mnuAIRbottles_Click()

' 25/2/98
If medium_speed Then
     sy123.ZOrder
Else
     sy123_LSS.ZOrder
End If

End Sub



Private Sub mnuAIRCOMP_Click()

If medium_speed Then
 sy121.ZOrder
Else
 sy121_LSS.ZOrder
End If

End Sub

Private Sub mnuAlarms_Click()
   Alarm_form.ZOrder
End Sub


Private Sub mnuANSI_Click()

'Global SI_units As Boolean
'Global Temp_char As String * 1      '        F or C
'Global Press_char As String * 3     '      psi or Bar
'Global Flow_chars As String         '   lbs/hr or kg/s
'
'Global Temp_mult As Single
'Global Temp_const As Single
'Global Press_mult As Single
'Global Flow_mult As Single

SI_units = False
Temp_char = "F"
Press_chars = " psi  "
Flow_chars = " lbs/hr"

Temp_mult = 1.8
Temp_const = 32#

Press_mult = 14.502     '     yyy psi = xxx Bar * 14.502 psi/Bar
Flow_mult = 7938#       '  yyy lbs/hr = xxx kg/sec * 2.205 lbs/kg  * 3600.0 secs/hour

mnuANSI.Checked = True
mnuSI.Checked = False


End Sub

Private Sub mnuBilgesys_Click()

If medium_speed Then
  sy131.ZOrder
Else
  sy131_LSS.ZOrder
End If

End Sub

Private Sub mnuBilgeTanks_Click()

If medium_speed Then
  sy130.ZOrder
Else
  sy130_LSS.ZOrder
End If

End Sub

Private Sub mnuBoiler_Port_Click()


End Sub

Private Sub mnuBoiler_Stbd_Click()


End Sub

Private Sub mnuBOILERctrl_Click()

boil.ZOrder

End Sub

Private Sub mnuBoilTest_Click()

Boilers.ZOrder

End Sub

Private Sub mnuBunkers_Click()

Bunkers.ZOrder
' === 20/11/02 17.011 ===
Bunkers.WindowState = 0 ' Normal


End Sub

Private Sub mnuCAMSHAFT_Click()

If Not medium_speed Then
    sy035.ZOrder
End If


End Sub

Private Sub mnuCondensate_Click()

' 12/7/99
Condensate.ZOrder
' === 20/11/02 17.011 ===
Condensate.WindowState = 0 ' Normal


End Sub

Private Sub mnuContinue_Click()

' 8/4/98
d_run = True    ' HALT  the     Diesel  Simulator
e_run = True    ' HALT  the Electrical  Simulator


' 13/10/09
'Me.StatusBar1.Panels.Item(2).text = "Simultor Running"
If SnapNum > 0 Then
    Me.StatusBar1.Panels.Item(2).text = "Simulator Run Snap " & Format(SnapNum, "#0")
Else
    Me.StatusBar1.Panels.Item(2).text = "Simultor Running"
End If

'23/4/98
RunIO = True
mnuHalt.Enabled = True
mnuContinue.Enabled = False


End Sub

Private Sub mnuCPP_Click()

If medium_speed Then
    sy046z.ZOrder
Else
    sy046z_LSS.ZOrder
End If

End Sub

Private Sub mnuCrank_Click()

If medium_speed Then
   sy021.ZOrder
Else
   sy021_LSS.ZOrder
End If

End Sub

Private Sub mnuCyllub_Click()

' 29/5/98
If medium_speed Then
  sy043.ZOrder
Else
  sy043_LSS.ZOrder
End If

End Sub

Private Sub mnuDA_Click()

SSDG.ZOrder
' === 20/11/02 17.011 ===
SSDG.WindowState = 0 ' Normal

End Sub

Private Sub mnuDAcompair_Click()

If medium_speed Then
  sy124.ZOrder
Else
  sy124_LSS.ZOrder
End If

End Sub

Private Sub mnuDistill_Click()

Distil.ZOrder
' === 20/11/02 17.011 ===
Distil.WindowState = 0 ' Normal

End Sub

Private Sub mnuDOBunkers_Click()

If medium_speed Then
    sy061.ZOrder
Else
    sy061_LSS.ZOrder
End If

End Sub

Private Sub mnuDOSEP_Click()

If medium_speed Then
    sy064.ZOrder
Else
    sy064_LSS.ZOrder
End If

End Sub

Private Sub mnuDtype_Click()

Dtypeblr.ZOrder

End Sub

Private Sub mnuEmergSWBD_Click()


'  16/9/99
Emergency_swbd.ZOrder

' === 20/11/02 17.011 ===
Emergency_swbd.WindowState = 0 ' Normal


'If medium_speed Then
'   sy207.zorder
'Else
'   sy207_LSS.zorder
'End If

End Sub

Private Sub mnuEnableSounds_Click()

' 7/4/98
Sounds_Enabled = Not Sounds_Enabled
If Sounds_Enabled Then
   mnuEnableSounds.Checked = True
Else
   mnuEnableSounds.Checked = False
End If

End Sub

Private Sub mnuEnglish_Click()

' 10/12/98
mnuOther.Checked = False          ' For Portuguese
mnuEnglish.Checked = True        ' For English
useEnglish = True


End Sub

Private Sub mnueot_Click()
  Frm_EOT.ZOrder
End Sub

Private Sub mnuExhaust_Click()

If medium_speed Then
  sy013.ZOrder
Else
  sy013_LSS.ZOrder
End If

End Sub

Private Sub mnuFAST_TIME_Click()

mnuFAST_TIME.Checked = Not mnuFAST_TIME.Checked

If mnuFAST_TIME.Checked Then
    status_flag = True
Else
    status_flag = False
End If

' Tell the model to go to fast or slow time...
Call set_STEAM_logicals(5, status_flag)


End Sub

Private Sub mnuFeed_Click()

Feed_Water.ZOrder
' === 20/11/02 17.011 ===
Feed_Water.WindowState = 0 ' Normal

End Sub

Private Sub mnuFire_Click()

'' 14/7/98
'If medium_speed Then
'  sy137.zorder
'Else
'  sy137_LSS.zorder
'End If

' 11/7/00
Fire_Main.ZOrder

' === 20/11/02 17.011 ===
Fire_Main.WindowState = 0 ' Normal


End Sub

Private Sub mnuFOMainEngine_Click()

'sy073.zorder
If medium_speed Then
    sy073.ZOrder
Else
    sy073_LSS.ZOrder
End If

End Sub

Private Sub mnuForced_Draft_Click()

Forced_Draft.ZOrder
' === 20/11/02 17.011 ===
Forced_Draft.WindowState = 0 ' Normal

End Sub

Private Sub mnuFOserv_Click()

FO_Service.ZOrder
' === 20/11/02 17.011 ===
FO_Service.WindowState = 0 ' Normal

End Sub

Private Sub mnuFuel_system_Click()



' 19/06/09
'Burner_FO_system.ZOrder
' === 20/11/02 17.011 ===
'Burner_FO_system.WindowState = 0 ' Normal


Burner_FO_system_NEW.ZOrder
' === 20/11/02 17.011 ===
Burner_FO_system_NEW.WindowState = 0 ' Normal



End Sub

Private Sub mnuFVCW_Click()

SY088.ZOrder

End Sub

Private Sub MNUFWG_Click()

' 14/7/98
If medium_speed Then
 sy087_med.ZOrder
Else
 sy087.ZOrder
End If


End Sub

Private Sub mnuGBOX_Click()

SY048.ZOrder

End Sub

Private Sub mnuGland_steam_Click()


' 14/2/00

'Gland_Steam.zorder

'Shaft_rpm.Caption = Format(LSS_reals(80), "##0") & " rpm"

' === 01/02/02 ===
'If LSS_reals(80) > 25# Then

  Gland_Steam_Full_away.ZOrder
  
'Else
'  Gland_Steam.ZOrder
'End If

' === 20/11/02 17.011 ===
Gland_Steam_Full_away.WindowState = 0 ' Normal


End Sub

Private Sub mnuHalt_Click()

' 8/4/98
d_run = False    ' HALT  the     Diesel  Simulator
e_run = False    ' HALT  the Electrical  Simulator

Me.StatusBar1.Panels.Item(2).text = "Simulator Halted"

'23/4/98
RunIO = False
mnuContinue.Enabled = True
mnuHalt.Enabled = False

End Sub

Private Sub mnuHOBunkers_Click()

If medium_speed Then
    sy051.ZOrder
Else
    sy051_LSS.ZOrder
End If

End Sub

Private Sub mnuHOSEP_Click()

If medium_speed Then
    sy053.ZOrder      ' HO Separator      (medium speed)
Else
    sy053_LSS.ZOrder  ' HO Separator No1  (slow speed)
End If

End Sub

Private Sub mnuHOSEP2_Click()

sy054_LSS.ZOrder

End Sub

Private Sub mnuHTCOOL_Click()

If medium_speed Then
  sy080.ZOrder
Else
  sy080_LSS.ZOrder
End If

End Sub

Private Sub mnuHTCWME_Click()

' 2/6/98
If medium_speed Then
     sy085.ZOrder
Else
   sy085_LSS.ZOrder
End If

End Sub

Private Sub mnuHWtest_Click()


'   3/9/99
'  OLD PORTUGAL Harware Test...
'IOtest.zorder
'IOtest.Visible = True

'  NEW APM ( Wiedmuller Win Block )  Hardware Test
' Get a form un-load event but why?....
mnuHalt_Click
FrmHardwareTest.ZOrder

End Sub

Private Sub mnuINaGROUP_Click()

' 26/3/98
In_a_group = Not In_a_group
If In_a_group Then
   mnuINaGROUP.Checked = True
Else
   mnuINaGROUP.Checked = False
End If

End Sub

Private Sub mnuLOSEP_Click()

If medium_speed Then
    sy041.ZOrder
Else
    sy041_LSS.ZOrder
End If

End Sub

Private Sub mnuLTCW_Click()

' 3/6/98
If medium_speed Then
 sy090.ZOrder
Else
 sy090_LSS.ZOrder
End If

End Sub

Private Sub mnuLub_Click()

' 22/5/98
If medium_speed Then
    sy030.ZOrder
Else
    sy030_LSS.ZOrder
End If

End Sub

Private Sub mnuLube_oil_Click()

Lube_Oil.ZOrder
' === 20/11/02 17.011 ===
Lube_Oil.WindowState = 0 ' Normal

End Sub

Private Sub mnuMain_Bearings_Click()

Main_Bearings.ZOrder
' === 20/11/02 17.011 ===
Main_Bearings.WindowState = 0 ' Normal

End Sub

Private Sub mnuMain_Condenser_Click()

Main_Condenser.ZOrder
' === 20/11/02 17.011 ===
Main_Condenser.WindowState = 0 ' Normal

End Sub

Private Sub mnuMain_Turbines_Click()

Main_Turbines.ZOrder
' === 20/11/02 17.011 ===
Main_Turbines.WindowState = 0 ' Normal

End Sub

Private Sub mnuME_Click()

If medium_speed Then
  sy011.ZOrder
Else
  sy011_LSS.ZOrder
End If

End Sub

Private Sub mnuMEcontrol_Click()


' === 13/11/02 ===
' DOES NOT EXIST on Steam_Sim...
'If medium_speed Then
'   MEcontrol.ZOrder
'Else
'   MEcontrol_LSS.ZOrder
'End If

End Sub

Private Sub mnuMedium_Click()

' 1/6/98 NOW IN INTRODUCTION.FRM
'   medium_speed = True
'   mnuSLOWSPEED.Checked = False
'   mnuMEDIUM.Checked = True

End Sub



Private Sub mnuModel_Constants_Click()

' === 23/08/02 ===
Model_Constants.ZOrder

End Sub

Private Sub mnuOMD_Click()

If medium_speed Then
 sy023.ZOrder
Else
 sy023_LSS.ZOrder
End If

End Sub

Private Sub mnuOther_Click()


' 10/12/98
mnuOther.Checked = True           ' For Portuguese
mnuEnglish.Checked = False        ' For English
useEnglish = False


End Sub

Private Sub mnuPCINCONTROL_Click()

Call change_PC_control

End Sub

Private Sub mnuReset_Click()

' 10/4/98
reset_flag = True
Unload Me

End Sub



Private Sub mnuRunGear_Click()

If medium_speed Then
   sy020.ZOrder
Else
   sy020_LSS.ZOrder
End If

End Sub

Private Sub mnuSalt_Water_Click()

Salt_Water.ZOrder
' === 20/11/02 17.011 ===
Salt_Water.WindowState = 0 ' Normal

End Sub

Private Sub mnuScavAir_Click()

If medium_speed Then
 sy012a.ZOrder
Else
 sy012_LSS.ZOrder
End If


End Sub

Private Sub mnuSI_Click()

' 14/7/99
SI_units = True
Temp_char = "C"
Press_chars = " Bar"
Flow_chars = " kg/s"

Temp_mult = 1#
Temp_const = 0#

Press_mult = 1#
Flow_mult = 1#

mnuANSI.Checked = False
mnuSI.Checked = True


End Sub

Private Sub mnuSlowspeed_Click()

' 1/6/98 NOW IN INTRODUCTION.FRM
'   medium_speed = False
'   mnuSLOWSPEED.Checked = True
'   mnuMEDIUM.Checked = False

End Sub


Private Sub mnuSoundMix_Click()

' 7/4/98
SoundMix.ZOrder

' 24/4/98
'''''frmHiddenSounds.Visible = True

End Sub

Private Sub mnuSteam_Calc_Click()

Steam_Calculator.ZOrder

End Sub

Private Sub mnuSWBD2_Click()


' 5/10/99

'If medium_speed Then
'   sy205.zorder
'Else
'   sy205_LSS.zorder
'End If



End Sub

Private Sub mnuSWBD4_Click()

If medium_speed Then
   sy206.ZOrder
Else
   sy206_LSS.ZOrder
End If

End Sub

Private Sub mnuSteamSteam_Click()


' 2/8/00
Stm_Stm_Generator.ZOrder

' === 20/11/02 17.011 ===
Stm_Stm_Generator.WindowState = 0 ' Normal

End Sub

Private Sub mnuSWBD11_Click()



main_SB_1.ZOrder
' === 20/11/02 17.011 ===
main_SB_1.WindowState = 0 ' Normal

End Sub

Private Sub mnuSWBD13_Click()


main_SB_2.ZOrder
' === 20/11/02 17.011 ===
main_SB_2.WindowState = 0 ' Normal

End Sub

Private Sub mnuSY010_Click()

' 4/6/98
If medium_speed Then
  sy010.ZOrder
Else
  sy010_LSS.ZOrder
End If

End Sub

Private Sub mnusy071_Click()

If medium_speed Then
    sy071.ZOrder
Else
    sy071_LSS.ZOrder
End If


End Sub

Private Sub mnusy072_Click()

If medium_speed Then
    sy072.ZOrder
Else
    sy072_LSS.ZOrder
End If

End Sub

Private Sub mnuSY091_Click()

' 3/6/98
If medium_speed Then
   sy091.ZOrder
Else
   sy091_LSS.ZOrder
End If

End Sub

Private Sub mnuSY096_Click()

' 3/6/98
If medium_speed Then
  sy096.ZOrder
Else
  sy096_LSS.ZOrder
End If
End Sub

Private Sub mnusy101_Click()

If medium_speed Then
    sy101.ZOrder
Else
    sy101_LSS.ZOrder
End If

End Sub

Private Sub mnusy104_Click()
 
 
If medium_speed Then
  sy104.ZOrder
Else
  sy104_LSS.ZOrder
End If

End Sub

Private Sub mnuSY150_Click()

sy150_LSS.ZOrder

End Sub

Private Sub mnuSY151_Click()

sy151_LSS.ZOrder


End Sub

Private Sub mnuSY155_Click()

sy155_LSS.ZOrder

End Sub

Private Sub mnuSY156_Click()

sy156_LSS.ZOrder


End Sub

Private Sub mnuSY161_Click()

sy161_LSS.ZOrder

End Sub

Private Sub mnuSY164_Click()

sy164_LSS.ZOrder


End Sub

Private Sub mnuSY171_Click()

sy171_LSS.ZOrder


End Sub

Private Sub mnuSY172_Click()

sy172_LSS.ZOrder


End Sub

Private Sub mnuSY200_Click()


Main_SB.ZOrder

' === 20/11/02 17.011 ===
Main_SB.WindowState = 0 ' Normal

End Sub

Private Sub Host_sock_Tx_Accept(Index As Integer, SocketId As Integer)

Debug.Print "Host_sock_Tx_Accept Event, Index = " & Index

End Sub

Private Sub Host_sock_Tx_Blocking(Index As Integer, Status As Integer, Cancel As Integer)

Debug.Print "Host_sock_Tx_Blocking Event, Index = " & Index

End Sub


Private Sub Host_sock_Tx_Connect(Index As Integer)

Debug.Print "Host_sock_Tx_Connect Event, Index = " & Index

HostSockTxConnected(Index) = True

End Sub



Private Sub Host_sock_Tx_Disconnect(Index As Integer)

Debug.Print "Entered Host_sock_Tx Disconnect,  Index = " & Index
Host_sock_Tx(Index).Disconnect
HostSockTxConnected(Index) = False

End Sub

Private Sub Host_sock_Tx_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)

Debug.Print "Host_sock_Tx(" & Index & ") - LastError : " & ErrorString & " Error Code: " & ErrorCode

' 16/3/98
Host_sock_Tx(Index).Flush

' 20/3/98 REMOVED...
'Host_sock_Tx(Index).Disconnect


End Sub

Private Sub mnu_tcpip_array_Click()
    
''' 20/3/98    frm_tcpip.zorder

End Sub


Private Sub mnuTest_Scope_Click()

'20/10/99
TEST_SCOPE.ZOrder

End Sub

Private Sub mnuTURBO_Click()


SSTG.ZOrder

' === 20/11/02 17.011 ===
SSTG.WindowState = 0 ' Normal

End Sub

Private Sub mnuUseEngineRoom_Click()


'1/12/99 make this flag permanent. make option invisible & set true in form load

mnuUseEngineRoom.Checked = Not mnuUseEngineRoom.Checked
'22/11/99
'status_flag = True
If mnuUseEngineRoom.Checked Then
  status_flag = True
Else
  status_flag = False
End If


' Tell the model to use the Engine Room...
Call set_STEAM_logicals(71, status_flag)


End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

' 23/7/98
flash_latest_alarm = False


' 5/4/00  TO TEST Engineer Call-out
' WORKED OK but we do NOT need this here ( actuate from DI(261) ONLY )
'Call setPCDI(261)




'    === 28/03/01 15.034 ===

'    LSS_LOGICALS(97) = Reserved ( See in VB it is used as Silence_BLEEP )

'If LSS_LOGICALS(97) Then
'  status_flag = False
'Else
'  status_flag = True
'End If

If I_AM_HOST Then
  status_flag = True
  Call set_LSS_logicals(97, status_flag)
End If



End Sub

Private Sub Super_Pro_Dongle_Timer_Timer()

' === 11/09/03 ===
If Super_Pro_Dongle Then

    Call Dongle_check
    
End If


End Sub

Private Sub Timer_IO_Timer()
Dim All_cards_read_OK As Boolean



' === 21/12/01 ===
If Exit_sharply Then
    Unload Me
End If



'21/4/98

If RunIO Then
    If I_AM_HOST Then
    
''''''' 27/4/98       Call io   ' ( DONE IN FORTRAN )

      '2/9/99  Allister's IO ( WiedMuller Win blocks )
      
      Call get_steam_DO(D_OUTPUT(0))
      Call get_steam_AO(A_OUTPUT(0))
  
      
      
    ' 22/10/99
    ' 29/10/99  use new flag hardware_flag  If PTT Then
        If Not HARDWARE_flag Then
        ' 22/10/99      We need this flag false ( for eg Main_SB.FRM AVR switch and pot to work )
          All_cards_read_OK = False
  
        Else
  
        ' 22/10/99      Read BOTH APPLICOM IO Cards...
            All_cards_read_OK = ReadAllCards(Timer_IO)
        
        End If
      
      
        If All_cards_read_OK Then
          ' 5/9/99
          Call set_steam_DI(D_INPUT(0))
      
         ' 27/9/99  Debug.Print "D_INPUT(751)" & D_INPUT(751)
      
          Call set_steam_AI(A_INPUT(0))
        
        Else
        
            ' === 7/01/02 ===
            '   Me.StatusBar1.Panels.Item(6).text = "I/O problem " & CardStatus
            If HARDWARE_flag Then
                Me.StatusBar1.Panels.Item(6).text = "I/O problem " & CardStatus
            Else
                Me.StatusBar1.Panels.Item(6).text = " "
           End If
              
        End If
      '24/10/99 add i/o logging
      Call LogDIAIs
        
    End If
End If


End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim this_alarm_text As String





' 16/8/99
flash_flag = Not flash_flag

' 8/12/98
'If MDIform_loaded = False Then
'    Exit Sub
'End If



'test only
'If LogFileOpen Then
   ' LogString = LogString & Chr(10) & Chr(13) & Time
'End If

'---------------------- experiment 29/4/98 ----------------------------------
' 29/4/98
If d_run Or e_run Then

    runIO_counter = runIO_counter + 1
    If runIO_counter > 10 Then
        runIO_counter = 10
    End If
'
 If runIO_counter > 4 Then
 
 
            ' === 15/03/01 ============

            If Not PTT And I_AM_HOST Then

                ' I must be HASP00 and running FULL MISSION

                If Dongle_Enable_Flag Then
    
                    ' Do nothing here
  
                Else
    
                        ' Send a message "DISABLE_dongle_requirement"
                        ' ( especially intended for HASP01 , which is a second IOS )

                        ' equivalence(STEAM_logicals(200),DISABLE_dongle_requirement)
                        ' logical*1           DISABLE_dongle_requirement
                        status_flag = True
                        Call set_STEAM_logicals(200, status_flag)
          
                End If
    
            End If

            'If This_PC_Name = "PC02" Then
             If This_PC_Name = "HASP01" Then

                If PTT Then
    
                    ' We are running PTT and we NEED a dongle
                    Dongle_Enable_Flag = True
   
                Else
    
    
                    ' === 26/09/01 REMOVED THE FOLLOWING CODE ===
                    '' We are running FULL MISSION
                    '' so look at the DISABLE_dongle_requirement flag ( sent by the HOST HASP00 )
                    'If VB_steam_logicals(200) Then
                    '  Dongle_Enable_Flag = False
                    'Else
                    '  Dongle_Enable_Flag = True
                    'End If
        
                End If
    
             End If

            ' === 15/03/01 end of edit ===

 
        ' 11/12/98
        '    Me.MousePointer = ccDefault
        If cold_start_hourglass Then
           cold_start_hourglass = False
            Me.MousePointer = ccDefault
        End If
        

''         'IOtest.zorder
 '        RunIO = True
 '        Call setPCDI(2567)
 '        IOtest.Visible = False
 '
 '        RunIO = True

 '   End If
'---------------------- end  experiment 29/4/98 ----------------------------------
    
    
'If d_run Or e_run Then

' 17/6/98
'  If (set_DI_flag And Not In_a_group) Then
'     Call setDI(nDI, True)
'     set_DI_flag = False
'  End If
  
  If (set_DI_flag) Then
  
     DI_counter = DI_counter + 1
     
     If DI_counter = 1 Then
       Call setDI(nDI, True)    ' Reset Slowdown alarm
     End If
     If DI_counter = 4 Then
       Call setDI(nDI, False)
       DI_counter = 0
       set_DI_flag = False
     End If
       
  End If
  
  
  
'  18/6/98
'  If I_AM_HOST Then
   If I_AM_HOST Or Not In_a_group Then
   
     

      ' 5/4/00   Engineer Call-out
        If D_OUTPUT(937) Then
            If flash_flag Then
              Me.StatusBar1.Panels.Item(7).text = "Engineer Call"
              Me.StatusBar1.Panels.Item(7).Visible = True
            Else
              Me.StatusBar1.Panels.Item(7).text = ""
              Me.StatusBar1.Panels.Item(7).Visible = False
            End If
            
        Else
              Me.StatusBar1.Panels.Item(7).text = ""
              Me.StatusBar1.Panels.Item(7).Visible = False
        End If

        '   2/9/99
        If VB_steam_logicals(5) Then
                Me.StatusBar1.Panels.Item(5).text = "FAST TIME"
        Else
        
            '   10/8/00
            '   Me.StatusBar1.Panels.Item(5).text = "NORMAL TIME"
           
            If VB_steam_logicals(121) Then
                Me.StatusBar1.Panels.Item(5).text = "SLOW TIME"
            Else
                Me.StatusBar1.Panels.Item(5).text = "NORMAL TIME"
            End If
            
        End If
        
      
    ' === 23/01/02 ===
    '  If medium_speed Then
    '    Call D_MODEL(split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    '  Else
           
        ' === 06/03/03 ===
        ' Call D_MODEL_LSS(CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
         Call D_MODEL_LSS(Ballast_flag, CMA_flag, PTT_mode, split_mode, ss_supply, diesel_init, elec_init, diesel_cold, elec_cold, d_run, e_run, snap_shot, hw_test)
    
     ' End If
    
      
     
     ' 29/4/98 MOVE to COLD and SNAPSHOT start...
     ' 28/4/98  RUN the IO task....
     '     If (diesel_init Or elec_init) Then
     '         nPCDI = 2567
     '         Call setPCDI(nPCDI)
     '      End If
           
           

     RtoS = True
     ItoS = True
   Else
     RtoS = False
     ItoS = False
   End If
  
  
'       9/8/00
'          diesel_init = False
'            elec_init = False
'          diesel_cold = False
'            elec_cold = False


' === 23/01/02 ===
' If medium_speed Then
'  Call get_reals(RtoS, PCCONTROL(1), PIDDATA(1), VV_position(1), PCA(1), PML(1), ZML(1), TML(1), FML(1), _
'  FLS(1), ZLS(1), TLS(1), ZFT(1), FMY(1), ZMY(1), PMY(1), PHT(1), THT(1), fHt(1), ZHT(1), _
'  PLT(1), TLT(1), FLT(1), ZLT(1), _
'  VB_ENGINE(1), _
'  PSW(1), TSW(1), FSW(1), _
'  PMF(1), TMF(1), FMF(1), _
'  PFT(1), TFT(1), FFT(1), _
'  ZBS(1), _
'  ALARM_HILO(1, 1), _
'  elec_reals(1), _
'  plot_Reals(1))
' Else

  Call get_reals_LSS(RtoS, LSS_reals(0), ALARM_HILO(1, 1), plot_Reals(1))
  
 'End If
 
 
' 9/8/00
If diesel_cold Or elec_cold Then
            Force_Instrp_update = True
End If
          diesel_init = False
            elec_init = False
          diesel_cold = False
            elec_cold = False

  
 If Not GotAlarmSetPoints Then
    For i = 1 To 450
        Default_ALARM_HILO(1, i) = ALARM_HILO(1, i)
        Default_ALARM_HILO(2, i) = ALARM_HILO(2, i)
        GotAlarmSetPoints = True
    Next
 End If
 
  
 If medium_speed Then
   Call get_integers(ItoS, _
    v_status(1), vv_status(1), p_status(1), bkr_status(1), rma_status(1), _
    VB_pots(0), _
    VB_lamps(1), elec_lamps(1), _
    alarm(1), elec_alarm(1), _
    active_actions(1), active_reset_actions(1), _
    active_faults(1), _
    I_sounds(1))
  Else
  
' 5/6/98 added active_actions  etc etc
' 14/7/99 ADDED VB_steam_logicals

' === 7/01/02 added VB_More_integers(1) ===
     Call get_integers_LSS(ItoS, PC_INT_STRING, alarm(1), elec_alarm(1), LSS_LOGICALS(1), _
     active_actions(1), active_reset_actions(1), _
     active_faults(1), _
     I_sounds(1), _
     VB_AI(1), _
     VB_VV_extras(1), VB_steam_logicals(1), VB_steam_integers(1), VB_More_integers(1))
     
  End If
  
  
          ' === 17/8/00 ======
        If I_AM_HOST Then
            i = 127
            If PTT Then
                status_flag = True
                ' REMEMBER ONLY SIMPLE IF Statements work on VB_steam_logicals so...
                If VB_steam_logicals(127) Then
                Else
                    Call set_STEAM_logicals(i, status_flag)
                End If
            Else
                status_flag = False
                ' REMEMBER ONLY SIMPLE IF Statements work on VB_steam_logicals so...
                If VB_steam_logicals(127) Then
                    Call set_STEAM_logicals(i, status_flag)
                End If
            End If
                       
        End If
        ' === 17/8/00 end ===

  
        
        
' === 13/11/02 OLD METHOD... ======================

'    If HARDWARE_flag Then
'        update_VBAIs = True
'    End If
'    If update_VBAIs Then
'            update_VBAIs = False
'            For i = 1 To 48
'                update_VBAI(i) = True
'            Next i
'    End If
  
' === 13/11/02 NEW METHOD ======================
  For i = 1 To 48
      If (VB_AI_old(i) <> VB_AI(i)) Then
          VB_AI_old(i) = VB_AI(i)
          update_VBAI(i) = True
     End If
  Next i
' === 13/11/02 NEW METHOD end of edit ==========

' === 18/11/02 ==============
'equivalence(STEAM_integers(46),Vscroll_value(1))
    For i = 1 To 50
        If (VB_steam_integers_OLD(45 + i) <> VB_steam_integers(45 + i)) Then
            VB_steam_integers_OLD(45 + i) = VB_steam_integers(45 + i)
            update_VSCROLL(i) = True
        End If
    Next i
' === 18/11/02 END of edit ===

        
        
 
' 23/3/98 ( NOW IN get_integers )
'    Call get_v_status(v_status(1))          ' Starter Box Valve Status
'    Call get_vv_status(vv_status(1))        ' Filter and Electrical Ancilliary Valve Status
'    Call get_p_status(p_status(1))          ' Pump Status
'    Call get_bkr_status(bkr_status(1))      ' BREAKER Status
'    Call get_rma_status(rma_status(1))      ' RMA Status
'    Call get_pots(VB_pots(0))
'
'    Call get_lamps(VB_lamps(1))
'    Call get_elec_lamps(elec_lamps(1))
'
'    Call get_alarms(alarm(1))               ' Status of ALL the DIESEL Alarms
'    Call get_elec_alarms(elec_alarm(1))     ' Status of ALL the ELECTRICAL Alarms
  
  
    current_time = Time        ' Current time i.e. NOW
    
    Call catch_alarms          ' Date the ALARMS and put in Chronological order for the ALARM_FORM
  
' 19/3/98 NOW in get_reals
'  Call get_HILO(ALARM_HILO(1, 1))
  
' 29/12/97 Display the LATEST Alarm
    If alarm_latest(1) > 0 Then
    
                 ' 28/5/98
                 '29/11/99
                'If medium_speed Then
                    'this_alarm_text = Mid$(alarmtxt(alarm_latest(1)), 6, 31)
                'Else
'                   this_alarm_text = Mid$(alarmtxt(alarm_no), 5, 32)
                    'this_alarm_text = Mid$(alarmtxt(alarm_latest(1)), 5, 32)
                'End If
                    
        '29/11/99 corrected for CMA
       ' === 6/12/00 Further Correction for CMA....
       'this_alarm_text = Mid$(alarmtxt(alarm_latest(1)), 9, Len(alarmtxt(alarm_latest(1))))
        this_alarm_text = Mid$(alarmtxt(alarm_latest(1)), 10, Len(alarmtxt(alarm_latest(1))))
           
        ' ===== 23/7/98 ==============
        ' StatusBar1.Panels(3).Text = alarm_time(1) & "   " & Format$(alarm_latest(1), "000") & "    " & this_alarm_text
           
        ' see catch_alarms    which sets: flash_latest_alarm = True
        ' see statusbar click which sets: flash_latest_alarm = false
        
        ' === 23/01/02 ===============
        If LSS_LOGICALS(97) Then    ' Sent by the HOST
          flash_latest_alarm = False
        End If
        ' === 23/01/02 end of edit ===
        
        
         If flash_latest_alarm Then
           If StatusBar1.Panels(3).text = "" Then
            StatusBar1.Panels(3).text = alarm_time(1) & "  #" & Format$(alarm_latest(1), "000") & "    " & this_alarm_text
            Beep
           Else
            StatusBar1.Panels(3).text = ""
           End If
         Else
            StatusBar1.Panels(3).text = alarm_time(1) & "  #" & Format$(alarm_latest(1), "000") & "    " & this_alarm_text
         End If
         
        ' ===== 23/7/98 end ===========
        
        
'8/6/98
          If OldAlarmTime <> alarm_time(1) Then
            'Call Logging(Format$(alarm_time(1) & alarm_latest(1), "000") & "    " & this_alarm_text, LogAlarms)
            Call Logging(" ALARM   " & this_alarm_text, LogAlarms)
            OldAlarmTime = alarm_time(1)
          End If
          
     
     Else
          StatusBar1.Panels(3).text = ""
     End If
   
'28/4/98
 End If



  If I_AM_HOST Then

    ' Call FILL_PCREALS
  
     
     'Send PC_REALS() to the first PC....(Host_sock_Tx(0))
 '25/3/98   If frm_Host.Host_sock_Tx(0).Connected Then
 '25/3/98    If frm_Host.Host_sock_Tx(0).Connect Then
 
 
 ' 17/6/98
 '           If HostSockTxConnected(Index) = True Then
 '              Call tcpip_send_array_single(0, PC_REALS())
 '           End If
 '
 '           If HostSockTxConnected(Index) = True Then
 '               Call tcpip_send_array_integer(1, PC_INTS())
 '           End If
 
 
        ' === 07/03/07 ===
        'For i = 0 To 4
        For i = 0 To 14
        
           If HostSockTxConnected(i * 2) = True Then
               Call tcpip_send_array_single(i * 2, PC_REALS())
            End If
     
            If HostSockTxConnected((i * 2) + 1) = True Then
                Call tcpip_send_array_integer((i * 2) + 1, PC_INTS())
            End If
        Next i
            
            
      
  End If
  
End If

' 15/10/98
If DEMO_flag Then
    If DEMO_halt_model Then
      d_run = False
      e_run = False
    End If
End If


End Sub


Private Sub Timer2_Timer()

'== 7/4/98 ==========
Dim i As Integer
'For Full mission, only HASP00 knows if model is in halt so Tx
'this flag to others - especially HASP01 to control the Feed Pumps

'13/12/00
If This_PC_number = 0 Then
    If d_run Then
        status_flag = True  'R.Beams 13/12/00
    Else
        status_flag = False
    End If
    If d_run <> old_d_run Then
        old_d_run = d_run
        Call set_STEAM_logicals(195, status_flag)
    End If
End If
    
    
'13/12/00 If d_run Then
'HASP0 to call sounds if model running. HASP 01 same, but flag received from HASP 00
If (This_PC_number = 0 And d_run) Or (This_PC_number = 1 And VB_steam_logicals(195)) Then
        frmHiddenSounds.SoundWhenModelRunning
Else
        If Not Sound_test_on Then   ' STOP Sounds
        '13/12/00
            If This_PC_number = 0 Then
                    '13/12/00 For i = 1 To 6
                For i = 1 To 16
                    frmHiddenSounds.snd.DSChannel = i
                    frmHiddenSounds.snd.DSCommand = 4 ' STOP Sounds
                Next i
            ElseIf This_PC_number = 1 Then
                For i = 1 To 2
                    frmHiddenSounds.snd.DSChannel = i
                    frmHiddenSounds.snd.DSCommand = 4 ' STOP Sounds
                Next i
            End If
        End If
End If
'============

End Sub
