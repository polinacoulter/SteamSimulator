Attribute VB_Name = "SocketTools"
'
' Catalyst SocketTools 1.0
' Copyright (c) 1995-1996, Catalyst Development Corp. All rights reserved.
'
' This file contains the constants and function declarations used
' with the SocketTools TCP/IP development toolkit for Visual Basic 3.0
'

'
' General constants used with most of the controls
'
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1

'
' Windows Sockets error codes. These error codes are returned by
' the DNS (Domain Name Service), Ping and SocketWrench controls.
'
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 25004

'
' Constants used with the Domain Name Service (DNS) cotrol
'
' DNS client actions
'
Public Const DNS_ACTION_NONE = 0
Public Const DNS_ACTION_CANCEL = 1

'
' Constants used with the File Transfer Protocol (FTP) control.
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event.
'
Public Const FTP_ERROR = -1
Public Const ERROR_FTP_NOHANDLES = 22001    ' no free handles
Public Const ERROR_FTP_ALLOCMEM = 22002     ' no Public memory
Public Const ERROR_FTP_NOTOWNER = 22003     ' not session owner
Public Const ERROR_FTP_NETSTART = 22004     ' cannot initialize network
Public Const ERROR_FTP_HOSTNAME = 22005     ' invalid hostname
Public Const ERROR_FTP_SOCKET = 22006       ' cannot create socket
Public Const ERROR_FTP_ACCEPT = 22007       ' cannot accept connection
Public Const ERROR_FTP_CONNECT = 22008      ' cannot connect to server
Public Const ERROR_FTP_PEERNAME = 22009     ' cannot get peer information
Public Const ERROR_FTP_BIND = 22010         ' cannot bind socket
Public Const ERROR_FTP_LISTEN = 22011       ' cannot listen for connection
Public Const ERROR_FTP_PARAM = 22012        ' invalid parameter
Public Const ERROR_FTP_HANDLE = 22013       ' invalid FTP handle
Public Const ERROR_FTP_SELECT = 22014       ' socket select error
Public Const ERROR_FTP_READ = 22015         ' cannot read from socket
Public Const ERROR_FTP_WRITE = 22016        ' cannot write to socket
Public Const ERROR_FTP_EOF = 22017          ' end of file on data connection
Public Const ERROR_FTP_BUSY = 22018         ' connection in use
Public Const ERROR_FTP_NOTBUSY = 22019      ' connection is not in use
Public Const ERROR_FTP_TIMEOUT = 22020      ' timeout waiting for socket
Public Const ERROR_FTP_SYNC = 22021         ' internal synchronization error
Public Const ERROR_FTP_CLOSED = 22022       ' server closed connection
Public Const ERROR_FTP_NOTREADY = 22023     ' server is not ready, try again
Public Const ERROR_FTP_NOTAVAIL = 22024     ' remote server is not available
Public Const ERROR_FTP_SERVICE = 22025      ' unknown problem with remote server
Public Const ERROR_FTP_DATAPORT = 22026     ' unable to set data port
Public Const ERROR_FTP_DATAOPEN = 22027     ' data port already open
Public Const ERROR_FTP_DATACLOSED = 22028   ' data port is closed
Public Const ERROR_FTP_ABORTED = 22029      ' file transfer aborted
Public Const ERROR_FTP_READONLY = 22030     ' data connection is read-only
Public Const ERROR_FTP_WRITEONLY = 22031    ' data connection is write-only
Public Const ERROR_FTP_FILETYPE = 22032     ' unable to set file type
Public Const ERROR_FTP_PASSIVE = 22033      ' data connection is passive
Public Const ERROR_FTP_ADDRESS = 22034      ' cannot determine passive address
Public Const ERROR_FTP_CREATEFILE = 22035   ' cannot create local file
Public Const ERROR_FTP_OPENFILE = 22036     ' cannot open local file
Public Const ERROR_FTP_READFILE = 22037     ' cannot read local file
Public Const ERROR_FTP_WRITEFILE = 22038    ' cannot write local file
Public Const ERROR_FTP_LOGIN = 22039        ' cannot login to remote server

'
' Result codes returned by the FTP server. The ResultCode property
' specifies the code from the last operation.
'
Public Const FTP_REPLY_RESTMARK = 110   ' restart marker reply
Public Const FTP_REPLY_NOTREADY = 120   ' service available in n minutes
Public Const FTP_REPLY_DATAOPEN = 125   ' data connection open, transfer started
Public Const FTP_REPLY_FILEOK = 150     ' file status okay
Public Const FTP_REPLY_CMDOK = 200      ' command okay
Public Const FTP_REPLY_CMDIGN = 202     ' command ignored
Public Const FTP_REPLY_SYSSTAT = 211    ' system status
Public Const FTP_REPLY_DIRSTAT = 212    ' directory status
Public Const FTP_REPLY_FILESTAT = 213   ' file status
Public Const FTP_REPLY_HELPMSG = 214    ' human-readable help response
Public Const FTP_REPLY_READY = 220      ' service ready for new user
Public Const FTP_REPLY_CLOSED = 221     ' service closing connection
Public Const FTP_REPLY_DATAOPENED = 225 ' data connection open
Public Const FTP_REPLY_DATACLOSED = 226 ' closing data connection
Public Const FTP_REPLY_PASVMODE = 227   ' entering passive mode
Public Const FTP_REPLY_LOGIN = 230      ' user logged in
Public Const FTP_REPLY_DONE = 250       ' requested file action completed
Public Const FTP_REPLY_GETPASS = 331    ' username okay, need password
Public Const FTP_REPLY_GETACCT = 332    ' need account for login
Public Const FTP_REPLY_PENDING = 350    ' file action pending
Public Const FTP_REPLY_OPENFAIL = 425   ' cannot open data connection
Public Const FTP_REPLY_ABORTED = 426    ' conenct closed, transfer aborted
Public Const FTP_REPLY_NOTAVAIL = 450   ' file is not available
Public Const FTP_REPLY_LOCALERR = 451   ' local error
Public Const FTP_REPLY_NOSPACE = 452    ' no space on server system
Public Const FTP_REPLY_BADSYN = 500     ' syntax error
Public Const FTP_REPLY_BADARG = 501     ' invalid command arguments
Public Const FTP_REPLY_BADCMD = 502     ' command not implemented
Public Const FTP_REPLY_BADSEQ = 503     ' bad sequence of commands
Public Const FTP_REPLY_BADPARM = 504    ' bad command parameter
Public Const FTP_REPLY_NOLOGIN = 530    ' user not logged in
Public Const FTP_REPLY_ACCTREQ = 532    ' account required for storing files
Public Const FTP_REPLY_NOFILE = 550     ' file unavailable
Public Const FTP_REPLY_BADPAGE = 551    ' page type unknown
Public Const FTP_REPLY_EXQUOTA = 552    ' exceeded file storage quota
Public Const FTP_REPLY_BADFILE = 553    ' invalid file name

Public Const FTP_FILE_READ = 0          ' read from file on remote host
Public Const FTP_FILE_WRITE = 1         ' write to file on remote host
Public Const FTP_FILE_APPEND = 2        ' append to file on remote host
Public Const FTP_FILE_UNIQUE = 3        ' create unique file on remote host

'
' FTP client actions
'
Public Const FTP_ACTION_NONE = 0
Public Const FTP_ACTION_CONNECT = 1
Public Const FTP_ACTION_LOGIN = 2
Public Const FTP_ACTION_LOGOUT = 3
Public Const FTP_ACTION_DISCONNECT = 4
Public Const FTP_ACTION_GETFILE = 5
Public Const FTP_ACTION_PUTFILE = 6
Public Const FTP_ACTION_FILELIST = 7
Public Const FTP_ACTION_OPEN = 8
Public Const FTP_ACTION_CREATE = 9
Public Const FTP_ACTION_APPEND = 10
Public Const FTP_ACTION_CLOSE = 11
Public Const FTP_ACTION_CANCEL = 12
Public Const FTP_ACTION_DELETE = 13
Public Const FTP_ACTION_RENAME = 14
Public Const FTP_ACTION_GETCWD = 15
Public Const FTP_ACTION_CHDIR = 16
Public Const FTP_ACTION_MKDIR = 17
Public Const FTP_ACTION_RMDIR = 18

'
' FTP client states
'
Public Const FTP_CLIENT_UNUSED = 0
Public Const FTP_CLIENT_IDLE = 1
Public Const FTP_CLIENT_CONNECTING = 2
Public Const FTP_CLIENT_LOGIN = 3
Public Const FTP_CLIENT_TRANSFER = 4

'
' Image viewer control constants
'
Public Const INVALID_IMAGE = 0

'
' Image types supported by the control
'
Public Const IMAGE_TYPE_NONE = 0
Public Const IMAGE_TYPE_BITMAP = 1
Public Const IMAGE_TYPE_GIF = 2
Public Const IMAGE_TYPE_PCX = 3
Public Const IMAGE_TYPE_JPEG = 4
Public Const IMAGE_TYPE_XBM = 5
Public Const IMAGE_TYPE_EXTERN = 6

'
' Network News Transfer Protocol (NNTP) control constants
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event.
'
Public Const NNTP_ERROR = -1
Public Const ERROR_NNTP_HANDLE = 22301
Public Const ERROR_NNTP_PARAM = 22302
Public Const ERROR_NNTP_INIT = 22303
Public Const ERROR_NNTP_ALLOC = 22304
Public Const ERROR_NNTP_HOSTNAME = 22305
Public Const ERROR_NNTP_SOCKET = 22306
Public Const ERROR_NNTP_CONNECT = 22307
Public Const ERROR_NNTP_SELECT = 22308
Public Const ERROR_NNTP_TIMEOUT = 22309
Public Const ERROR_NNTP_SERVER = 22310
Public Const ERROR_NNTP_READ = 22311
Public Const ERROR_NNTP_WRITE = 22312
Public Const ERROR_NNTP_CLOSED = 22313
Public Const ERROR_NNTP_BUSY = 22314
Public Const ERROR_NNTP_ABORTED = 22315
Public Const ERROR_NNTP_NOSERVICE = 22316
Public Const ERROR_NNTP_BADGROUP = 22317
Public Const ERROR_NNTP_NOGROUP = 22318
Public Const ERROR_NNTP_EMPTYGROUP = 22319
Public Const ERROR_NNTP_NOARTICLE = 22320
Public Const ERROR_NNTP_LASTARTICLE = 22321
Public Const ERROR_NNTP_FIRSTARTICLE = 22322
Public Const ERROR_NNTP_BADARTICLE = 22323
Public Const ERROR_NNTP_NOTRANSFER = 22324
Public Const ERROR_NNTP_BADTRANSFER = 22325
Public Const ERROR_NNTP_REJECTED = 22326
Public Const ERROR_NNTP_NOPOSTING = 22327
Public Const ERROR_NNTP_BADPOST = 22328
Public Const ERROR_NNTP_COMMAND = 22329
Public Const ERROR_NNTP_RESTRICTED = 22330
Public Const ERROR_NNTP_FAULT = 22331
Public Const ERROR_NNTP_AUTHINFO = 22332
Public Const ERROR_NNTP_AUTHFAIL = 22333
Public Const ERROR_NNTP_AUTHTYPE = 22334
Public Const ERROR_NNTP_NOAUTH = 22335

'
' Result codes returned by the NNTP server. The ResultCode property
' specifies the code from the last operation.
'
Public Const NNTP_REPLY_READY = 200
Public Const NNTP_REPLY_SELECTED = 211
Public Const NNTP_REPLY_GROUPLIST = 215
Public Const NNTP_REPLY_ARTICLE = 220
Public Const NNTP_REPLY_HEADER = 221
Public Const NNTP_REPLY_BODY = 222
Public Const NNTP_REPLY_STATUS = 223
Public Const NNTP_REPLY_NEWARTICLES = 230
Public Const NNTP_REPLY_NEWGROUPS = 231
Public Const NNTP_REPLY_TRANSFERRED = 235
Public Const NNTP_REPLY_POSTED = 240
Public Const NNTP_REPLY_GROUPTITLE = 282
Public Const NNTP_REPLY_XFERARTICLE = 335
Public Const NNTP_REPLY_AUTHCONT = 350
Public Const NNTP_REPLY_AUTHPASS = 381
Public Const NNTP_REPLY_POSTARTICLE = 340
Public Const NNTP_REPLY_AUTHSIMPLE = 450
Public Const NNTP_REPLY_AUTHORIGIN = 480

'
' NNTP client actions
'
Public Const NNTP_ACTION_NONE = 0
Public Const NNTP_ACTION_CONNECT = 1
Public Const NNTP_ACTION_DISCONNECT = 2
Public Const NNTP_ACTION_LISTGROUPS = 3
Public Const NNTP_ACTION_NEWGROUPS = 4
Public Const NNTP_ACTION_LISTARTICLES = 5
Public Const NNTP_ACTION_GETHEADER = 6
Public Const NNTP_ACTION_GETARTICLE = 7
Public Const NNTP_ACTION_NEXTARTICLE = 8
Public Const NNTP_ACTION_PREVARTICLE = 9
Public Const NNTP_ACTION_NEWARTICLE = 10
Public Const NNTP_ACTION_POSTARTICLE = 11
Public Const NNTP_ACTION_AUTHENTICATE = 12

'
' NNTP client states
'
Public Const NNTP_CLIENT_UNUSED = 0
Public Const NNTP_CLIENT_IDLE = 1
Public Const NNTP_CLIENT_CONNECTING = 2
Public Const NNTP_CLIENT_GROUPLIST = 3
Public Const NNTP_CLIENT_ARTICLELIST = 4
Public Const NNTP_CLIENT_READING = 5
Public Const NNTP_CLIENT_POSTING = 6

'
' Post Office Protocol (POP) client control constants
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event.
'
Public Const POP_ERROR = -1
Public Const ERROR_POP_HANDLE = 22301   ' invalid handle to pop client
Public Const ERROR_POP_PARAM = 22302    ' invalid paramter passed to function
Public Const ERROR_POP_INIT = 22303     ' unable to initialize socket library
Public Const ERROR_POP_ALLOC = 22304    ' unable to allocate memory
Public Const ERROR_POP_HOSTNAME = 22305 ' unable to resolve hostname
Public Const ERROR_POP_SOCKET = 22306   ' unable to create socket
Public Const ERROR_POP_CONNECT = 22307  ' unable to connect to remote host
Public Const ERROR_POP_SELECT = 22308   ' unable to select socket
Public Const ERROR_POP_TIMEOUT = 22309  ' timeout reading socket
Public Const ERROR_POP_SERVER = 22310   ' invalid POP server response
Public Const ERROR_POP_READ = 22311     ' unable to read socket
Public Const ERROR_POP_WRITE = 22312    ' unable to write socket
Public Const ERROR_POP_LOGIN = 22313    ' unable to login to server
Public Const ERROR_POP_NOLOGIN = 22314  ' not logged in to server
Public Const ERROR_POP_MESGID = 22315   ' invalid message identifier
Public Const ERROR_POP_CLOSED = 22316   ' server connection closed
Public Const ERROR_POP_BUSY = 22317     ' server is busy

'
' POP client actions
'
Public Const POP_ACTION_NONE = 0
Public Const POP_ACTION_CONNECT = 1
Public Const POP_ACTION_DISCONNECT = 2
Public Const POP_ACTION_GETHDR = 3
Public Const POP_ACTION_GETMSG = 4
Public Const POP_ACTION_DELMSG = 5
Public Const POP_ACTION_RESET = 6

'
' POP client states
'
Public Const POP_CLIENT_UNUSED = 0
Public Const POP_CLIENT_IDLE = 1
Public Const POP_CLIENT_CONNECTING = 2
Public Const POP_CLIENT_LOGIN = 3
Public Const POP_CLIENT_TRANSFER = 4

'
' Simple Mail Transfer Protocol (SMTP) client control constants
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event.
'
Public Const SMTP_ERROR = -1
Public Const ERROR_SMTP_HANDLE = 22201  ' invalid handle to SMTP client
Public Const ERROR_SMTP_PARAM = 22202   ' invalid paramter passed to function
Public Const ERROR_SMTP_INIT = 22203    ' unable to initialize socket library
Public Const ERROR_SMTP_ALLOC = 22204   ' unable to allocate memory
Public Const ERROR_SMTP_HOSTNAME = 22205 ' unable to resolve hostname
Public Const ERROR_SMTP_SOCKET = 22206  ' unable to create socket
Public Const ERROR_SMTP_CONNECT = 22207 ' unable to connect to remote host
Public Const ERROR_SMTP_SELECT = 22208  ' unable to select socket
Public Const ERROR_SMTP_TIMEOUT = 22209 ' timeout reading socket
Public Const ERROR_SMTP_SERVER = 22210  ' invalid SMTP server response
Public Const ERROR_SMTP_READ = 22211    ' unable to read socket
Public Const ERROR_SMTP_WRITE = 22212   ' unable to write socket
Public Const ERROR_SMTP_BUSY = 22213    ' server connection in use
Public Const ERROR_SMTP_DATA = 22214    ' server receiving message data
Public Const ERROR_SMTP_NODATA = 22215  ' no data received
Public Const ERROR_SMTP_CLOSED = 22216  ' server connection closed
Public Const ERROR_SMTP_CANCEL = 22217  ' client operation was canceled

'
' SMTP client actions
'
Public Const SMTP_ACTION_NONE = 0
Public Const SMTP_ACTION_CONNECT = 1
Public Const SMTP_ACTION_DISCONNECT = 2
Public Const SMTP_ACTION_ADDRESS = 3
Public Const SMTP_ACTION_SENDMAIL = 4
Public Const SMTP_ACTION_EXPAND = 5
Public Const SMTP_ACTION_VERIFY = 6
Public Const SMTP_ACTION_RESET = 7
Public Const SMTP_ACTION_CANCEL = 8

'
' SMTP client states
'
Public Const SMTP_CLIENT_UNUSED = 0
Public Const SMTP_CLIENT_IDLE = 1
Public Const SMTP_CLIENT_CONNECTING = 2
Public Const SMTP_CLIENT_ADDRESSING = 3
Public Const SMTP_CLIENT_COMPOSING = 4
Public Const SMTP_CLIENT_SENDING = 5

'
' Remote Access Service (RAS) control contants
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event. These are the
' same codes returned by the RAS library, with 25000 added
' to the base value
'
Public Const ERROR_INVALID_PORT_HANDLE = 25601
Public Const ERROR_PORT_ALREADY_OPEN = 25602
Public Const ERROR_BUFFER_TOO_SMALL = 25603
Public Const ERROR_WRONG_INFO_SPECIFIED = 25604
Public Const ERROR_CANNOT_SET_PORT_INFO = 25605
Public Const ERROR_PORT_NOT_CONNECTED = 25606
Public Const ERROR_EVENT_INVALID = 25607
Public Const ERROR_DEVICE_DOES_NOT_EXIST = 25608
Public Const ERROR_DEVICETYPE_DOES_NOT_EXIST = 25609
Public Const ERROR_INVALID_BUFFER = 25610
Public Const ERROR_ROUTE_NOT_AVAILABLE = 25611
Public Const ERROR_ROUTE_NOT_ALLOCATED = 25612
Public Const ERROR_INVALID_COMPRESSION_SPECIFIED = 25613
Public Const ERROR_OUT_OF_BUFFERS = 25614
Public Const ERROR_PORT_NOT_FOUND = 25615
Public Const ERROR_ASYNC_REQUEST_PENDING = 25616
Public Const ERROR_ALREADY_DISCONNECTING = 25617
Public Const ERROR_PORT_NOT_OPEN = 25618
Public Const ERROR_PORT_DISCONNECTED = 25619
Public Const ERROR_NO_ENDPOINTS = 25620
Public Const ERROR_CANNOT_OPEN_PHONEBOOK = 25621
Public Const ERROR_CANNOT_LOAD_PHONEBOOK = 25622
Public Const ERROR_CANNOT_FIND_PHONEBOOK_ENTRY = 25623
Public Const ERROR_CANNOT_WRITE_PHONEBOOK = 25624
Public Const ERROR_CORRUPT_PHONEBOOK = 25625
Public Const ERROR_CANNOT_LOAD_STRING = 25626
Public Const ERROR_KEY_NOT_FOUND = 25627
Public Const ERROR_DISCONNECTION = 25628
Public Const ERROR_REMOTE_DISCONNECTION = 25629
Public Const ERROR_HARDWARE_FAILURE = 25630
Public Const ERROR_USER_DISCONNECTION = 25631
Public Const ERROR_INVALID_SIZE = 25632
Public Const ERROR_PORT_NOT_AVAILABLE = 25633
Public Const ERROR_CANNOT_PROJECT_CLIENT = 25634
Public Const ERROR_UNKNOWN = 25635
Public Const ERROR_WRONG_DEVICE_ATTACHED = 25636
Public Const ERROR_BAD_STRING = 25637
Public Const ERROR_REQUEST_TIMEOUT = 25638
Public Const ERROR_CANNOT_GET_LANA = 25639
Public Const ERROR_NETBIOS_ERROR = 25640
Public Const ERROR_SERVER_OUT_OF_RESOURCES = 25641
Public Const ERROR_NAME_EXISTS_ON_NET = 25642
Public Const ERROR_SERVER_GENERAL_NET_FAILURE = 25643
Public Const ERROR_AUTH_INTERNAL = 25645
Public Const ERROR_RESTRICTED_LOGON_HOURS = 25646
Public Const ERROR_ACCT_DISABLED = 25647
Public Const ERROR_PASSWD_EXPIRED = 25648
Public Const ERROR_NO_DIALIN_PERMISSION = 25649
Public Const ERROR_SERVER_NOT_RESPONDING = 25650
Public Const ERROR_FROM_DEVICE = 25651
Public Const ERROR_UNRECOGNIZED_RESPONSE = 25652
Public Const ERROR_MACRO_NOT_FOUND = 25653
Public Const ERROR_MACRO_NOT_DEFINED = 25654
Public Const ERROR_MESSAGE_MACRO_NOT_FOUND = 25655
Public Const ERROR_DEFAULTOFF_MACRO_NOT_FOUND = 25656
Public Const ERROR_FILE_COULD_NOT_BE_OPENED = 25657
Public Const ERROR_DEVICENAME_TOO_LONG = 25658
Public Const ERROR_DEVICENAME_NOT_FOUND = 25659
Public Const ERROR_NO_RESPONSES = 25660
Public Const ERROR_NO_COMMAND_FOUND = 25661
Public Const ERROR_WRONG_KEY_SPECIFIED = 25662
Public Const ERROR_UNKNOWN_DEVICE_TYPE = 25663
Public Const ERROR_ALLOCATING_MEMORY = 25664
Public Const ERROR_PORT_NOT_CONFIGURED = 25665
Public Const ERROR_DEVICE_NOT_READY = 25666
Public Const ERROR_READING_INI_FILE = 25667
Public Const ERROR_NO_CONNECTION = 25668
Public Const ERROR_BAD_USAGE_IN_INI_FILE = 25669
Public Const ERROR_READING_SECTIONNAME = 25670
Public Const ERROR_READING_DEVICETYPE = 25671
Public Const ERROR_READING_DEVICENAME = 25672
Public Const ERROR_READING_USAGE = 25673
Public Const ERROR_READING_MAXCONNECTBPS = 25674
Public Const ERROR_READING_MAXCARRIERBPS = 25675
Public Const ERROR_LINE_BUSY = 25676
Public Const ERROR_VOICE_ANSWER = 25677
Public Const ERROR_NO_ANSWER = 25678
Public Const ERROR_NO_CARRIER = 25679
Public Const ERROR_NO_DIALTONE = 25680
Public Const ERROR_IN_COMMAND = 25681
Public Const ERROR_WRITING_SECTIONNAME = 25682
Public Const ERROR_WRITING_DEVICETYPE = 25683
Public Const ERROR_WRITING_DEVICENAME = 25684
Public Const ERROR_WRITING_MAXCONNECTBPS = 25685
Public Const ERROR_WRITING_MAXCARRIERBPS = 25686
Public Const ERROR_WRITING_USAGE = 25687
Public Const ERROR_WRITING_DEFAULTOFF = 25688
Public Const ERROR_READING_DEFAULTOFF = 25689
Public Const ERROR_EMPTY_INI_FILE = 25690
Public Const ERROR_AUTHENTICATION_FAILURE = 25691
Public Const ERROR_PORT_OR_DEVICE = 25692
Public Const ERROR_NOT_BINARY_MACRO = 25693
Public Const ERROR_DCB_NOT_FOUND = 25694
Public Const ERROR_STATE_MACHINES_NOT_STARTED = 25695
Public Const ERROR_STATE_MACHINES_ALREADY_STARTED = 25696
Public Const ERROR_PARTIAL_RESPONSE_LOOPING = 25697
Public Const ERROR_UNKNOWN_RESPONSE_KEY = 25698
Public Const ERROR_RECV_BUF_FULL = 25699
Public Const ERROR_CMD_TOO_LONG = 25700
Public Const ERROR_UNSUPPORTED_BPS = 25701
Public Const ERROR_UNEXPECTED_RESPONSE = 25702
Public Const ERROR_INTERACTIVE_MODE = 25703
Public Const ERROR_BAD_CALLBACK_NUMBER = 25704
Public Const ERROR_INVALID_AUTH_STATE = 25705
Public Const ERROR_WRITING_INITBPS = 25706
Public Const ERROR_INVALID_WIN_HANDLE = 25707
Public Const ERROR_NO_PASSWORD = 25708
Public Const ERROR_NO_USERNAME = 25709
Public Const ERROR_CANNOT_START_STATE_MACHINE = 25710
Public Const ERROR_GETTING_COMMSTATE = 25711
Public Const ERROR_SETTING_COMMSTATE = 25712
Public Const ERROR_COMM_FUNCTION = 25713
Public Const ERROR_CONFIGURATION_PROBLEM = 25714
Public Const ERROR_X25_DIAGNOSTIC = 25715
Public Const ERROR_TOO_MANY_LINE_ERRORS = 25716
Public Const ERROR_OVERRUN = 25717
Public Const ERROR_ACCT_EXPIRED = 25718
Public Const ERROR_CHANGING_PASSWORD = 25719
Public Const ERROR_NO_ACTIVE_ISDN_LINES = 25720
Public Const ERROR_NO_ISDN_CHANNELS_AVAILABLE = 25721

'
' Remote command client actions
'
Public Const RCMD_ACTION_NONE = 0
Public Const RCMD_ACTION_LOGIN = 1
Public Const RCMD_ACTION_EXECUTE = 2
Public Const RCMD_ACTION_COMMAND = 3
Public Const RCMD_ACTION_DISCONNECT = 4

'
' Remote command client states
'
Public Const RCMD_CLIENT_UNUSED = 0
Public Const RCMD_CLIENT_IDLE = 1
Public Const RCMD_CLIENT_CONNECTING = 2
Public Const RCMD_CLIENT_RECEIVING = 3
Public Const RCMD_CLIENT_SENDING = 4
Public Const RCMD_CLIENT_CLOSING = 5

'
' Telnet (Network Terminal) client control constants
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event.
'
Public Const TELNET_ERROR = -1
Public Const ERROR_TELNET_NOHANDLES = 22101   ' no free handles
Public Const ERROR_TELNET_ALLOCMEM = 22102    ' no Public memory
Public Const ERROR_TELNET_NOTOWNER = 22103    ' not session owner
Public Const ERROR_TELNET_NETSTART = 22104    ' cannot initialize network
Public Const ERROR_TELNET_HOSTNAME = 22105    ' invalid hostname
Public Const ERROR_TELNET_SOCKET = 22106      ' cannot create socket
Public Const ERROR_TELNET_CONNECT = 22108     ' cannot connect to server
Public Const ERROR_TELNET_PEERNAME = 22109    ' cannot get peer information
Public Const ERROR_TELNET_PARAM = 22112       ' invalid parameter
Public Const ERROR_TELNET_HANDLE = 22113      ' invalid TELNET handle
Public Const ERROR_TELNET_SELECT = 22114      ' socket select error
Public Const ERROR_TELNET_READ = 22115        ' cannot read from socket
Public Const ERROR_TELNET_RDBUF = 22117       ' receive buffer is full
Public Const ERROR_TELNET_WRITE = 22118       ' cannot write to socket
Public Const ERROR_TELNET_WRBUF = 22119       ' send buffer is empty
Public Const ERROR_TELNET_BUSY = 22120        ' connection in use
Public Const ERROR_TELNET_NOTBUSY = 22121     ' connection is not in use
Public Const ERROR_TELNET_TIMEOUT = 22122     ' timeout waiting for socket
Public Const ERROR_TELNET_SYNC = 22123        ' internal synchronization error
Public Const ERROR_TELNET_OPTION = 22124      ' unsupported TELNET option
Public Const ERROR_TELNET_CLOSED = 22125      ' server closed connection

'
' Telnet client actions
'
Public Const TELNET_ACTION_CONNECT = 1
Public Const TELNET_ACTION_DISCONNECT = 2

'
' Telnet client states
'
Public Const TELNET_CLIENT_UNUSED = 0
Public Const TELNET_CLIENT_IDLE = 1
Public Const TELNET_CLIENT_CONNECTING = 2
Public Const TELNET_CLIENT_RECEIVING = 3
Public Const TELNET_CLIENT_SENDING = 4
Public Const TELNET_CLIENT_CLOSING = 5

'
' World Wide Web control contants
'
' The error codes are returned by the LastError property and
' passed as an argument to the LastError event.
'
Public Const HTTP_ERROR = -1
Public Const ERROR_HTTP_HANDLE = 22401       ' invalid handle to client
Public Const ERROR_HTTP_PARAM = 22402        ' invalid paramter
Public Const ERROR_HTTP_INIT = 22403         ' unable to initialize library
Public Const ERROR_HTTP_ALLOC = 22404        ' unable to allocate memory
Public Const ERROR_HTTP_HOSTNAME = 22405     ' unable to resolve hostname
Public Const ERROR_HTTP_SOCKET = 22406       ' unable to create socket
Public Const ERROR_HTTP_CONNECT = 22407      ' unable to connect to remote host
Public Const ERROR_HTTP_SELECT = 22408       ' unable to select socket
Public Const ERROR_HTTP_TIMEOUT = 22409      ' timeout reading socket
Public Const ERROR_HTTP_SERVER = 22410       ' invalid HTTP server response
Public Const ERROR_HTTP_NOFILE = 22411       ' cannot open local file
Public Const ERROR_HTTP_OPEN = 22412         ' cannot open local file
Public Const ERROR_HTTP_READ = 22413         ' unable to read socket
Public Const ERROR_HTTP_WRITE = 22414        ' unable to write socket
Public Const ERROR_HTTP_EOF = 22415          ' end-of-file on socket
Public Const ERROR_HTTP_BUSY = 22416         ' server is busy
Public Const ERROR_HTTP_MOVED = 22417        ' resource has been moved
Public Const ERROR_HTTP_NOTMODIFIED = 22418  ' resource has not been modified
Public Const ERROR_HTTP_BADREQUEST = 22419   ' bad request to server
Public Const ERROR_HTTP_UNAUTHORIZED = 22420 ' unauthorized request to server
Public Const ERROR_HTTP_PAYMENT = 22421      ' payment required for resource
Public Const ERROR_HTTP_FORBIDDEN = 22422    ' access to resource forbidden
Public Const ERROR_HTTP_NOTFOUND = 22423     ' resource not found on server
Public Const ERROR_HTTP_BADMETHOD = 22424    ' invalid method for this resource
Public Const ERROR_HTTP_NOTACCEPTED = 22425  ' resource not acceptable to client
Public Const ERROR_HTTP_PROXYAUTH = 22426    ' proxy authentication required
Public Const ERROR_HTTP_GONE = 22427         ' resource is permanently removed
Public Const ERROR_HTTP_INTERNAL = 22428     ' internal server error
Public Const ERROR_HTTP_NOMETHOD = 22429     ' method not supported by server
Public Const ERROR_HTTP_BADGATEWAY = 22430   ' invalid response from gateway
Public Const ERROR_HTTP_UNAVAILABLE = 22431  ' server is unavailable
Public Const ERROR_HTTP_GATETIMEOUT = 22432  ' server gateway timeout
Public Const ERROR_HTTP_UNEXPECTED = 22433   ' unexpected error

'
' Web client actions
'
Public Const WEB_ACTION_LOAD = 1
Public Const WEB_ACTION_RELOAD = 2
Public Const WEB_ACTION_UNLOAD = 3
Public Const WEB_ACTION_NEXT = 4
Public Const WEB_ACTION_PREV = 5
Public Const WEB_ACTION_CANCEL = 6
Public Const WEB_ACTION_RESET = 7
Public Const WEB_ACTION_REFRESH = 8
Public Const WEB_ACTION_SEARCH = 9

'
' Web client states
'
Public Const WEB_CLIENT_UNUSED = 0
Public Const WEB_CLIENT_IDLE = 1
Public Const WEB_CLIENT_CONNECT = 2
Public Const WEB_CLIENT_COMMAND = 3
Public Const WEB_CLIENT_LOAD = 4
Public Const WEB_CLIENT_UNLOAD = 5

'
' Web client resource types
'
Public Const WEB_RESOURCE_UNKNOWN = 0
Public Const WEB_RESOURCE_HTML = 1
Public Const WEB_RESOURCE_TEXT = 2
Public Const WEB_RESOURCE_IMAGE = 3

'
' SocketWrench error response values
'
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

'
' SocketWrench control actions
'
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_ABORT = 8

'
' SocketWrench control states
'
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

'
' Address families
'
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

'
' Socket types
'
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

'
' Protocol types
'
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256

'
' Common ports
'
Public Const IPPORT_ANY = 0
Public Const IPPORT_ECHO = 7
Public Const IPPORT_DISCARD = 9
Public Const IPPORT_SYSTAT = 11
Public Const IPPORT_DAYTIME = 13
Public Const IPPORT_NETSTAT = 15
Public Const IPPORT_FTP = 21
Public Const IPPORT_TELNET = 23
Public Const IPPORT_SMTP = 25
Public Const IPPORT_TIMESERVER = 37
Public Const IPPORT_NAMESERVER = 42
Public Const IPPORT_WHOIS = 43
Public Const IPPORT_MTP = 57
Public Const IPPORT_FINGER = 79
Public Const IPPORT_TFTP = 69
Public Const IPPORT_RESERVED = 1024
Public Const IPPORT_USERRESERVED = 5000

'
' Network addresses
'
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.255.255.255"

'
' Shutdown values
'
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

'
' Declarations for functions to encode and decode files, typically
' used as with attachments to mail messages or news articles
'
#If Win32 Then
Declare Function DecodeFile Lib "UUCODE32.DLL" (ByVal FileName As String) As Long
Declare Function EncodeFile Lib "UUCODE32.DLL" (ByVal InputFile As String, ByVal OutputFile As String) As Long
Declare Function DecodeBase64File Lib "UUCODE32.DLL" (ByVal InputFile As String, ByVal OutputFile As String) As Long
Declare Function EncodeBase64File Lib "UUCODE32.DLL" (ByVal InputFile As String, ByVal OutputFile As String) As Long
#Else
Declare Function DecodeFile Lib "UUCODE16.DLL" (ByVal FileName As String) As Integer
Declare Function EncodeFile Lib "UUCODE16.DLL" (ByVal InputFile As String, ByVal OutputFile As String) As Integer
Declare Function DecodeBase64File Lib "UUCODE16.DLL" (ByVal InputFile As String, ByVal OutputFile As String) As Integer
Declare Function EncodeBase64File Lib "UUCODE16.DLL" (ByVal InputFile As String, ByVal OutputFile As String) As Integer
#End If

