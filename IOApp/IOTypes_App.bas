Attribute VB_Name = "IOTypes_App"
Option Explicit

' Mirror of the simulator's IOTypes.bas array bounds. When Phase 3 adds
' Profibus reads/writes here, this lets us reuse Set_*_Output / Get_*_Input
' code from the simulator with the same address mapping.

Public Const REALADDR_DI As Integer = 2400
Public Const REALADDR_DO As Integer = 2560
Public Const REALADDR_AO As Integer = 280
Public Const REALADDR_AI As Integer = 64

Public Const ABSTRACTADDR_DI As Integer = 1600
Public Const ABSTRACTADDR_DO As Integer = 1440
Public Const ABSTRACTADDR_AI As Integer = 136
Public Const ABSTRACTADDR_AO As Integer = 220

Public Const NO_D_OUTPUT = REALADDR_DO + ABSTRACTADDR_DO  ' 4000
Public Const NO_D_INPUT = REALADDR_DI + ABSTRACTADDR_DI   ' 4000
Public Const NO_A_OUTPUT = REALADDR_AO + ABSTRACTADDR_AO  ' 500
Public Const NO_A_INPUT = REALADDR_AI + ABSTRACTADDR_AI   ' 200

Public A_INPUT(NO_A_INPUT) As Byte
Public A_OUTPUT(NO_A_OUTPUT) As Byte
Public D_INPUT(NO_D_INPUT) As Byte
Public D_OUTPUT(NO_D_OUTPUT) As Byte

' Python server endpoints (this app is the inverse of the simulator -
' it READS aout/dout from /ioio/status and WRITES ain/din to /ioio/inputs).
Public Const PYTHON_STATUS_URL As String = "http://127.0.0.1:8080/ioio/status"
Public Const PYTHON_INPUTS_URL As String = "http://127.0.0.1:8080/ioio/inputs"

Public ioStatusHttp As MSXML2.XMLHTTP
Public ioInputsHttp As MSXML2.XMLHTTP
Public StatusHandler As IOAppHandler
