Attribute VB_Name = "IOTypes_App"
Option Explicit

' Mirror of the simulator's IOTypes.bas array bounds, address constants,
' and per-card state. Lets the I/O App use the same Set_*_Output /
' Get_*_Input wrappers (in IOAppProfibus.bas) with the same address
' mapping as the simulator did when it had Profibus enabled.

' === Profibus card numbers ===
Public Const WINCARD_A As Integer = 1
Public Const WINCARD_B As Integer = 2

' === Per-block channel counts ===
Public Const NO_D_INPUT_CHANNELS As Integer = 32
Public Const NO_D_OUTPUT_CHANNELS As Integer = 64
Public Const NO_A_INPUT_CHANNELS As Integer = 4
Public Const NO_A_OUTPUT_CHANNELS As Integer = 4

' === Card A block address ranges ===
Public Const D_OUTPUT_START_ADDR As Integer = 60
Public Const D_OUTPUT_END_ADDR As Integer = 78
Public Const D_INPUT_START_ADDR As Integer = 80
Public Const D_INPUT_END_ADDR As Integer = 93
Public Const A_OUTPUT_START_ADDR As Integer = 1
Public Const A_OUTPUT_END_ADDR As Integer = 42
Public Const A_INPUT_START_ADDR As Integer = 50
Public Const A_INPUT_END_ADDR As Integer = 54

' === Card B block address ranges ===
Public Const D_OUTPUT_START_ADDR_B As Integer = 25
Public Const D_OUTPUT_END_ADDR_B As Integer = 37
Public Const D_INPUT_START_ADDR_B As Integer = 45
Public Const D_INPUT_END_ADDR_B As Integer = 56
Public Const A_OUTPUT_START_ADDR_B As Integer = 1
Public Const A_OUTPUT_END_ADDR_B As Integer = 18
Public Const A_INPUT_START_ADDR_B As Integer = 20
Public Const A_INPUT_END_ADDR_B As Integer = 21

' === Card B array offsets ===
Public Const AO_OFFSET As Integer = 200
Public Const AI_OFFSET As Integer = 40
Public Const DO_OFFSET As Integer = 1280
Public Const DI_OFFSET As Integer = 640

' === Total array sizes (real + abstract) ===
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

' === I/O arrays ===
Public A_INPUT(NO_A_INPUT) As Byte
Public A_OUTPUT(NO_A_OUTPUT) As Byte
Public D_INPUT(NO_D_INPUT) As Byte
Public D_OUTPUT(NO_D_OUTPUT) As Byte

' === Per-card state ===
' g_CardA_OK / g_CardB_OK are True when IO_Init succeeded for that card.
' If a card failed to initialize, the timer skips its reads and writes
' but still does the HTTP traffic.
Public g_CardA_OK As Boolean
Public g_CardB_OK As Boolean

' Lamp flash toggles, flipped each cycle by Set_D_Output.
Public Illuminate As Boolean
Public Illuminate2 As Boolean

' === Python server endpoints ===
' This app is the inverse of the simulator: it READS aout/dout from
' /ioio/status and WRITES ain/din to /ioio/inputs.
Public Const PYTHON_STATUS_URL As String = "http://127.0.0.1:8080/ioio/status"
Public Const PYTHON_INPUTS_URL As String = "http://127.0.0.1:8080/ioio/inputs"

Public ioStatusHttp As MSXML2.XMLHTTP
Public ioInputsHttp As MSXML2.XMLHTTP
Public StatusHandler As IOAppHandler
