Attribute VB_Name = "IOTypes"
Option Explicit
'card numbers
Global Const WINCARD_A = 1
Global Const WINCARD_B = 2

'block values
Global Const NO_D_INPUT_CHANNELS = 32
Global Const NO_D_OUTPUT_CHANNELS = 64
Global Const NO_A_INPUT_CHANNELS = 4
Global Const NO_A_OUTPUT_CHANNELS = 4

'io range address
'win card A
'set 14/7/99. Max Range 80 - 99
Global Const D_OUTPUT_START_ADDR = 60
Global Const D_OUTPUT_END_ADDR = 78 '78 CORRECT
'set 14/7/99. Max Range 60 - 79
Global Const D_INPUT_START_ADDR = 80
Global Const D_INPUT_END_ADDR = 93 '93 CORRECT
'set on 14/7/99.  Max Range 1 - 49
Global Const A_OUTPUT_START_ADDR = 1 '1 CORRECT
Global Const A_OUTPUT_END_ADDR = 42 '42 CORRECT
'set 14/7/99. Max Range 50 - 59
Global Const A_INPUT_START_ADDR = 50

' 15/11/99
'Global Const A_INPUT_END_ADDR = 53 '53 correct
Global Const A_INPUT_END_ADDR = 54 '53 correct


' Card 2 array offsets
'set 14/7/99
Global Const AO_OFFSET = 200
Global Const AI_OFFSET = 40
Global Const DO_OFFSET = 1280
Global Const DI_OFFSET = 640

'win card b
'set 15/7/99. Max Range 25 - 44
Global Const D_OUTPUT_START_ADDR_B = 25 ' 25 correct
Global Const D_OUTPUT_END_ADDR_B = 37  '37 correct
'set 15/7/99. Max Range 45 - 99
Global Const D_INPUT_START_ADDR_B = 45 '45 correct
Global Const D_INPUT_END_ADDR_B = 56 '56 correct
'set 15/7/99. Max Range 1 - 19
Global Const A_OUTPUT_START_ADDR_B = 1 '1 correct
Global Const A_OUTPUT_END_ADDR_B = 18 '18 correct
'set 15/7/99. Max Range 20 - 24
Global Const A_INPUT_START_ADDR_B = 20 '20 correct
Global Const A_INPUT_END_ADDR_B = 21 ' 21 correct

'io bounds for both cards. Includes memeory for free modules
'set 15/7/99
Global Const REALADDR_DI = 2400
Global Const REALADDR_DO = 2560
Global Const REALADDR_AO = 280
Global Const REALADDR_AI = 64

' 4/9/99   Make DI and DO arrays have 4000 elements  ( 0 to 4000 )
'Const ABSTRACTADDR_DI = 500
'Const ABSTRACTADDR_DO = 500
Const ABSTRACTADDR_DI = 1600     ' Total 2400 + 1600 = 4000
Const ABSTRACTADDR_DO = 1440     ' Total 2560 + 1440 = 4000

'  5/9/99
'Const ABSTRACTADDR_AI = 500      ' Total  64 + 500 = 564
'Const ABSTRACTADDR_AO = 500      ' Total 280 + 500 = 780
Const ABSTRACTADDR_AI = 136       ' Total  64 + 136 = 200
Const ABSTRACTADDR_AO = 220       ' Total 280 + 220 = 500

Global Const NO_D_OUTPUT = REALADDR_DO + ABSTRACTADDR_DO '2560 'NO_D_OUTPUT_CHANNELS * (((D_OUTPUT_END_ADDR - D_OUTPUT_START_ADDR) + 1) + ((D_OUTPUT_END_ADDR_B - D_OUTPUT_START_ADDR_B) + 1))
Global Const NO_D_INPUT = REALADDR_DI + ABSTRACTADDR_DI '2400  'NO_D_INPUT_CHANNELS * (((D_INPUT_END_ADDR - D_INPUT_START_ADDR) + 1) + ((D_INPUT_END_ADDR_B - D_INPUT_START_ADDR_B) + 1))
Global Const NO_A_OUTPUT = REALADDR_AO + ABSTRACTADDR_AO '280  'NO_A_OUTPUT_CHANNELS * (((A_OUTPUT_END_ADDR - A_OUTPUT_START_ADDR) + 1) + ((A_OUTPUT_END_ADDR_B - A_OUTPUT_START_ADDR_B) + 1))
Global Const NO_A_INPUT = REALADDR_AI + ABSTRACTADDR_AI '64    'NO_A_INPUT_CHANNELS * (((A_INPUT_END_ADDR - A_INPUT_START_ADDR) + 1) + ((A_INPUT_END_ADDR_B - A_INPUT_START_ADDR_B) + 1))

'ranges
Global Const AO_MAX = 255
Global Const AO_MIN = 0

'io data arrays
'25/10/99
Global VB_DI_Reload(0 To 3071) As Byte
'Global VB_AI_Reload(0 To 255) As Integer
Global VB_AI_Reload(0 To 255) As Byte




' even index in DOUTPUT means flash
' odd index in DOUTPUT means steady

'  5/9/99  ( Can I/O program handle this ?? )
'Global D_OUTPUT(NO_D_OUTPUT) As Boolean
'Global D_INPUT(NO_D_INPUT) As Boolean
Global D_OUTPUT(NO_D_OUTPUT) As Byte
Global D_INPUT(NO_D_INPUT) As Byte



Global A_OUTPUT(NO_A_OUTPUT) As Byte

Global A_INPUT(NO_A_INPUT) As Byte

' === 27/11/01 ===
Global A_INPUT_old(NO_A_INPUT) As Byte

Global ioioXmlHttp As MSXML2.XMLHTTP
Global profiIgnoreList As String
Global profiIgnoreListRead As Boolean

'15/11/99
Global OLD_A_input(NO_A_INPUT) As Byte
Global old_A_value As Integer
Global A_value As Integer

' variables
Global Illuminate As Boolean
Global Illuminate2 As Boolean
Global Changes(REALADDR_DI) As Boolean
'global AIChanges(realaddr_ai) as byte
