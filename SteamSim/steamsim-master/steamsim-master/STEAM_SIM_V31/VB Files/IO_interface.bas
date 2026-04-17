Attribute VB_Name = "IO_interface"
'26/2/97 hBoard is used elsewhere not hBoard%        Global hBoard%
Global hBoard As Integer
Global hS%, hE%, hB%
Global chip%, chan%, DIOchip%
Global flash_count As Integer

' 28/2/97 jbf declare values for dll calls
Global isInput As Integer               ' config port for input
Global isOutput As Integer              ' config port for output
Global Adata As Long                    ' data on port A
Global Bdata As Long                    ' data on port A
Global Cdata As Long                    ' data on port A
Global Cdata_L As Long
Global Cdata_H As Long

' Boards supported...
Global Const PC212E = 212
Global Const PC214E = 214
Global Const PC215E = 215
Global Const PC218E = 218
Global Const PC272E = 272

Global Const PC226E = 226
Global Const PC30AT = 30
Global Const PC26AT = 26
Global Const PC27E = 27

' Group base address offsets...
Global Const PPIX = 0                   ' PPI X
Global Const PPIY = 8                   ' PPI Y
Global Const PPIZ = 16                  ' PPI Z
Global Const X1 = 0                     ' Timer X1
Global Const X2 = 4                     ' Timer X2
Global Const Y1 = 8                     ' Timer Y1
Global Const Y2 = 12                    ' Timer Y2
Global Const Z1 = 16                    ' Timer Z1
Global Const Z2 = 20                    ' Timer Z2

' PPI ports for setting mode...
Global Const PORTA = 0
Global Const PORTB = 1
Global Const PORTC_L = 2
Global Const PORTC_H = 3
'27/2/97 seems PORTC_L does whole byte anyway
Global Const PORTC = 2

' Timer counter clock sources...
Global Const CLK_CLK = 0
Global Const CLK_10MHZ = 1
Global Const CLK_1MHZ = 2
Global Const CLK_100KHZ = 3
Global Const CLK_10KHZ = 4
Global Const CLK_1KHZ = 5
Global Const CLK_OUTN_1 = 6
Global Const CLK_EXT = 7

' Timer counter gate sources...
Global Const GAT_VCC = 0
Global Const GAT_GND = 1
Global Const GAT_EXT = 2
Global Const GAT_OUTN_2 = 3
 
' Error codes...
Global Const OK = 0
Global Const ERRSUPPORT = -1
Global Const ERRBASE = -2
Global Const ERRIRQ = -3
Global Const ERRHANDLE = -4
Global Const ERRCHAN = -5
Global Const ERRDATA = -6
Global Const ERRRANGE = -7
Global Const ERRMEMORY = -8
Global Const ERRBUFFER = -9
Global Const ERRPC226 = -10

Declare Function registerBoard% Lib "dio_tc.dll" (ByVal model%, ByVal ba%, ByVal irq%)
Declare Function GetBoardModel% Lib "dio_tc.dll" (ByVal h%)
Declare Function FreeBoard% Lib "dio_tc.dll" (ByVal h%)

Declare Function setIntMask% Lib "dio_tc.dll" (ByVal h%, ByVal mask%)
Declare Function getIntStat% Lib "dio_tc.dll" (ByVal h%)
Declare Function enableInterrupts% Lib "dio_tc.dll" (ByVal h%)
Declare Function disableInterrupts% Lib "dio_tc.dll" (ByVal h%)

Declare Function allocateLongBuf% Lib "dio_tc.dll" (ByVal nItems&)
Declare Function freeLongBuf% Lib "dio_tc.dll" (ByVal b%)
Declare Function readLongBuf% Lib "dio_tc.dll" (ByVal b%, ByVal item&, p&)
Declare Function writeLongBuf% Lib "dio_tc.dll" (ByVal b%, ByVal item&, ByVal dat&)
Declare Function copyToLongBuf% Lib "dio_tc.dll" (ByVal b%, ByVal start&, ByVal nItems&, p&)
Declare Function copyFromLongBuf% Lib "dio_tc.dll" (ByVal b%, ByVal start&, ByVal nItems&, p&)
Declare Function getLongIntItem% Lib "dio_tc.dll" (ByVal b%, item&)

Declare Function allocateIntegerBuf% Lib "dio_tc.dll" (ByVal nItems%)
Declare Function freeIntegerBuf% Lib "dio_tc.dll" (ByVal b%)
Declare Function readIntegerBuf% Lib "dio_tc.dll" (ByVal b%, ByVal item&, p%)
Declare Function writeIntegerBuf% Lib "dio_tc.dll" (ByVal b%, ByVal item&, ByVal dat%)
Declare Function copyToIntegerBuf% Lib "dio_tc.dll" (ByVal b%, ByVal start&, ByVal nItems&, p%)
Declare Function copyFromIntegerBuf% Lib "dio_tc.dll" (ByVal b%, ByVal start&, ByVal nItems&, p%)
Declare Function getIntegerIntItem% Lib "dio_tc.dll" (ByVal b%, item&)

Declare Function TCisAvailable% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%)
Declare Function TCfreeResource% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%)
Declare Function TCsetClock% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal clk%)
Declare Function TCsetGate% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal gat%)
Declare Function TCsetMode% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal mde%)
Declare Function TCgetStatus% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%)
Declare Function TCsetCount% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal count&)
Declare Function TCgetCount% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, count&)
Declare Function TCgetUpCount% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, count&)

Declare Function TCsetDiffCounters% Lib "dio_tc.dll" (ByVal h%, ByVal chip1%, ByVal chan1%, ByVal clk1%, ByVal gat1%, ByVal chip2%, ByVal chan2%, ByVal clk2%, ByVal gat2%)
Declare Function TCgetDiffCount% Lib "dio_tc.dll" (ByVal h%, ByVal hD%, diff&)
Declare Function TCgetRatio% Lib "dio_tc.dll" (ByVal h%, ByVal hD%, ratio!)
Declare Function TCfreeDiffCounters% Lib "dio_tc.dll" (ByVal h%, ByVal hD%)

Declare Function TCsetMonoShot% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal duration!)
Declare Function TCsetAstable% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal chipS%, ByVal chanS%, ByVal freq!, ByVal msratio!)
Declare Function TCfreeAstable% Lib "dio_tc.dll" (ByVal h%, ByVal ha%)

Declare Function TCsetStopwatch% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%)
Declare Function TCstartStopwatch% Lib "dio_tc.dll" (ByVal h%, ByVal hS%)
Declare Function TCpauseStopwatch% Lib "dio_tc.dll" (ByVal h%, ByVal hS%)
Declare Function TCgetElapsedTime% Lib "dio_tc.dll" (ByVal h%, ByVal hS%, lPtr&)
Declare Function TCsetEventRecorder% Lib "dio_tc.dll" (ByVal h%, ByVal hS%, ByVal chip%, ByVal hB%)
Declare Function TCfreeEventRecorder% Lib "dio_tc.dll" (ByVal h%, ByVal hE%)
Declare Function TCgetTimeStr% Lib "dio_tc.dll" (ByVal ms&, strPtr$)
Declare Function TCfreeStopwatch% Lib "dio_tc.dll" (ByVal h%, ByVal hS%)

Declare Function TCgetExtPeriod% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, fPtr!)
Declare Function TCgetExtFreq% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, fPtr&)
Declare Function TCgenerateFreq% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal freq!)
Declare Function TCgenerateAccFreq% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal freq!)
Declare Function TCmultiplyFreq% Lib "dio_tc.dll" (ByVal h%, ByVal ipChip%, ByVal ipChan%, ByVal opChip%, ByVal opChan%, ByVal factor!)
Declare Function TCdivideFreq% Lib "dio_tc.dll" (ByVal h%, ByVal ipChip%, ByVal ipChan%, ByVal opChip%, ByVal opChan%, ByVal divisor!)

Declare Function TCsetDCO% Lib "dio_tc.dll" (ByVal h%, ByVal diChip%, ByVal diChan%, ByVal opChip%, ByVal opChan%, ByVal udFreq!, ByVal udChip!, ByVal MinF!, ByVal MaxF!)
Declare Function TCsetVCO% Lib "dio_tc.dll" (ByVal h%, ByVal AImodel%, ByVal AIbaseAddr%, ByVal AIchan%, ByVal opChip%, ByVal opChan%, ByVal udFreq!, ByVal udChip!, ByVal freq0V!, ByVal freq10V!)
Declare Function TCfreeDCO% Lib "dio_tc.dll" (ByVal h%, ByVal hO%)

Declare Function DIOisAvailable% Lib "dio_tc.dll" (ByVal h%, ByVal chip%)
Declare Function DIOsetMode% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal isInput%)
Declare Function DIOsetChanWidth% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal numBits%)
Declare Function DIOsetData% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, ByVal dat&)
Declare Function DIOgetData% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal chan%, dat&)

Declare Function DIOsetSwitchMatrix% Lib "dio_tc.dll" (ByVal h%, ByVal order%)
Declare Function DIOgetSwitchStatus% Lib "dio_tc.dll" (ByVal h%, ByVal xcoord%, ByVal ycoord%)
Declare Function DIOfreeSwitchMatrix% Lib "dio_tc.dll" (ByVal h%)

Declare Function DIOsetBiDirectionalBus% Lib "dio_tc.dll" (ByVal h%, ByVal chip%, ByVal isPBip%, ByVal isPCLip%, ByVal bIn%, ByVal bOut%)
Declare Function DIOsendBDBdata% Lib "dio_tc.dll" (ByVal h%, ByVal hI%)
Declare Function DIOfreeBiDirectionalBus% Lib "dio_tc.dll" (ByVal h%, ByVal hI%)


