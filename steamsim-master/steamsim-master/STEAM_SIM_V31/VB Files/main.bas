Attribute VB_Name = "Main_module"


' === 07/03/07 ADDED ===================================
Global HOST_PC_address As String

' === 14/01/05 ===
'Global CLIENT_PC_address(10) As String
'Global CLIENT_PC_IS_ON_LINE(10) As String
Global CLIENT_PC_address(15) As String
Global CLIENT_PC_IS_ON_LINE(15) As String

' ================================================



' === 22/11/02 17.049 ===
Global Flashing_Flag As Boolean


' 11/12/98
Global cold_start_hourglass As Boolean

' 10/12/98
Global HARDWARE_flag As Boolean
Public Const DEFAULT_DISABLE_PROFIBUS As Boolean = True
Global DISABLE_PROFIBUS As Boolean

' 8/12/98
Global MDIform_loaded As Boolean

' 3/12/98
Global sy064_open As Boolean
Global sy071_open As Boolean
Global drain_mouse_down(200) As Boolean
'Global v_status(200) As Integer

' 14/9/98
Global DEMO_flag As Boolean
Global DEMO_halt_model As Boolean


' 6/7/98
Global turning_ahead As Integer

' 25/6/98
Global valve_mouse_flag As Integer

' 29/4/98
Global runIO_counter As Integer

' 10/4/98
Global reset_flag As Byte

' 6/4/98
Global Sounds_Enabled As Byte

' 3/4/98
Global AlarmChanges As Byte

' 2/4/98
Global clutch_flag As Byte

'31/3/98
Global old_OMD_count As Integer, OMD_count As Integer

' 26/3/98
Global Tx_delay As Single

' 10/3/98   SYNCH.FRM
Global AIPHASE As Single
Global Static_radius As Integer, Moving_radius As Integer
Global centreX As Integer, centreY As Integer
Global synch_count As Integer
Global synch_In_a_group As Integer
Global LOCK_SCOPE As Integer, RUNALL As Integer, FAIL_SCOPE As Integer
'Global test_flag As Integer
Global VRUN As Single, VINC As Single
Global FRUN As Single, fInc As Single

' 25/6/98
Global TG_speed As Single
Global TG_STEP As Single
Global TG_phase As Single
Global TG_RED_angle As Single, TG_BLUE_angle As Single, TG_YELLOW_angle As Single
Global TG_Aradians As Single
Global TG_centreX As Integer, TG_centreY As Integer
Global TG_Moving_radius As Integer
Global TG_select As Integer



Global Speed As Single, test_speed As Single

Global AIPHASE_model As Integer, AIPHASE_model_old As Integer
Global AI_STEP As Single
Global Aradians As Single
Global RED_angle As Single, BLUE_angle As Single, YELLOW_angle As Single
Global synchroscope As Integer


' 24/2/98 For now
' now in exec.bas Global LogFaults As Integer

' 24/2/98
'Global set_dummy_focus As Integer

Global Xval As Single
Global Yval As Single

' 13/1/98
Global PC_IN_CONTROL As Integer         '     -1=There is NO PC in CONTROL , 0=PC00 is in CONTROL , 1=PC01 is in CONTROL etc etc
Global This_PC_number As Integer         '      Read from  SIM.INI   ( e.g.SIM.INI=PC00 gives This_PC_number=00 )
Global This_Control_status As Integer    '      TRUE if this PC is in CONTROL , FALSE=This PC is NOT in CONTROL
Global I_AM_HOST As Integer                 '      TRUE if this PC is the HOST PC . ONLY the HOST PC will run the I/O software
'11/10/99
Global I_AM_INSTRUCTOR As Boolean

' 30/1/98
Global sep_water_low As Integer

' 28/1/98
Global greenness(10) As Integer, old_greenness(10) As Integer

' 27/1/98
Global redness(10) As Integer, old_redness(10) As Integer

' 2/2/98
Global izhset As Integer, izhset_old As Integer, izhset_sep_old As Integer
Global izhbunk1 As Integer, izhbunk1_old As Integer
Global izhbunk2 As Integer, izhbunk2_old As Integer

' 27/1/98
Global izhservice As Integer, izhservice_old As Integer

' 26/1/98
Global nRMA_Fuel As Integer, Fuel_temp_control As Integer

' 13/1/98
Global form_loading As Integer
' 13/1/98
Global response As Integer

' 6/1/98
Global medium_speed As Integer
Global old_medium_speed As Integer

' 6/1/98
Global LOSEP_flash As Integer
Global DOSEP_flash As Integer
Global HOSEP_flash As Integer

' 29/12/97
Global ial As Integer, jal As Integer, kal As Integer
Global alarm_count As Integer

'23/12/97
Global alarm_height As Integer
Global alarm_width As Integer

' 11/3/98
Global elec_ctrl_height As Integer
Global elec_ctrl_width As Integer

' 23/12/97 added:
'Global alarm_logged(350) As Integer, alarm_latest(351) As Integer
'Global alarm_time(351) As Date
Global alarm_logged(450) As Integer, alarm_latest(451) As Integer


'Global alarm_time(451) As Date
Global alarm_time(451) As String

'23/3/98
Global old_alarm_latest(451) As Integer

'Global current_time As Date
Global current_time As String


' 2/12/98
Global TARGET_this_sim As Integer
Global Const NOWHERE = 0      ' USE THIS for a MAXSIM type simulator (NO HARDWARE)
Global Const THAILAND = 1
Global Const PORTUGAL = 2
Global Const VIGO = 3
Global Const FERROL = 4


' 5/7/99
Global Const CALIFORNIA = 5

' === 20/12/01 Note: St_Peter is 6 and Sakhalin is 7 so.... ====
Global Const MARYLAND = 8

' === 31/05/02 ===
Global Const TEXAS = 9
Global Ballast_IS_included As Boolean


' 17/3/98
Declare Sub string_real Lib "c:\HAVEN\Model.DLL" (REAL_ARR As Single, str_to_real As Byte)


'18/3/98   ' 22/5/98 NOT USED:
'Declare Sub Get_Real_String Lib "c:\HAVEN\Model.DLL" (RtoS As Byte)



' 6/3/98
Declare Sub get_elec_addresses Lib "c:\HAVEN\Model.DLL" (ESTARTSW As Integer, _
ESTOPSW As Integer, EAUTOSW As Integer, EMANSW As Integer, ESTBY1SW As Integer, ESTBY2SW As Integer, _
AVRAUTOSW As Integer, AVRMANSW As Integer, MBKSW As Integer, OPENMBKSW As Integer, AVR_pot As Integer, _
SYNCHSELSW As Integer, CHSYNCHSW As Integer, SYNCHOFFSW As Integer, SELLAMPSSW As Integer, SELSYNCHSW As Integer, _
REVRESET As Integer)

Global ESTARTSW(2) As Integer
Global ESTOPSW(2) As Integer
Global EAUTOSW(2) As Integer
Global EMANSW(2) As Integer
Global ESTBY1SW(2) As Integer
Global ESTBY2SW(2) As Integer

Global AVRAUTOSW(4) As Integer
Global AVRMANSW(4) As Integer
Global MBKSW(4) As Integer
Global OPENMBKSW(4) As Integer
Global AVR_pot(4) As Integer
Global SYNCHSELSW(4) As Integer, CHSYNCHSW As Integer, SYNCHOFFSW As Integer, SELLAMPSSW As Integer, SELSYNCHSW As Integer
Global REVRESET(4) As Integer

' 29/12/97
Declare Sub set_HILO Lib "c:\HAVEN\Model.DLL" (nalarm As Integer, HILO_Value As Single)
Global nalarm As Integer, HILO_Value As Single
Global nAlarm_HIHI As Integer, nAlarm_HI As Integer, nAlarm_LO As Integer, nAlarm_LOLO As Integer

' 5/3/98
'Declare Sub get_HILO Lib "c:\HAVEN\Model.DLL" (ALARM_HILO As Single)
'Global ALARM_HILO(1 To 2, 1 To 350) As Single
Declare Sub get_HILO Lib "c:\HAVEN\Model.DLL" (ALARM_HILO As Single)
Global ALARM_HILO(1 To 2, 1 To 450) As Single
'16/11/99
Global Default_ALARM_HILO(1 To 2, 1 To 450) As Single
Global GotAlarmSetPoints As Boolean


' 29/12/97
Declare Sub set_DI Lib "c:\HAVEN\Model.DLL" (nDI As Integer, status_flag As Byte)
Global nDI As Integer
Global status_flag As Byte
Global set_DI_flag As Byte

' 17/6/98
Global DI_counter As Integer

' 23/12/97
' 14/12/98....
'Declare Sub get_alarm_text Lib "c:\HAVEN\Model.DLL" (ByVal alarm_no As Integer, ByVal alarm_text As String)

Global alarm_no As Integer
' 30/12/97 32 changed to 36
Global alarm_text As String * 36
' 5/3/98
'Global alarmtxt(1 To 350) As String * 36
Global alarmtxt(1 To 450) As String * 50



' 22/12/97 TESTING:
Declare Sub forstr Lib "c:\HAVEN\Model.DLL" (ByVal Bstring As String)
Global Bstring As String * 40

' 22/12/97
Global In_a_group As Integer

'30/4/97
Global RecordOn As Byte
Global RecordOff As Byte
Global ReplayOn As Byte
Global ReplayOff As Byte

' 1/4/97
Global tasks_running As Integer
' 26/3/97
Global fault_count As Integer
Global fault_in(20) As Integer

' 28/1/97
Global mouse_down As Integer, Mywidth As Integer, Myheight As Integer, Myscalewidth As Integer, Myscaleheight As Integer
Global resize_flag As Integer

' 19/3/98
Declare Sub get_R_string Lib "c:\HAVEN\Model.DLL" (ByVal R_string As String)
Declare Sub set_R_string Lib "c:\HAVEN\Model.DLL" (ByVal R_string As String)
Global R_string As String * 8004

' 23/3/98
Declare Sub get_I_string Lib "c:\HAVEN\Model.DLL" (ByVal I_string As String)
Declare Sub set_I_string Lib "c:\HAVEN\Model.DLL" (ByVal I_string As String)
Global I_string As String * 4002

'30/4/97
Declare Sub open_com Lib "c:\HAVEN\Model.DLL" (ncomm As Integer)


' === 23/01/02 ===
'Declare Sub D_MODEL Lib "c:\HAVEN\Model.DLL" (split_mode As Byte, ss_supply As Byte, diesel_init As Byte, elec_init As Byte, diesel_cold As Byte, elec_cold As Byte, d_run As Byte, e_run As Byte, snap_shot As Byte, hw_test As Byte)



' 8/12/98  18:57
'Global split_mode As Byte
'Global ss_supply As Byte
'Global diesel_init As Byte
'Global elec_init As Byte
'Global diesel_cold As Byte
'Global elec_cold As Byte
'Global d_run As Byte
'Global e_run As Byte
'Global snap_shot As Byte
'Global hw_test As Byte

' 8/12/98 18:57
Global split_mode As Byte, ss_supply As Byte, diesel_init As Byte, elec_init As Byte, diesel_cold As Byte, elec_cold As Byte, d_run As Byte, e_run As Byte, snap_shot As Byte, hw_test As Byte


' === 23/01/02 ===
'Declare Sub get_reals Lib "c:\HAVEN\Model.DLL" (RtoS As Byte, PCCONTROL As Single, PIDDATA As Single, VV_position As Single, PCA As Single, _
'PML As Single, ZML As Single, TML As Single, FML As Single, FLS As Single, ZLS As Single, TLS As Single, ZFT As Single, FMY As Single, ZMY As Single, PMY As Single, _
'PHT As Single, THT As Single, fHt As Single, ZHT As Single, _
'PLT As Single, TLT As Single, FLT As Single, ZLT As Single, _
'VB_ENGINE As Single, _
'PSW As Single, TSW As Single, FSW As Single, _
'PMF As Single, TMF As Single, FMF As Single, _
'PFT As Single, FFT As Single, TFT As Single, _
'ZBS As Single, _
'ALARM_HILO As Single, _
'elec_reals As Single, _
'plot_Reals As Single)


' 19/3/98
Global RtoS As Byte
Global PCCONTROL(80) As Single, PIDDATA(200) As Single, VV_position(20) As Single, PCA(15) As Single
Global PML(15) As Single, ZML(15) As Single, TML(8) As Single, FML(20) As Single
Global FLS(20) As Single, ZLS(8) As Single, TLS(8) As Single, ZFT(10) As Single
Global FMY(15) As Single, ZMY(4) As Single, PMY(12) As Single
Global PHT(20) As Single, THT(20) As Single, fHt(30) As Single, ZHT(5) As Single
Global PLT(20) As Single, TLT(20) As Single, FLT(20) As Single, ZLT(2) As Single

' 20/3/98
'Global VB_ENGINE(30) As Single
' 9/4/98
'Global VB_ENGINE(100) As Single
Global VB_ENGINE(150) As Single

Global PSW(20) As Single, TSW(15) As Single, FSW(25) As Single

Global PMF(30) As Single, TMF(30) As Single, FMF(30) As Single
Global PFT(30) As Single, TFT(30) As Single, FFT(30) As Single
Global ZBS(8) As Single

' 28/1/98
'Declare Sub get_logicals Lib "c:\HAVEN\Model.DLL" (VB_DO As Byte)
Global VB_DO(0 To 5119) As Byte


Global elec_reals(100) As Single


Declare Sub set_PCDI Lib "c:\HAVEN\Model.DLL" (nPCDI As Integer)
Global nPCDI As Integer

Declare Sub set_AI Lib "c:\HAVEN\Model.DLL" (nCHANNEL As Integer, nValue As Integer)
Global nCHANNEL As Integer, nValue As Integer

' 13/1/98
Declare Sub set_PC_CONTROL Lib "c:\HAVEN\Model.DLL" (nCHANNEL As Integer, Value As Single)
Global Value As Single


Declare Sub get_V_addresses Lib "c:\HAVEN\Model.DLL" (DI_VIO As Integer, DO_VIO As Integer)
Global DI_VIO(200) As Integer, DO_VIO(200) As Integer

Declare Sub get_VV_addresses Lib "c:\HAVEN\Model.DLL" (DI_ON_VVIO As Integer, DI_OFF_VVIO As Integer)
Global DI_ON_VVIO(30) As Integer, DI_OFF_VVIO(30) As Integer

Declare Sub get_P_addresses Lib "c:\HAVEN\Model.DLL" (PumpIO As Integer)
Global PumpIO(1 To 16, 1 To 55) As Integer
' 4/3/98
Declare Sub get_bkr_addresses Lib "c:\HAVEN\Model.DLL" (BKR As Integer)
Global BKR(1 To 11, 1 To 66) As Integer

Declare Sub get_RMA_addresses Lib "c:\HAVEN\Model.DLL" (RMAIO As Integer)
Global RMAIO(1 To 7, 1 To 8) As Integer


' 24/3/98
Declare Sub set_ACTION Lib "c:\HAVEN\Model.DLL" (nset As Integer, naction As Integer)
Global nset As Integer, naction As Integer
Declare Sub set_FAULT Lib "c:\HAVEN\Model.DLL" (nset As Integer, nFault As Integer)
Global nFault As Integer


' 23/3/98
Declare Sub get_integers Lib "c:\HAVEN\Model.DLL" (ItoS As Byte, _
v_status As Integer, vv_status As Integer, p_status As Integer, bkr_status As Integer, rma_status As Integer, _
VB_pots As Integer, _
VB_lamps As Byte, elec_lamps As Byte, _
alarm As Byte, elec_alarm As Byte, _
active_actions As Integer, active_reset_actions As Integer, _
active_faults As Integer, _
I_sounds As Integer)

' 24/3/98
Global active_actions(32) As Integer

' 5/6/98
'Global active_reset_actions(32) As Integer
Global active_reset_actions(128) As Integer

Global active_faults(32) As Integer
' 7/4/98
Global I_sounds(8) As Integer
Global old_I_sounds(8) As Integer

Global ItoS As Byte

Declare Sub get_v_status Lib "c:\HAVEN\Model.DLL" (v_status As Integer)
Global v_status(200) As Integer

Declare Sub get_vv_status Lib "c:\HAVEN\Model.DLL" (vv_status As Integer)
Global vv_status(30) As Integer

Declare Sub get_p_status Lib "c:\HAVEN\Model.DLL" (p_status As Integer)
Global p_status(55) As Integer

Declare Sub get_bkr_status Lib "c:\HAVEN\Model.DLL" (bkr_status As Integer)
Global bkr_status(66) As Integer

Declare Sub get_rma_status Lib "c:\HAVEN\Model.DLL" (rma_status As Integer)
Global rma_status(8) As Integer

Declare Sub get_pots Lib "c:\HAVEN\Model.DLL" (VB_pots As Integer)
Global VB_pots(0 To 255) As Integer

Declare Sub get_lamps Lib "c:\HAVEN\Model.DLL" (VB_lamps As Byte)
Global VB_lamps(500) As Byte

' 3/3/98
Declare Sub get_elec_lamps Lib "c:\HAVEN\Model.DLL" (elec_lamps As Byte)
Global elec_lamps(100) As Byte


Declare Sub get_alarms Lib "c:\HAVEN\Model.DLL" (alarm As Byte)
' 5/3/98
'Global alarm(1 To 350) As Byte
'Global alrm(1 To 350) As Integer
Global alarm(1 To 450) As Byte
Global alrm(1 To 450) As Integer

Declare Sub get_elec_alarms Lib "c:\HAVEN\Model.DLL" (elec_alarm As Byte)
Global elec_alarm(1 To 100) As Byte




Global itest_type As Integer
Global itest_no As Integer, ichannel As Integer, ivalue As Integer



Global VB_Droop As Single



Global record_on As Byte, replay_on As Byte


' 24/4/97
'Declare Sub read_SDI Lib "c:\sultan\Version1\debug\Version1.dll" (VB_SDIchanges As Integer)
Global VB_SDIchanges(0 To 20) As Integer
'Declare Sub Rset_SDI Lib "c:\sultan\Version1\debug\Version1.dll" (Snch As Integer)
Global Snch As Integer

' 21/4/97
'Declare Sub readDIch Lib "c:\sultan\Version1\debug\Version1.dll" (VB_DIchanges As Integer)
Global VB_DIchanges(0 To 20) As Integer
Global nch As Integer

' 10/4/97
Global AIOLDVALUE(0 To 27) As Integer
Global AIACTUAL(0 To 255) As Integer

Global par_alts(4) As Integer
Global SPLITMODE As Byte

' 12/3/97
Global VB_FAULT(1000) As Byte
' 13/2/97
Global Ifault As Integer, VDU_TYPE As Integer, Ifreq As Integer, Ivolts As Integer
Global xx As Integer

' 7/2/97
' 19/3/98
'Global PC_REALS(200) As Single
Global PC_REALS(2000) As Single


' 10/2/97
Global PC_INTS(200) As Integer

' 3/3/97
Global T42SCOPER(100) As Single
Global EDCR(100) As Single
Global T42AVRR(150) As Single
Global T42BUSR(200) As Single
Global EXECMODELR(100) As Single
Global T42DIESELR(200) As Single
Global T42BKRL(100) As Byte
Global T42BUSL(100) As Byte
Global VB_IO_TIME As Single

'IO added 28/2/97
'Global AI(0 To 255) As Byte
Global AI(0 To 255) As Integer
'Global AO(0 To 767) As Byte
Global AO(0 To 767) As Integer
Global AOold(0 To 767) As Byte ' need an AOold for sound o/p (0 to 255)
Global DI(0 To 3071) As Byte
Global DIold(0 To 3071) As Byte
Global DIchanges(0 To 128) As Byte ' used for logging
'Global VBB_DO(0 To 5119) As Byte 'can't use "DO" as an array name

''''' 10/12/97 Global VB_DO(0 To 5119) As Byte 'can't use "DO" as an array name

Global DelayLoopCount As Double 'used to determine fixed delay during i/o (dependant on processor speed)

Global Const AObase As Byte = &H12
Global Const SNDbase As Byte = &H9
Global Const AIbase As Byte = &H31
Global Const DIbase As Byte = &H52
Global Const DObase As Byte = &H82


Global Const AOtop As Byte = &H18
Global Const SNDtop As Byte = &HD
Global Const AItop As Byte = &H33
Global Const DItop As Byte = &H5B
Global Const DOtop As Byte = &H87

Global AOBoardAddr As Byte
Global SNDBoardAddr As Byte
Global AIBoardAddr As Byte
Global DOBoardAddr As Byte

Global BoardAddress As Byte
Global MsBoardAddr As Byte
Global LsBoardAddr As Byte
Global AOChip As Byte
Global AIChip As Byte
Global DOChip As Byte
Global diChip As Byte
Global ChipAddr As Byte
Global Offset As Integer

'Misc
'20/3/97 Global BitValue As Byte
Global BitValue As Integer
Global SlowFlash As Integer

Global FlashCountActual As Integer 'counter for flash rate of digital o/p
Global FlashCountMax As Integer 'counter for flash rate of digital o/p
Global IOCounter As Integer
Global GPIOTestOn As Integer
Global HWTestOn As Byte

Global test_flag As Byte
Global ICODE As Integer
Global Instructor_name As String * 20
Global Class_name As String * 20
Global Course_name As String * 20


Global EXECL(300) As Byte
Global EXECMODELL(100) As Byte


Global AI_chan_no As Integer, AI_Value As Integer


'''''Global task_array(100) As Byte



'Global Rx_strings(5) As String
Global Rx_strings(16) As String

' ================================================================


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Global Search_String As String
Global String1 As String, String2 As String, string3 As String

' Operating Modes
Global Const MODE_SINGLE = 1
Global Const MODE_NORMAL = 2

' Operation Mode - either MODE_SINGLE or MODE_NORMAL
Global Op_Mode As Integer

Dim LastSocket As Integer

' An array of 2500 singles can be reliably transmitted on one machine in a loop using 32k buffers
' Above 2500 errors begin to appear
' When going through one machine, data speed is aprox 1000 items per second

' When using two machines, transmitting the data takes less than half a second

' The Ports are used for the communiction sockets - 2000..2005 are known to be useable.

' 13/3/98 Multiple Definition
'Global Const Port_0 = 2000
'Global Const Port_1 = 2001
Global counter As Integer


' 27/1/97  TESTING ONLY:
Global J2_flag As Integer, X1_flag As Integer, M1_off As Integer

'30/1/97
'Global gen_ON(6) As Integer, Main_BKR_closed(6) As Integer, Cross_BKR_closed(8) As Integer, EDC_BKR_closed(6) As Integer
'Global Xlive(8) As Integer
'Global In_a_group As Integer
' 3/2/97
Global GEN_on(6) As Byte, Main_BKR_closed(6) As Byte, Cross_BKR_closed(8) As Byte, EDC_BKR_closed(6) As Byte
Global Xlive(8) As Byte
Global old_gen_ON(6) As Byte, old_Main_BKR_closed(6) As Byte, old_Cross_BKR_closed(8) As Byte, old_EDC_BKR_closed(6) As Byte
Global old_Xlive(8) As Byte

' Status of simulator - Running, Stopped, Reset
Global Sim_Running_Status As String




' These definitions are used by the various
' graphics modules - the values come from the
' Fortran dlls.

Global eg(3) As Single
Global gen_volt(3) As Single
Global cosphi(3) As Single
Global ra(3) As Single
Global tandelta(3) As Single
Global p(3) As Single
Global q(3) As Single
Global cos_val(3) As Integer

Global bg_switch_closed(3) As Boolean
Global busbar_switch_closed As Boolean
' RWJ code to interface with Fortran model
'
' NOTE - The directory on the machine being accessed,
' e.g. 'sultan' must be Shared on that machine
' (open Explorer, right click on the directory to
'  share, select 'Sharing...' and then 'Shared As')






Sub fill_PCINTS()

' 10/2/97   Fill the array PC_INTS for Transmission via TCPIP
Dim i As Integer


    PC_INTS(0) = 0
    
    
End Sub


Sub FILL_PCREALS()

' 7/2/97   Fill the array PC_REALS for Transmission via TCPIP
Dim i As Integer


    PC_REALS(0) = 0#
    
' 17/3/98
' 19/3/98
'For i = 0 To 200
For i = 0 To 2000
    PC_REALS(i) = i + 2000.123
Next i



  
End Sub


'Sub strip_string(Search_String, string1, string2, string3)
Sub strip_Rx_strings()

Dim SearchChar As String * 1
'Dim Pos1 As Integer, Pos2 As Integer, Pos3 As Integer, Pos4 As Integer
Dim StartPosition As Integer, EndPosition As Integer, i As Integer

'Search_String = "12.3#Richard#4567"   ' String to search in.
SearchChar = "#"    ' Search for "#".

' A textual comparison starting at position 1
'Pos1 = 1
'Pos2 = InStr(1, Search_String, SearchChar, 1)
'Pos3 = InStr(Pos2 + 1, Search_String, SearchChar, 1)
'Pos4 = InStr(Pos3 + 1, Search_String, SearchChar, 1)
 
'string1 = Mid$(Search_String, 1, Pos2 - 1)
'string2 = Mid$(Search_String, Pos2 + 1, Pos3 - Pos2 - 1)
'string3 = Mid$(Search_String, Pos3 + 1, Len(Search_String) - Pos3)


  StartPosition = 1
' 26/3/98
'  For i = 1 To 5
  For i = 1 To 9
     EndPosition = InStr(StartPosition + 1, Search_String, "#", 1)
     Debug.Print "EndPostion = " & EndPosition
     If ((EndPosition > StartPosition - 1) And (EndPosition > 0)) Then
' 2/4/97
 '       Rx_strings(i) = Mid(Search_String, StartPosition + 1, EndPosition - StartPosition - 1)
         Rx_strings(i) = Trim(Mid(Search_String, StartPosition + 1, EndPosition - StartPosition - 1))
    Else
        Rx_strings(i) = "x"
     End If
     StartPosition = EndPosition
  Next i
  
 
End Sub


Function Op_Mode_Name() As String
Dim txt As String

Select Case Op_Mode
    Case MODE_SINGLE
    txt = "Split"
    Case MODE_NORMAL
    txt = "Normal"
End Select

txt = txt + " Mode"
Op_Mode_Name = txt

End Function

Sub update_pump(frm As Form, index_no As Integer, pump_no As Integer)

' 18/12/97
If p_status(pump_no) = 0 Then    '  Pump is OFF ( and  not in Stand-bye)
        frm.Havenpump1(index_no).PumpOn = False
        
'        frm.Pump_OFF(index_no).ForeColor = &HFFFFFF         ' WHITE
'        frm.Pump_ON(index_no).ForeColor = &H808080           ' GREY
'        frm.Pump_SBY(index_no).ForeColor = &H808080         ' GREY
        
 ' 24/2/98
'       frm.OFFpump(index_no).BackColor = &HFF&               ' BRIGHT RED

        frm.OFFpump(index_no).BackColor = &HFFFFFF             ' WHITE
'        frm.ONpump(index_no).BackColor = &H5000&             ' DARK GREEN
'        frm.sbyPUMP(index_no).BackColor = &H8080&             '  DARK YELLOW
        frm.ONpump(index_no).BackColor = &HC0C0C0           ' GREY
        frm.sbyPUMP(index_no).BackColor = &HC0C0C0           ' GREY
       
        
        
ElseIf p_status(pump_no) = 1 Then    '  Pump is OFF ( and  in Stand-bye)
       frm.Havenpump1(index_no).PumpOn = False
       
'        frm.Pump_OFF(index_no).ForeColor = &HFFFFFF          ' WHITE
'        frm.Pump_ON(index_no).ForeColor = &H808080            ' GREY
'        frm.Pump_SBY(index_no).ForeColor = &HFFFF&           ' YELLOW
       
  ' 24/2/98
'       frm.OFFpump(index_no).BackColor = &HFF&               ' BRIGHT RED

        frm.OFFpump(index_no).BackColor = &HFFFFFF             ' WHITE
 '       frm.ONpump(index_no).BackColor = &H5000&             ' DARK GREEN
         frm.ONpump(index_no).BackColor = &HC0C0C0           ' GREY
       frm.sbyPUMP(index_no).BackColor = &HFFFF&                    ' BRIGHT YELLOW
        
        
ElseIf p_status(pump_no) = 2 Then   '  Pump is ON ( and  in Stand-bye)
       frm.Havenpump1(index_no).PumpOn = True
       
'        frm.Pump_OFF(index_no).ForeColor = &H808080           ' GREY
'        frm.Pump_ON(index_no).ForeColor = &HC000&       ' DULL GREEN
'        frm.Pump_SBY(index_no).ForeColor = &HFFFF          ' YELLOW
        
' 24/2/98
 '       frm.OFFpump(index_no).BackColor = &HC0&                 ' DARK RED
 
'         frm.OFFpump(index_no).BackColor = &H606060           ' DARK GREY
          frm.OFFpump(index_no).BackColor = &HC0C0C0           ' DARK GREY
        frm.ONpump(index_no).BackColor = &HFF00&               ' BRIGHT GREEN
         frm.sbyPUMP(index_no).BackColor = &HFFFF&                 ' BRIGHT YELLOW
        
ElseIf p_status(pump_no) = 3 Then    '  Pump is ON ( and  not in Stand-bye)
      frm.Havenpump1(index_no).PumpOn = True
      
'       frm.Pump_OFF(index_no).ForeColor = &H808080           ' GREY
'        frm.Pump_ON(index_no).ForeColor = &HC000&             ' DULL GREEN
'        frm.Pump_SBY(index_no).ForeColor = &H808080         ' GREY
        
' 24/2/98
'        frm.OFFpump(index_no).BackColor = &HC0&               ' DARK RED

'        frm.OFFpump(index_no).BackColor = &H606060           ' DARK GREY
          frm.OFFpump(index_no).BackColor = &HC0C0C0           ' DARK GREY
          
         frm.ONpump(index_no).BackColor = &HFF00&              ' BRIGHT GREEN
'         frm.sbyPUMP(index_no).BackColor = &H8080&             '  DARK YELLOW
          frm.sbyPUMP(index_no).BackColor = &HC0C0C0           ' DARK GREY
        
End If

End Sub
Sub update_ONOFF_pump(frm As Form, index_no As Integer, pump_no As Integer)

' 3/2/98
' These pumps have ONLY ON/OFF . There is NO Stand-by  (  e.g.  DO Transfer pump )
' They also have a HavenPump1
If p_status(pump_no) = 0 Then    '  Pump is OFF ( and  not in Stand-bye)
        frm.Havenpump1(index_no).PumpOn = False
        frm.Pump_ON(index_no).ForeColor = &H808080           ' GREY
ElseIf p_status(pump_no) = 1 Then    '  Pump is OFF ( and  in Stand-bye)
       frm.Havenpump1(index_no).PumpOn = False
        frm.Pump_ON(index_no).ForeColor = &H808080            ' GREY
ElseIf p_status(pump_no) = 2 Then   '  Pump is ON ( and  in Stand-bye)
       frm.Havenpump1(index_no).PumpOn = True
        frm.Pump_ON(index_no).ForeColor = &HC000&       ' DULL GREEN
ElseIf p_status(pump_no) = 3 Then    '  Pump is ON ( and  not in Stand-bye)
      frm.Havenpump1(index_no).PumpOn = True
        frm.Pump_ON(index_no).ForeColor = &HC000&       ' DULL GREEN
End If

End Sub
Sub update_heater(frm As Form, index_no As Integer, pump_no As Integer)

' 3/2/98
' 24/2/98 up-date
' These have ONLY ON/OFF . There is NO Stand-by  and NO Pump symbol
If p_status(pump_no) = 0 Then    '  Pump is OFF ( and  not in Stand-bye)
 '       frm.Pump_ON(index_no).ForeColor = &H808080           ' GREY
         frm.ONpump(index_no).BackColor = &HC0C0C0           ' GREY
ElseIf p_status(pump_no) = 1 Then    '  Pump is OFF ( and  in Stand-bye)
'        frm.Pump_ON(index_no).ForeColor = &H808080            ' GREY
         frm.ONpump(index_no).BackColor = &HC0C0C0           ' GREY
ElseIf p_status(pump_no) = 2 Then   '  Pump is ON ( and  in Stand-bye)
'        frm.Pump_ON(index_no).ForeColor = &HC000&       ' DULL GREEN
         frm.ONpump(index_no).BackColor = &HFF00&               ' BRIGHT GREEN
ElseIf p_status(pump_no) = 3 Then    '  Pump is ON ( and  not in Stand-bye)
'        frm.Pump_ON(index_no).ForeColor = &HC000&       ' DULL GREEN
         frm.ONpump(index_no).BackColor = &HFF00&               ' BRIGHT GREEN
End If

End Sub

Sub update_valve(Formname As Form, valve_number As Integer, i As Integer)

'  Valves with 1 DI and 1 DO as found on "Starter Boxes"

' 12/11/97 edited (removed ColOutlineTop & bottom )
If v_status(valve_number) = 1 Then

         Formname.Valve(i).OpenTopLeft = Not Formname.Valve(i).OpenTopLeft
         Formname.Valve(i).OpenBtmRight = Not Formname.Valve(i).OpenBtmRight
         
                  Formname.Valve(i).ColourOpen = &HFF00&        ' GREEN
                  
                 ' 15/12/98
                 Formname.Valve(i).ColourClosed = &HFFFFFF     ' WHITE
                 
ElseIf v_status(valve_number) = 2 Then

         Formname.Valve(i).OpenTopLeft = Not Formname.Valve(i).OpenTopLeft
         Formname.Valve(i).OpenBtmRight = Not Formname.Valve(i).OpenBtmRight
                 Formname.Valve(i).ColourOpen = &HFF&          ' RED

                 ' 15/12/98
                 Formname.Valve(i).ColourClosed = &HFFFFFF     ' WHITE
                 
ElseIf v_status(valve_number) = 0 Then

         Formname.Valve(i).OpenTopLeft = False
         Formname.Valve(i).OpenBtmRight = False
         
                 Formname.Valve(i).ColourOpen = &HFF00&        ' GREEN
                 
                 ' 15/12/98
                 Formname.Valve(i).ColourClosed = &HFF&        ' RED

Else
          Formname.Valve(i).OpenTopLeft = True
         Formname.Valve(i).OpenBtmRight = True
         
                 Formname.Valve(i).ColourOpen = &HFF00&        ' GREEN
                 
                 ' 15/12/98
                 Formname.Valve(i).ColourClosed = &HFFFFFF     ' WHITE
End If

  
End Sub
Sub copy_valve(Formname As Form, Vindex1 As Integer, Vindex2 As Integer)

         Formname.Valve(Vindex2).OpenTopLeft = Formname.Valve(Vindex1).OpenTopLeft
         Formname.Valve(Vindex2).OpenBtmRight = Formname.Valve(Vindex1).OpenBtmRight
         
                  Formname.Valve(Vindex2).ColourOpen = Formname.Valve(Vindex1).ColourOpen
                  
                 ' 15/12/98
                  Formname.Valve(Vindex2).ColourClosed = Formname.Valve(Vindex1).ColourClosed
                                    
End Sub
Sub update_rma(nRMA As Integer, set_point As Single, rma_text As String)

' 19 / 12 / 97  NOTE:
'      rma_status( ) = 0   LOCAL  MANUAL
'      rma_status( ) = 1   LOCAL  AUTO
'      rma_status( ) = 2   REMOTE MANUAL
'      rma_status( ) = 3   REMOTE AUTO
'      rma_status( ) = 4   REAL PID selected

If rma_status(nRMA) = 0 Then     '   (1) for MELUB (2) for Scavenge air etc.
      rma_text = "LCL MANUAL"
      set_point = PCCONTROL(nRMA + 40)
ElseIf rma_status(nRMA) = 1 Then
      rma_text = "LCL AUTO"
      set_point = PCCONTROL(nRMA + 60)
ElseIf rma_status(nRMA) = 2 Then
      rma_text = "REM MANUAL"
      set_point = PCCONTROL(nRMA)
ElseIf rma_status(nRMA) = 3 Then
      rma_text = "REM AUTO"
     set_point = PCCONTROL(nRMA + 20)
ElseIf rma_status(nRMA) = 4 Then
      rma_text = "REAL PID"
      set_point = PCCONTROL(nRMA + 20)
Else
     rma_text = "undefined"
      set_point = 0#
End If

End Sub
Sub update_vvalve(Formname As Form, valve_number As Integer, indx1 As Integer, indx2 As Integer)

'  Valves with 2 DI's and 2 DO's   ( If the second valve (indx2) does not exist then set indx2 = -1 )

If vv_status(valve_number) = 1 Then

         Formname.Valve(indx1).OpenTopLeft = Not Formname.Valve(indx1).OpenTopLeft
         Formname.Valve(indx1).OpenBtmRight = Not Formname.Valve(indx1).OpenBtmRight
                  Formname.Valve(indx1).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                 Formname.Valve(indx1).ColourClosed = &HFFFFFF      ' WHITE
                  
        If indx2 > -1 Then
            Formname.Valve(indx2).OpenTopLeft = Not Formname.Valve(indx2).OpenTopLeft
            Formname.Valve(indx2).OpenBtmRight = Not Formname.Valve(indx2).OpenBtmRight
                     Formname.Valve(indx2).ColourOpen = &HFF&          ' RED
                  ' 15/12/98
                 Formname.Valve(indx2).ColourClosed = &HFFFFFF      ' WHITE
        End If
        
ElseIf vv_status(valve_number) = 2 Then

         Formname.Valve(indx1).OpenTopLeft = Not Formname.Valve(indx1).OpenTopLeft
         Formname.Valve(indx1).OpenBtmRight = Not Formname.Valve(indx1).OpenBtmRight
                 Formname.Valve(indx1).ColourOpen = &HFF&          ' RED
                  ' 15/12/98
                 Formname.Valve(indx1).ColourClosed = &HFFFFFF      ' WHITE
         If indx2 > -1 Then
               Formname.Valve(indx2).OpenTopLeft = Not Formname.Valve(indx2).OpenTopLeft
                Formname.Valve(indx2).OpenBtmRight = Not Formname.Valve(indx2).OpenBtmRight
                     Formname.Valve(indx2).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                 Formname.Valve(indx2).ColourClosed = &HFFFFFF      ' WHITE
          End If
          
ElseIf vv_status(valve_number) = 0 Then

         Formname.Valve(indx1).OpenTopLeft = False
         Formname.Valve(indx1).OpenBtmRight = False
                 Formname.Valve(indx1).ColourOpen = &HFF00&        ' GREEN
                 
                 ' 15/12/98
                 Formname.Valve(indx1).ColourClosed = &HFF&          ' RED
                 
        If indx2 > -1 Then
             Formname.Valve(indx2).OpenTopLeft = True
             Formname.Valve(indx2).OpenBtmRight = True
                 Formname.Valve(indx2).ColourOpen = &HFF00&        ' GREEN
                 
                  ' 15/12/98
                 Formname.Valve(indx2).ColourClosed = &HFFFFFF      ' WHITE
        End If
        
Else

          Formname.Valve(indx1).OpenTopLeft = True
         Formname.Valve(indx1).OpenBtmRight = True
                 Formname.Valve(indx1).ColourOpen = &HFF00&        ' GREEN
                   ' 15/12/98
                 Formname.Valve(indx1).ColourClosed = &HFFFFFF      ' WHITE
                
         If indx2 > -1 Then
                 Formname.Valve(indx2).OpenTopLeft = False
                 Formname.Valve(indx2).OpenBtmRight = False
                     Formname.Valve(indx2).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                 Formname.Valve(indx2).ColourClosed = &HFF&          ' RED
        End If
                  
End If

End Sub
Sub catch_alarms()


'08/09/00
If ResetAllAlarms Then
    For i = 1 To 450
      old_alarm_latest(i) = -123
      alarm_latest(i) = 0
      alarm_logged(i) = False
      alarm_time(i) = "00 00 00"
    Next
    ResetAllAlarms = False
End If



' 5/3/98
For ial = 1 To 100
  alarm(ial + 350) = elec_alarm(ial)
Next ial



' 29/12/97 Transfer alarm( ) array from fortran to the alrm( ) array in VB
' NB  VB gets complex logical expressions WRONG if the BYTE array alarm( ) is used in a complex expression
'EG     If alarm(ial) And (Not alarm_logged(ial)) Then      '  THIS GIVES A TRUE RESULT
'          If (Not alarm(ial)) And alarm_logged(ial) Then      '  THIS ALSO GIVES A TRUE RESULT
'
'    However VB does get the following correct:

'For ial = 1 To 350
For ial = 1 To 450

  If ial = 69 Or ial = 70 Then GoTo skip_these   ' (Alarm(4) and Alarm(12) are copies , log these instead 0
    If alarm(ial) Then
       alrm(ial) = True
    Else
       alrm(ial) = False
    End If
skip_these:

Next ial


' 23/12/97 modified from Version 2.5 (Iceland)

' 5/3/98  Changed 350 to 450

' For ial = 1 To 350
 For ial = 1 To 450
   If alrm(ial) And (Not alarm_logged(ial)) Then
'    For jal = 351 To 2 STEP -1
     For jal = 451 To 2 Step -1
        alarm_latest(jal) = alarm_latest(jal - 1)
          alarm_time(jal) = alarm_time(jal - 1)
     Next jal
     
        alarm_latest(1) = ial
          alarm_time(1) = current_time
       alarm_logged(ial) = True
       
       ' 23/7/98 ADDED:
         flash_latest_alarm = True
         
         
'        Debug.Print "Log " & alarm_latest(1) & "  " & alarm_time(1)
'    If ial = 1 Then
'       Debug.Print "LOG   Alrm(1) = " & alrm(1)
'    End If
 
  End If
 Next ial

' For ial = 1 To 350
 For ial = 1 To 450
  If (Not alrm(ial)) And alarm_logged(ial) Then
    

' 5/3/98  For jal = 1 To 351
          For jal = 1 To 451
             If alarm_latest(jal) = ial Then
             
'                  Debug.Print "UNLog " & ial & "  " & alarm_time(jal)
                  
' 5/3/98          For kal = jal To 350
                  For kal = jal To 450
                    alarm_latest(kal) = alarm_latest(kal + 1)
                      alarm_time(kal) = alarm_time(kal + 1)
                  Next kal
                 
              End If
           Next jal
           
 ' 5/3/98    alarm_latest(351) = 0
             alarm_latest(451) = 0
             alarm_logged(ial) = False
             
   End If
 Next ial
 
End Sub
Sub inv_copy_valve(Formname As Form, valve_number As Integer, Vindex1 As Integer, Vindex2 As Integer)

         Formname.Valve(Vindex2).OpenTopLeft = Not Formname.Valve(Vindex1).OpenTopLeft
         Formname.Valve(Vindex2).OpenBtmRight = Not Formname.Valve(Vindex1).OpenBtmRight
         
           If v_status(valve_number) = 1 Then
                  Formname.Valve(Vindex2).ColourOpen = &HFF&          ' RED
           Else
                  Formname.Valve(Vindex2).ColourOpen = &HFF00&       ' GREEN
           End If
           
           
                ' 15/12/98 ADDED:
            If v_status(valve_number) = 3 Then
                 Formname.Valve(Vindex2).ColourClosed = &HFF&       ' RED
            Else
                Formname.Valve(Vindex2).ColourClosed = &HFFFFFF     ' WHITE
            End If
           
End Sub
Sub Control_PC_check()


' 19/3/98
DoEvents

' 14/1/98
If PC_IN_CONTROL = This_PC_number Then
             This_Control_status = True
Else
            This_Control_status = False
        If PC_IN_CONTROL = -1 Then
           response = MsgBox("There is NO PC in CONTROL", vbOK, "Information")
        Else
           response = MsgBox("This PC is NOT the CONTROL PC. The CONTOL PC is PC " & Format(PC_IN_CONTROL, "#00"), vbOK, "Information")
        End If
End If

End Sub
Sub toggle_PC_control(frm As Form, indx As Integer)

' 16/1/98
If indx = 0 Then
     '  We are trying to set this PC to be the ONE and ONLY PC_in_control
       If PC_IN_CONTROL = -1 Then
               frm.PC_icon(1).Visible = True
               frm.PC_icon(0).Visible = False
                 PC_IN_CONTROL = This_PC_number
      Else
            response = MsgBox("PC " & Format(PC_IN_CONTROL, "#00") & " has control" & Chr(10) & "This PC can not be the control PC", vbOK, "Information")
      End If

ElseIf indx = 1 Then
      If PC_IN_CONTROL = This_PC_number Then         ' It is OK to de-select this PC
               frm.PC_icon(1).Visible = False
               frm.PC_icon(0).Visible = True
                 PC_IN_CONTROL = -1
       End If
End If

End Sub
Sub update_PC_ICON(frm As Form)

'  IF this PC is in CONTROL then set the PC ICON to show this:
If PC_IN_CONTROL = This_PC_number Then
               frm.PC_icon(1).Visible = True
               frm.PC_icon(0).Visible = False
Else
               frm.PC_icon(1).Visible = False
               frm.PC_icon(0).Visible = True
End If
frm.This_PC_box.Caption = "PC " & Format(This_PC_number, "#00")

End Sub
Sub setPCDI(nPCDI As Integer)
Dim DITextStr As String
Dim Fnum3 As Integer

If I_AM_HOST Or Not In_a_group Then

  ' 22/5/98
   If medium_speed Then
   
        Call set_PCDI(nPCDI)                        '  SEND  to the "local"  MODEL.DLL
  
   Else
   
          ' 6/1/00
        '   Call set_PCDI_LSS(nPCDI)                    '  SEND  to the "local"  LSS_SIM.DLL
        
        If (PTT Or I_AM_INSTRUCTOR) Then
        
            Call set_PCDI_LSS(nPCDI)                    '  SEND  to the "local"  LSS_SIM.DLL
        
        Else
        
        ' === 13/03/01 ===
        '  NOTE: Ther are VIRTUAL PUMPS between DI 3000 and 3199 so...
        ' If (nPCDI > 3199) Then
          If (nPCDI > 2999) Then
            Call set_PCDI_LSS(nPCDI)                    '  SEND  to the "local"  LSS_SIM.DLL
          End If
          
        End If
 
    

   End If
    
    

   
   
   
Else

        MessageString = "#SET_PCDI#" & Format(nPCDI, "00000") & "#"

    ' 3/4/00
    If Not I_AM_INSTRUCTOR And Not PTT Then
    
          ' I MUST be an ACT Terminal (Full Mission) so
          
        ' === 13/03/01 ===
        '  NOTE: Ther are VIRTUAL PUMPS between DI 3000 and 3199 so...
        ' If (nPCDI > 3199) Then
          If (nPCDI > 2999) Then
          
            ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"
            Call frm_Host.SendMessage(MessageString)
            
          End If
          
    Else

       ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"
        Call frm_Host.SendMessage(MessageString)
    
    End If
    
    
    'Debug.Print "Sending PCDI Message = " & MessageString

End If

'5/6/98 Add Logging for PCDIs


'4/4/00If LogFileOpen Then

      ' === 13/03/01 Ther are VIRTUAL pumps between DI 3000 and 3199
      ' ( See Breaker Ref V15.xls
      ' If nPCDI <= 3071 Then
        If nPCDI < 3000 Then
        
            'If Mid(DITEXT(nPCDI), 1, 10) <> "          " Then
                'DITextStr = DITEXT(nPCDI)
                
             '25/10/99 for Steam Sim
            '29/10/99 db open anyway If Not PTT Then
            
                'jf no database from 11/11/99
                'DIDataBase.MoveFirst
                'DIDataBase.Find "[i/o ref] =" & nPCDI
                'If Not (DIDataBase.EOF) Then
                    'DITextStr = DIDataBase(1)
                'Else
                    'DITextStr = "Unknown Switch"
                'End If

                Call GetHardIOText("DI", nPCDI, DITextStr)
                Call Logging(DITextStr, LogPCDIs)
            'Else
                'Call Logging("Can't log #" & nPCDI & " yet - no text", LogPCDIs)
            'End If
            'Else
                'Call Logging("Digital Input # " & str(nPCDI), LogPCDIs)
            'End If
        Else
        
            ' === 13/03/01 Ther are VIRTUAL pumps between DI 3000 and 3199
            ' ( See Breaker Ref V15.xls
            If nPCDI < 3200 Then
            
                  DITextStr = VIRTUAL_PUMP_TEXT(nPCDI - 3000)
                  
            Else
            
                'Call Logging("Digital Input #" & nPCDI & " is out of range", LogPCDIs)
                Fnum3 = FreeFile(0)
                Open "C:\steam_sim\text\VirtualvalveText.dat" For Random Access Read As #Fnum3 Len = Rec_Length
            
                Get #Fnum3, nPCDI, DITextStr
                Close #Fnum3
                
                ' === 07/06/02 ===
                If nPCDI = 4402 Then
                    DITextStr = "Bridge Request (Bridge)"
                ElseIf nPCDI = 4401 Then
                    DITextStr = "Transfer Request (Bridge)"
                ElseIf nPCDI = 4403 Then
                    DITextStr = "ER Request (Bridge)"
                End If
                
                
            End If
            
            
                Call Logging(DITextStr, LogPCDIs)
            
        
        End If
'4/4/00End If



End Sub

Sub setSTEAMlogicals(nlog As Integer, status_flag As Byte)

' === 22/11/02 ===
If I_AM_HOST Or Not In_a_group Then

    Call set_STEAM_logicals(nlog, status_flag)
  
Else

   ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"
    If status_flag Then
       MessageString = "#setSTEAMlogicals#" & Format(nlog, "00000") & "#" & "TRUE" & "#"
    Else
       MessageString = "#setSTEAMlogicals#" & Format(nlog, "00000") & "#" & "FALSE" & "#"
    End If
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending PCDI Message = " & MessageString

End If


End Sub
Sub setDI(nDI As Integer, status_flag As Byte)

' 25/3/98
'  Call setDI(nDI, True)
If I_AM_HOST Or Not In_a_group Then

  If medium_speed Then
    Call set_DI(nDI, status_flag)                       '  SEND  to the "local"  MODEL.DLL
  Else
    Call set_DI_LSS(nDI, status_flag)                   '  SEND  to the "local"  LSS_SIM.DLL
  End If
  
Else
   ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"   MODEL.DLL
    If status_flag Then
       MessageString = "#setDI#" & Format(nDI, "00000") & "#" & "TRUE" & "#"
    Else
       MessageString = "#setDI#" & Format(nDI, "00000") & "#" & "FALSE" & "#"
    End If
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending PCDI Message = " & MessageString

End If

End Sub
Sub setPCcontrol(nCHANNEL As Integer, Value As Single)

' 26/3/98
If I_AM_HOST Or Not In_a_group Then

  If medium_speed Then
    Call set_PC_CONTROL(nCHANNEL, Value)         '  SEND  to the "local"  MODEL.DLL
  Else
    Call set_PC_CONTROL_LSS(nCHANNEL, Value)     '  SEND  to the "local"  SIM_LSS.DLL
  End If
  
Else
   ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"   MODEL.DLL

    MessageString = "#setPCcontrol#" & Format(nCHANNEL, "00000") & "#" & Format(Value, "#####0.000") & "#"
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending PC CONTROL Message = " & MessageString

End If


End Sub
Sub setAI(nCHANNEL As Integer, nValue As Integer)

' 25/3/98
'Call setAI(nCHANNEL, nVALUE)

If I_AM_HOST Or Not In_a_group Then

  If medium_speed Then
    Call set_AI(nCHANNEL, nValue)        '  SEND  to the "local"  MODEL.DLL
  Else
    Call set_AI_LSS(nCHANNEL, nValue)    '  SEND  to the "local"  LSS_SIM.DLL
  End If
  
Else
   ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"   MODEL.DLL

    MessageString = "#setAI#" & Format(nCHANNEL, "00000") & "#" & Format(nValue, "00000") & "#"
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending setAI Message = " & MessageString

End If


End Sub

Sub setVscroll(nCHANNEL As Integer, nValue As Integer)

' === 7/12/00 NEW Sub ===

If I_AM_HOST Or Not In_a_group Then

    Call set_VSCROLL(nCHANNEL, nValue)    '  SEND  to the "local"  STEAM_SIM.DLL
  
Else
   ' Send to the HOST PC via a SOCKET     '  SEND to the "Host"   STEAM_SIM.DL

    MessageString = "#setVSCROLL#" & Format(nCHANNEL, "00000") & "#" & Format(nValue, "00000") & "#"
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending setVSCROLL Message = " & MessageString

End If


End Sub

Sub setHILO(nalarm As Integer, HILO_Value As Single)

' 25/3/98
If I_AM_HOST Or Not In_a_group Then

    If medium_speed Then
     Call set_HILO(nalarm, HILO_Value)           '  SEND  to the "local"  MODEL.DLL
    Else
     Call set_HILO_LSS(nalarm, HILO_Value)       '  SEND  to the "local"  SIM_LSS.DLL
    End If
Else
   ' Send the PCDI to the HOST PC via a SOCKET   '  SEND to the "Host"   MODEL.DLL

    MessageString = "#SET_HILO#" & Format(nalarm, "00000") & "#" & Format(HILO_Value, "#####0.000") & "#"
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending HILO Message = " & MessageString

End If

End Sub
Sub setCONTROLLERvars(nCHANNEL As Integer, nValue As Integer)

' === 11/11/02 ===
If I_AM_HOST Or Not In_a_group Then

     Call set_CONTROLLER_VARS(nCHANNEL, nValue)

Else
   ' Send to the HOST PC via a SOCKET   '  SEND to the "Host"

    MessageString = "#SET_CV#" & Format(nCHANNEL, "00000") & "#" & Format(nValue, "00000") & "#"
    Call frm_Host.SendMessage(MessageString)

End If


End Sub
Sub setACTION(nset As Integer, naction As Integer)

' 24/3/98
If I_AM_HOST Or Not In_a_group Then


     If medium_speed Then
         Call set_ACTION(nset, naction)                '  SEND  to the "local"  MODEL.DLL
     Else
         Call set_ACTION_LSS(nset, naction)            '  SEND  to the "local"  LSS_SIM.DLL
     End If
     
Else
   ' Send the PCDI to the HOST PC via a SOCKET       '  SEND to the "Host"   MODEL.DLL
   
    MessageString = "#SET_ACTION#" & Format(nset, "00000") & "#" & Format(naction, "00000") & "#"
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending ACTION Message = " & MessageString

    
End If

End Sub
Sub setplotchannels()
Dim i As Integer

'global plot_channels(16) As Integer
If I_AM_HOST Or Not In_a_group Then

    'Call set_plot_channels(plot_channels(1))    '  SEND  to the "local"  MODEL.DLL
    Call set_plot_channels_LSS(plot_channels(1))    '  SEND  to the "local"  MODEL.DLL

Else
   ' Send the PCDI to the HOST PC via a SOCKET       '  SEND to the "Host"   MODEL.DLL
   
    MessageString = "#SET_PLOT_CHANNELS#"
    For i = 1 To 8
     MessageString = MessageString & str(plot_channels(i)) & "#"
    Next i
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending PLOT CHANNELS  Message = " & MessageString

    
End If

End Sub

Sub setFAULT(nset As Integer, nFault As Integer)

' 24/3/98
If I_AM_HOST Or Not In_a_group Then

   ' 5/6/98
   If medium_speed Then
     Call set_FAULT(nset, nFault)                   '  SEND  to the "local"  MODEL.DLL
   Else
     Call set_FAULT_LSS(nset, nFault)               '  SEND  to the "local"  LSS_SIM.DLL
   End If
   
Else
   ' Send the PCDI to the HOST PC via a SOCKET      '  SEND to the "Host"   MODEL.DLL
   
    MessageString = "#SET_FAULT#" & Format(nset, "00000") & "#" & Format(nFault, "00000") & "#"
    Call frm_Host.SendMessage(MessageString)
    Debug.Print "Sending ACTION Message = " & MessageString

    
End If

End Sub
Sub update_Hscroll1(frm As Form, nRMA As Integer)


' 27/5/98
If medium_speed Then

  If rma_status(nRMA) = 0 Then
     frm.HScroll1.Value = PCCONTROL(40 + nRMA) * 100              '  LOCAL MANUAL
  ElseIf rma_status(nRMA) = 1 Then
     frm.HScroll1.Value = PCCONTROL(60 + nRMA)                    '  LOCAL AUTO
  ElseIf rma_status(nRMA) = 2 Then
     frm.HScroll1.Value = PCCONTROL(nRMA) * 100                   '  REMOTE MANUAL
  ElseIf rma_status(nRMA) = 3 Or rma_status(nRMA) = 4 Then
     frm.HScroll1.Value = PCCONTROL(20 + nRMA)                    '  REMOTE AUTO or REAL PID
  End If
  
Else

 ' 9/7/98
 ' If rma_status_LSS(nRMA) = 0 Then
 '    frm.HScroll1.Value = PCCONTROL(40 + PC_control_LSS(nRMA)) * 100              '  LOCAL MANUAL
 ' ElseIf rma_status_LSS(nRMA) = 1 Then
 '    frm.HScroll1.Value = PCCONTROL(60 + PC_control_LSS(nRMA))                    '  LOCAL AUTO
 ' ElseIf rma_status_LSS(nRMA) = 2 Then
 '    frm.HScroll1.Value = PCCONTROL(PC_control_LSS(nRMA)) * 100                   '  REMOTE MANUAL
 ' ElseIf rma_status_LSS(nRMA) = 3 Or rma_status(nRMA) = 4 Then
 '    frm.HScroll1.Value = PCCONTROL(20 + PC_control_LSS(nRMA))                    '  REMOTE AUTO or REAL PID
 ' End If
  
  ' see sy030_lss form load which worked OK
  If rma_status_LSS(nRMA) = 0 Then
     frm.HScroll1.Value = LSS_reals(840 + PC_control_LSS(nRMA)) * 100       '  LOCAL MANUAL
  ElseIf rma_status_LSS(nRMA) = 1 Then
     frm.HScroll1.Value = LSS_reals(860 + PC_control_LSS(nRMA))             '  LOCAL AUTO
  ElseIf rma_status_LSS(nRMA) = 2 Then
     frm.HScroll1.Value = LSS_reals(800 + PC_control_LSS(nRMA)) * 100       '  REMOTE MANUAL
' 20/7/98
' ElseIf rma_status_LSS(nRMA) = 3 Or rma_status(nRMA) = 4 Then
  ElseIf rma_status_LSS(nRMA) = 3 Or rma_status_LSS(nRMA) = 4 Then
     frm.HScroll1.Value = LSS_reals(820 + PC_control_LSS(nRMA))             '  REMOTE AUTO or REAL PID
  End If
  


End If

End Sub
Sub update_water_valve(frm As Form, vv_state As Byte, indx As Integer)

If vv_state Then
   frm.Valve(indx).OpenBtmRight = True
   frm.Valve(indx).OpenTopLeft = True
Else
   frm.Valve(indx).OpenBtmRight = False
   frm.Valve(indx).OpenTopLeft = False
End If

End Sub
Sub sep_water_check()

' 30/1/98
        If ZFT(10) < 0.02 Then    ' ZLwater
           sep_water_low = True
           response = MsgBox("There is insufficient water in the Separator water tank", vbOK, "Information")
        Else
          sep_water_low = False
        End If
        
End Sub
Sub change_PC_control()

If frm_Host.StatusBar1.Panels.Item(5).text = "" Then

     '  We are trying to set this PC to be the ONE and ONLY PC_in_control
       If PC_IN_CONTROL = -1 Then

'               frm.PC_icon(1).Visible = True
'               frm.PC_icon(0).Visible = False
                 PC_IN_CONTROL = This_PC_number
                 frm_Host.StatusBar1.Panels.Item(5).text = "PC " & PC_IN_CONTROL & " has Control"
      Else
            response = MsgBox("PC " & Format(PC_IN_CONTROL, "#00") & " has control" & Chr(10) & "This PC can not be the control PC", vbOK, "Information")
      End If

Else
      If PC_IN_CONTROL = This_PC_number Then         ' It is OK to de-select this PC
'               frm.PC_icon(1).Visible = False
'               frm.PC_icon(0).Visible = True
                 PC_IN_CONTROL = -1
                 frm_Host.StatusBar1.Panels.Item(5).text = ""
       End If
End If

End Sub
Sub get_alarm_text()

' 14/12/98
Dim IN_string As String
Dim start_pos As Long
Dim end_pos As Long
Dim filenum As Integer
Dim resp As Integer

'Dim star_pos As Long
Dim alarm_number As Integer

'List1.Clear

'Open "c:\Haven\Text\E_AlarmText.txt" For Input As #1 ' Open file for input.
'12/12/00 getting error here sometimes. Need to specify error better  On Error GoTo Errhandler
On Error GoTo Errh1

If useEnglish Then
    filenum = FreeFile
    Open "c:\Steam_Sim\Text\E_SteamAlarmText.txt" For Input As #filenum ' Open file for input.

Else
    
    'Filepathname = App.Path & "\text\" & LanguagePrefix & "_" & ModuleName Use this version once App.Path is sensible
     
     Open "c:\Steam_Sim\Text\" & LanguagePrefix & "_AlarmText.txt" For Input As #filenum  ' Open file for input.

End If

On Error GoTo Errh2
Do While Not EOF(1) ' Loop until end of file.
    Input #1, IN_string
    'If Mid$(IN_string, 1, 10) = "ALARM_TEXT" Then
     If Val(IN_string) > 0 Then 'it must be valid text
   
         alarm_number = Val(IN_string)
         'If alarm_number = 405 Then
         '   alarm_number = 405
         'End If
         
         
        'alarmtxt(alarm_number) = Mid$(IN_string, start_pos + 1, end_pos - 1 - start_pos)
            If alarm_number >= LBound(alarmtxt) And alarm_number <= UBound(alarmtxt) Then
                start_pos = 1 'first character
                end_pos = Len(IN_string)
                '12/12/00 possibility that for blank lines of text, end >= start for error
                'alarmtxt(alarm_number) = Mid$(IN_string, start_pos, end_pos)
                'Debug.Print "Read Alarm # " & alarm_number & "  OK"
                If end_pos > start_pos Then
                    alarmtxt(alarm_number) = Mid$(IN_string, start_pos, end_pos)
                End If
            End If
    End If
Loop
Close #filenum    ' Close file.

Debug.Print " In get alarm text, text(1) = " & alarmtxt(1)

Exit Sub

'Errhandler:
Errh1:
resp = MsgBox("The file E_SteamAlarmText.txt is missing or corrupt", vbOKOnly, "File Error")
Exit Sub

Errh2:
resp = MsgBox("Error retrieving text from E_SteamAlarmText.txt, check text format is OK", vbOKOnly, "File Error")
Close #filenum    ' Close file.
Exit Sub

End Sub

