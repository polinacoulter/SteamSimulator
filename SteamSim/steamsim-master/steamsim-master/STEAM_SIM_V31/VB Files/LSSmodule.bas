Attribute VB_Name = "LSSmodule"


' === 07/06/02 ===
Global EOT_instructor(11) As Integer

' === 31/01/02 Screen Background Colour ===
Global Bcolour As Long
Global B_colour_default As Long

'
'  *** Large Slow Speed STEAM Simulator declarations ***
'
'

' === 28/01/02 ===
Global White_background As Boolean



' === 23/01/02 ===
Global Silence_BLEEP_TIME As Integer


' 9/8/00
Global Force_Instrp_update As Boolean

' 2/2/00
Global Fuel_Type As Integer

' 27/7/00
Global T_foot As Single

' 26/7/00
'      SUBROUTINE set_FUEL(VB_FUEL)
'      real*4    VB_FUEL(10)
Declare Sub set_FUEL Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (VB_fuel As Single)
Global VB_fuel(10) As Single


' 20/7/00
'      SUBROUTINE set_BEARING_severity(VB_Bearing_severity)
Declare Sub set_BEARING_severity Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (BEARING_severity As Single)
Declare Sub get_BEARING_severity Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (BEARING_severity As Single)
Global BEARING_severity(25) As Single


' === 14/7/00  ===
' FAULT_severity  ( note: severity i 0.0 to 100.0 % )
Declare Sub set_FAULT_severity Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nFault As Integer, severity As Single)
Declare Sub get_FAULT_severity Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nFault As Integer, severity As Single)

Global severity As Single
Global FAULT_severity_needed(1400) As Boolean

' 21/7/00
Global flash_bearing_severity As Boolean
Global Sticky_valve As Boolean


Global OLD_DO_BRIDGE As Byte
Global OLD_DO_ER As Byte



' 17/8/99
'Global OLE_flag As Boolean
Global flash_flag As Boolean

' 11/11/99
'Global old_bearing_value As Integer
Global old_bearing_value(25) As Integer

Global gov_position As Single
Global eng_rpm As Single
Global Pscav_bar As Single
'
Global gov_low As Single
Global gov_high As Single

' 23/7/98
Global flash_latest_alarm As Integer

' 21/7/98
Global update_VBAIs As Integer
Global update_VBAI(48) As Integer

' === 18/11/02 ===
Global update_VSCROLL(50) As Integer





' 13/10/99
Global DO_used(7000) As Integer
Global DI_used(5000) As Integer

' 9/7/98
Global slider_value(32) As Integer
'      slider_value(1) = SY030_LSS Vscroll1

' 9/7/98
'Global rma_text As String
'Global set_point As Single

' 10/6/98
Global flash_status As Integer    ' ( BOIL.FRM )


'25/10/99
Declare Sub get_DI_Reloads Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (VB_DI_Reload As Byte)
'Declare Sub get_AI_Reloads Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (VB_AI_Reload As Integer)
Declare Sub get_AI_Reloads Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (VB_AI_Reload As Byte)



' 16/7/98
Declare Sub set_EXECI_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, nValue As Integer)
'===========================================


' 2/7/99
' Path changed from c:\HAVEN to  c:\Steam_sim\Fortran Files\Debug....


' 5/9/99 ADDED:
Declare Sub get_steam_DO Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (D_OUTPUT As Byte)
Declare Sub get_steam_AO Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (A_OUTPUT As Byte)
Declare Sub set_steam_DI Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (D_INPUT As Byte)
Declare Sub set_steam_AI Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (A_INPUT As Byte)


' 7/7/99
Declare Sub set_STEAM_logicals Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nlog As Integer, status_flag As Byte)

' === 22/11/02 17.033 ===
Global nlog As Integer

' === 13/11/01 ===
Declare Sub set_TRIP_BREAKER Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nBreaker As Integer, status_flag As Byte)

' === 28/03/01 Point 10.29 ===
Declare Sub set_LSS_logicals Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nlog As Integer, status_flag As Byte)


Declare Sub set_VSCROLL Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, nValue As Integer)


Declare Sub AA_time_test Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" ()
Declare Sub open_com_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer)

Declare Sub get_Psat Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psat As Single, Tsat As Single)

Declare Sub get_Tsat Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psat As Single, Tsat As Single)
Declare Sub get_hf_hfg Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psat As Single, hf As Single, hfg As Single)

Declare Sub get_Tsteam Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psteam As Single, Tsat As Single, _
Hsteam As Single, hf As Single, hfg As Single, Tsteam As Single, qDryness As Single, CP_Steam As Single)


Declare Sub get_Hsteam Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psteam As Single, Tsat As Single, _
Hsteam As Single, hf As Single, hfg As Single, Tsteam As Single, qDryness As Single, CP_Steam As Single)

' 24/6/99
Declare Sub get_ssteam Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psteam As Single, Tsat As Single, _
ssteam As Single, hf As Single, hfg As Single, Tsteam As Single, qDryness As Single, CP_Steam As Single)
Declare Sub get_sf_sfg Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Psat As Single, sf As Single, sfg As Single)

'===========================================


' 5/6/98
Declare Sub set_FAULT_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nset As Integer, nFault As Integer)
Declare Sub set_ACTION_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nset As Integer, naction As Integer)
Declare Sub LSS_SNAP Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (SnapNum As Integer, SnapType As Integer)
Declare Sub LSS_RELSNAP Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (SnapNum As Integer, SnapType As Integer)

' 27/5/98
Declare Sub get_R_string_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal R_string As String)
Declare Sub set_R_string_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal R_string As String)
''' In main.bas:   Global R_string As String * 8004
Declare Sub get_I_string_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal I_string As String)
Declare Sub set_I_string_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal I_string As String)


' 17/2/00
Global flash_burner_1_1 As Boolean
Global flash_burner_1_2 As Boolean
Global flash_burner_2_1 As Boolean
Global flash_burner_2_2 As Boolean

' 27/5/98
Declare Sub get_alarm_text_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal alarm_no As Integer, ByVal alarm_text As String)
Declare Sub set_HILO_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nalarm As Integer, HILO_Value As Single)


'8/10/99
'set_CONTROLLER_PID_constants(nCHANNEL,R_VALUE,nLOOPNUMBER)
Declare Sub set_CONTROLLER_PID_constants Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, R_VALUE As Single, nLOOPNUMBER As Integer)
Global nLOOPNUMBER As Integer
Global n_Ti As Integer
Global n_Pband As Integer
Global Ti_value As Single
Global Pband_value As Single


' 8/6/98
Declare Sub set_DI_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nDI As Integer, status_flag As Byte)

Declare Sub set_PCDI_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nPCDI As Integer)
Declare Sub set_PC_CONTROL_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, Value As Single)
Declare Sub set_AI_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, nValue As Integer)

' === 04/03/02 set_CONSTANT(nCHANNEL,VALUE)===
Declare Sub set_CONSTANT Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, Value As Single)
Global Shaft_Inertia_multiplier As Single
Global Ship_Inertia_multiplier As Single
Global Pitch_multiplier As Single

' === 24/09/02 ===
Global White_Smoke_multiplier As Single



' 16/9/99
Declare Sub set_AO_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, nValue As Integer)

' 27/9/99
Declare Sub set_CONTROLLER_VARS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, nValue As Integer)

' 21/8/98
Declare Sub set_TA_control Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (nCHANNEL As Integer, Value As Single)


' 5/7/99 Already declared in RWJ_condense module...
'Declare Sub open_com_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ncomm As Integer)


Declare Sub get_VV_addr_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (DI_ON_VVIO_LSS As Integer, DI_OFF_VVIO_LSS As Integer)
' 16/12/98
'Global DI_ON_VVIO_LSS(250) As Integer, DI_OFF_VVIO_LSS(250) As Integer
'  16/9/99
'Global DI_ON_VVIO_LSS(300) As Integer, DI_OFF_VVIO_LSS(300) As Integer
Global DI_ON_VVIO_LSS(600) As Integer, DI_OFF_VVIO_LSS(600) As Integer


' 4/9/99
Global breakers(1 To 100, 1 To 18) As Integer
Global VB_Valves(1 To 600, 1 To 14) As Integer

' 4/9/99      set_Breakers(VB_Breakers)
Declare Sub set_Breakers Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (breakers As Integer)
'14/9/99
Declare Sub set_valves Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (valves As Integer)


' 26/5/98 Brazil Pump addresses.... ( BP_addr )
Declare Sub get_BP_addr Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (PumpIO_LSS As Integer)
Global PumpIO_LSS(1 To 15, 1 To 51) As Integer

' 26/5/98
Declare Sub get_RMA_addr_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (RMAIO_LSS As Integer)
Global RMAIO_LSS(1 To 7, 1 To 15) As Integer

' 5/6/98
Declare Sub get_elec_addr_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ESTARTSW As Integer, _
ESTOPSW As Integer, EAUTOSW As Integer, EMANSW As Integer, ESTBY1SW As Integer, ESTBY2SW As Integer, _
AVRAUTOSW As Integer, AVRMANSW As Integer, MBKSW As Integer, OPENMBKSW As Integer, AVR_pot As Integer, _
SYNCHSELSW As Integer, CHSYNCHSW As Integer, SYNCHOFFSW As Integer, SELLAMPSSW As Integer, SELSYNCHSW As Integer, _
REVRESET As Integer)
' NOTE ESTARTSW etc are declared in main.bas

' 5/6/98
Declare Sub get_bkr_addr_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (BKR As Integer)
' NOTE BKR is declared in main.bas





Declare Sub get_reals_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (RtoS As Byte, LSS_reals As Single, ALARM_HILO As Single, plot_Reals As Single)
Global LSS_reals(899) As Single

Declare Sub INIT_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" ()


' === 23/01/02 15.019 added PTT_mode ===
'Declare Sub D_MODEL_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (split_mode As Byte, ss_supply As Byte, diesel_init As Byte, elec_init As Byte, diesel_cold As Byte, elec_cold As Byte, d_run As Byte, e_run As Byte, snap_shot As Byte, hw_test As Byte)

' === 06/03/03 Added Ballast_flag
'Declare Sub D_MODEL_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (CMA_flag As Byte, PTT_mode As Byte, split_mode As Byte, ss_supply As Byte, diesel_init As Byte, elec_init As Byte, diesel_cold As Byte, elec_cold As Byte, d_run As Byte, e_run As Byte, snap_shot As Byte, hw_test As Byte)
Declare Sub D_MODEL_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (Ballast_flag As Byte, CMA_flag As Byte, PTT_mode As Byte, split_mode As Byte, ss_supply As Byte, diesel_init As Byte, elec_init As Byte, diesel_cold As Byte, elec_cold As Byte, d_run As Byte, e_run As Byte, snap_shot As Byte, hw_test As Byte)

' === 29/01/02 ===
Global CMA_flag As Byte
Global PTT_mode As Byte

' === 06/03/03 ===
Global Ballast_flag As Byte


' 28/5/98 WORKED OK  ( NB ByVal PC_INT_STRING )
'Declare Sub get_integers_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal ItoS As Byte, ByVal PC_INT_STRING As String)

' 28/5/98 ADD alarm  NB NEXT LINE CRASHES THE SIM ( DO NOT USE ByVal for alarm )
'Declare Sub get_integers_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal ItoS As Byte, ByVal PC_INT_STRING As String, ByVal alarm As Byte)

' 28/5/98 NEXT LINE WORKED OK
'Declare Sub get_integers_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal ItoS As Byte, ByVal PC_INT_STRING As String, alarm As Byte)
Declare Sub get_integers_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ByVal ItoS As Byte, ByVal PC_INT_STRING As String, alarm As Byte, elec_alarm As Byte, LSS_LOGICALS As Byte, _
active_actions As Integer, active_reset_actions As Integer, _
active_faults As Integer, _
I_sounds As Integer, _
VB_AI As Integer, _
VB_VV_extras As Integer, _
VB_steam_logicals As Byte, _
VB_steam_integers As Integer, _
VB_More_integers As Integer)

Global VB_steam_integers(1 To 100) As Integer

Global VB_steam_integers_OLD(1 To 100) As Integer


' === 7/01/02 ===
Global VB_More_integers(1 To 100) As Integer

' === 03/12/02 ===
Global VB_More_integers_OLD(1 To 100) As Integer


' 10/8/00
'Global VB_steam_logicals(1 To 100) As Byte
Global VB_steam_logicals(1 To 200) As Byte

' 22/7/99
Global DO_sett As Integer
Global old_DO_sett As Integer

Global HO_sett(2) As Integer
Global old_HO_sett(2) As Integer

Global HO_deep(2) As Integer
Global old_HO_deep(2) As Integer

Global DO_double_bottom(2) As Integer
Global OLD_DO_double_bottom(2) As Integer


' 17/12/98
'  16/9/99
'Global VB_VV_extras(50) As Integer
Global VB_VV_extras(350) As Integer

' 21/7/98
Global VB_AI(48) As Integer


' === 13/11/02 ===
Global VB_AI_old(48) As Integer



Global LSS_LOGICALS(100) As Byte

' 28/5/98 NEXT LINE DID NOT CRASH BUT DID NOT Rx valves and alarms on PC01
'Declare Sub get_integers_LSS Lib "c:\Steam_SIM\Fortran Files\Debug\Steam_SimV31.dll" (ItoS As Byte, ByVal PC_INT_STRING As String, alarm As Byte)

Global PC_INT_STRING As String * 512

' 16/12/98
'Global vv_status_LSS(250) As Integer
'  16/9/99
'Global vv_status_LSS(300) As Integer
Global vv_status_LSS(600) As Integer

' 22/9/99
'Global OLD_vv_status_LSS(300) As Integer
Global OLD_vv_status_LSS(600) As Integer

Global pp_status_LSS(51) As Integer
Global rma_status_LSS(15) As Integer
' 9/7/98
Global old_rma_status_LSS(15) As Integer

' 27/5/98
Global PC_control_LSS(13) As Integer
Global greenness_LSS(10) As Integer
Global old_greenness_LSS(10) As Integer
Global HOSEP1_flash As Integer
Global HOSEP2_flash As Integer

' 5/6/98 added
Global MSTARTSW(2) As Integer
Global MSTOPSW(2) As Integer
Global MAUTOSW(2) As Integer
Global MMANSW(2) As Integer
Global MSTBY1SW(2) As Integer
Global MSTBY2SW(2) As Integer
'''







Sub update_vvalve_LSS(Formname As Form, valve_number As Integer, indx1 As Integer, indx2 As Integer)
Dim single_char As String * 1

' 17/12/98 EXTRA valves added from 251 to 300
If valve_number < 251 Then

    single_char = Mid$(PC_INT_STRING, valve_number, 1)

    If single_char = "0" Then
        vv_status_LSS(valve_number) = 0
    ElseIf single_char = "1" Then
        vv_status_LSS(valve_number) = 1
    ElseIf single_char = "2" Then
        vv_status_LSS(valve_number) = 2
    Else
        vv_status_LSS(valve_number) = 3
    End If
    
Else
        ' 17/12/98  EXTRA valves added from 251 to 300
        vv_status_LSS(valve_number) = VB_VV_extras(valve_number - 250)

End If


'  Valves with 2 DI's and 2 DO's   ( If the second valve (indx2) does not exist then set indx2 = -1 )

If vv_status_LSS(valve_number) = 1 Then

         Formname.valve(indx1).OpenTopLeft = Not Formname.valve(indx1).OpenTopLeft
         Formname.valve(indx1).OpenBtmRight = Not Formname.valve(indx1).OpenBtmRight
                  Formname.valve(indx1).ColourOpen = &HFF00&        ' GREEN
                  
                  ' 15/12/98
                  Formname.valve(indx1).ColourClosed = &HFFFFFF      ' WHITE
                  
        If indx2 > -1 Then
            Formname.valve(indx2).OpenTopLeft = Not Formname.valve(indx2).OpenTopLeft
            Formname.valve(indx2).OpenBtmRight = Not Formname.valve(indx2).OpenBtmRight
                  Formname.valve(indx2).ColourOpen = &HFF&          ' RED
                   ' 15/12/98
                  Formname.valve(indx2).ColourClosed = &HFFFFFF      ' WHITE
       End If
        
ElseIf vv_status_LSS(valve_number) = 2 Then

         Formname.valve(indx1).OpenTopLeft = Not Formname.valve(indx1).OpenTopLeft
         Formname.valve(indx1).OpenBtmRight = Not Formname.valve(indx1).OpenBtmRight
                 Formname.valve(indx1).ColourOpen = &HFF&          ' RED
                  ' 15/12/98
                  Formname.valve(indx1).ColourClosed = &HFFFFFF      ' WHITE
                 
         If indx2 > -1 Then
               Formname.valve(indx2).OpenTopLeft = Not Formname.valve(indx2).OpenTopLeft
                Formname.valve(indx2).OpenBtmRight = Not Formname.valve(indx2).OpenBtmRight
                     Formname.valve(indx2).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                  Formname.valve(indx2).ColourClosed = &HFFFFFF      ' WHITE
         End If
          
ElseIf vv_status_LSS(valve_number) = 0 Then

         Formname.valve(indx1).OpenTopLeft = False
         Formname.valve(indx1).OpenBtmRight = False
                 Formname.valve(indx1).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                  Formname.valve(indx1).ColourClosed = &HFF&       ' RED
                  
        If indx2 > -1 Then
             Formname.valve(indx2).OpenTopLeft = True
             Formname.valve(indx2).OpenBtmRight = True
                 Formname.valve(indx2).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                  Formname.valve(indx2).ColourClosed = &HFFFFFF      ' WHITE
        End If
        
Else

          Formname.valve(indx1).OpenTopLeft = True
         Formname.valve(indx1).OpenBtmRight = True
                 Formname.valve(indx1).ColourOpen = &HFF00&        ' GREEN
                 
                  ' 15/12/98
                  Formname.valve(indx1).ColourClosed = &HFFFFFF      ' WHITE
                  
         If indx2 > -1 Then
                 Formname.valve(indx2).OpenTopLeft = False
                 Formname.valve(indx2).OpenBtmRight = False
                     Formname.valve(indx2).ColourOpen = &HFF00&        ' GREEN
                  ' 15/12/98
                  Formname.valve(indx2).ColourClosed = &HFF&          ' RED
         End If
                  
End If

End Sub
Sub valve_pcdi_LSS(nValve As Integer)

   If vv_status_LSS(nValve) = 0 Or vv_status_LSS(nValve) = 2 Then
   
            Call setPCDI(DI_ON_VVIO_LSS(nValve))
   Else
            Call setPCDI(DI_OFF_VVIO_LSS(nValve))
   End If

End Sub

Sub update_pump_LSS(frm As Form, index_no As Integer, pump_no As Integer)
Dim single_char As String * 1

single_char = Mid$(PC_INT_STRING, pump_no + 250, 1)

If single_char = "0" Then
   pp_status_LSS(pump_no) = 0
ElseIf single_char = "1" Then
   pp_status_LSS(pump_no) = 1
ElseIf single_char = "2" Then
   pp_status_LSS(pump_no) = 2
Else
   pp_status_LSS(pump_no) = 3
End If


If pp_status_LSS(pump_no) = 0 Then    '  Pump is OFF ( and  not in Stand-bye)
        frm.HavenPump1(index_no).PumpOn = False

        frm.OFFpump(index_no).BackColor = &HFFFFFF           ' WHITE
        frm.ONpump(index_no).BackColor = &HC0C0C0            ' GREY
        frm.sbyPUMP(index_no).BackColor = &HC0C0C0           ' GREY
               
ElseIf pp_status_LSS(pump_no) = 1 Then    '  Pump is OFF ( and  in Stand-bye)
        frm.HavenPump1(index_no).PumpOn = False
       
        frm.OFFpump(index_no).BackColor = &HFFFFFF          ' WHITE
        frm.ONpump(index_no).BackColor = &HC0C0C0           ' GREY
        frm.sbyPUMP(index_no).BackColor = &HFFFF&           ' BRIGHT YELLOW
              
ElseIf pp_status_LSS(pump_no) = 2 Then   '  Pump is ON ( and  in Stand-bye)
        frm.HavenPump1(index_no).PumpOn = True
       
        frm.OFFpump(index_no).BackColor = &HC0C0C0           ' DARK GREY
        frm.ONpump(index_no).BackColor = &HFF00&             ' BRIGHT GREEN
        frm.sbyPUMP(index_no).BackColor = &HFFFF&            ' BRIGHT YELLOW
        
ElseIf pp_status_LSS(pump_no) = 3 Then    '  Pump is ON ( and  not in Stand-bye)
      frm.HavenPump1(index_no).PumpOn = True
      
         frm.OFFpump(index_no).BackColor = &HC0C0C0           ' DARK GREY
         frm.ONpump(index_no).BackColor = &HFF00&              ' BRIGHT GREEN
         frm.sbyPUMP(index_no).BackColor = &HC0C0C0           ' DARK GREY
        
End If

End Sub

Sub update_rma_LSS(nRMA As Integer, set_point As Single, rma_text As String)
Dim single_char As String * 1

single_char = Mid$(PC_INT_STRING, nRMA + 310, 1)

If single_char = "0" Then
   rma_status_LSS(nRMA) = 0
ElseIf single_char = "1" Then
   rma_status_LSS(nRMA) = 1
ElseIf single_char = "2" Then
   rma_status_LSS(nRMA) = 2
ElseIf single_char = "3" Then
   rma_status_LSS(nRMA) = 3
Else
   rma_status_LSS(nRMA) = 4
End If

'  NOTE:
'      rma_status( ) = 0   LOCAL  MANUAL
'      rma_status( ) = 1   LOCAL  AUTO
'      rma_status( ) = 2   REMOTE MANUAL
'      rma_status( ) = 3   REMOTE AUTO
'      rma_status( ) = 4   REAL PID selected

If rma_status_LSS(nRMA) = 0 Then     '   (5) for MELUB etc
      rma_text = "LCL MANUAL"
'      set_point = PCCONTROL_LSS(nRMA + 40)
ElseIf rma_status_LSS(nRMA) = 1 Then
      rma_text = "LCL AUTO"
'      set_point = PCCONTROL_LSS(nRMA + 60)
ElseIf rma_status_LSS(nRMA) = 2 Then
      rma_text = "REM MANUAL"
'      set_point = PCCONTROL_LSS(nRMA)
ElseIf rma_status_LSS(nRMA) = 3 Then
      rma_text = "REM AUTO"
'     set_point = PCCONTROL_LSS(nRMA + 20)
ElseIf rma_status_LSS(nRMA) = 4 Then
      rma_text = "REAL PID"
'      set_point = PCCONTROL_LSS(nRMA + 20)
Else
     rma_text = "undefined"
'      set_point = 0#
End If

End Sub
Sub paint_alarm(objname As Object, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer)
Dim alarm1 As Integer, alarm2 As Integer, alarm3 As Integer, alarm4 As Integer

' 9/7/98 OK
If n1 > 0 Then
  alarm1 = alarm(n1)
Else
  alarm1 = False
End If
If n2 > 0 Then
  alarm2 = alarm(n2)
Else
  alarm2 = False
End If
If n3 > 0 Then
  alarm3 = alarm(n3)
Else
  alarm3 = False
End If
If n4 > 0 Then
  alarm4 = alarm(n4)
Else
  alarm4 = False
End If

If alarm1 Or alarm2 Or alarm3 Or alarm4 Then
   objname.BackColor = &HFF&                   ' RED
   objname.ForeColor = &HFFFFFF                ' WHITE
Else
   objname.BackColor = &HFF00&                 ' GREEN
   objname.ForeColor = &H0&                    ' BLACK
End If


End Sub
Sub find_VV_status_LSS(valve_number As Integer)

Dim single_char As String * 1

If valve_number < 251 Then

    single_char = Mid$(PC_INT_STRING, valve_number, 1)

    If single_char = "0" Then
        vv_status_LSS(valve_number) = 0
    ElseIf single_char = "1" Then
        vv_status_LSS(valve_number) = 1
    ElseIf single_char = "2" Then
        vv_status_LSS(valve_number) = 2
    Else
        vv_status_LSS(valve_number) = 3
    End If
    
Else
        ' 17/12/98  EXTRA valves added from 251 to 300
        vv_status_LSS(valve_number) = VB_VV_extras(valve_number - 250)

End If



End Sub
