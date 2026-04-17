Attribute VB_Name = "RWJ_condense_module"

' 15/7/99
Global Scoop_pos As Integer
Global OLD_Scoop_pos As Integer

'Global SI_units As Boolean
Global Temp_char As String      '         'F' or 'C'
Global Press_chars As String    '      ' psi' or ' Bar'
Global Flow_chars As String     '   ' lbs/hr' or ' kg/s  '

Global Temp_mult As Single
Global Temp_const As Single
Global Press_mult As Single
Global Flow_mult As Single


Global Damper_count As Integer
Global Flame_count As Integer

Global Fast_time As Boolean
Global Fast_time_count As Integer

Global run_flag As Boolean
Global boil_run_flag As Boolean

Global Pshaft As Single

' 14/6/99
Global Vscroll_val(2) As Integer

' ======== Model Variables =======

' 7/7/99
'Global Burner_ON(4) As Boolean
Global Burner_ON(4) As Byte

' 27/8/99
Global Burner1_ON(2) As Byte
Global Burner2_ON(2) As Byte

Global Bnumb As Integer
Global Boffset As Integer


Global Turbine_efficiency As Single
Global Turbine_K As Single

Global Qcondenser As Single
Global dEnth_required As Single
Global dEnth_possible As Single
Global dEnth_actual As Single

Global Psteam14_set As Single
Global P_dot As Single

' kNm
Global Torque_HP(2) As Single
Global Torque_LP(2) As Single
' kW
Global Power_HP(2) As Single
Global Power_LP(2) As Single

Global shaft_rpm As Single
Global shaft_rpm_dot As Single
Global TQ_shaft As Single
Global TQ_shaft_load As Single


Global R_HP(2) As Single       ' Bar.sec / kg
Global R_LP(2) As Single       ' Bar.sec / kg

' 21/6/99
Global TQ_HP(2) As Single       ' kNm/Bar
Global TQ_LP(2) As Single       ' kNm/Bar

' 21/6/99
Global R_turb(4) As Single
Global TQ_turb(4) As Single


Global R_turbines As Single
Global RV_ahead As Single       ' 0 to 1.0e8
Global V_ahead_pos As Single        ' 0 to 100%


' 7/6/99
Global R_HP_bleed As Single
Global R_bleed As Single
Global R_LP_bleed As Single

' 8/6/99
Global R_turbine_A As Single
Global R_turbine_B As Single
Global R_turbine_C As Single
Global Res_A As Single
Global Res_B As Single



Global Pdiff As Single
Global Rtotal As Single



Global V_attemp_pos_set(2) As Single
Global V_attemp_pos_dot As Single
Global Vattemp_POS(2) As Single


Global T_dot As Single
Global M_dot As Single

'Global t(100) As Single
'Global Q(100) As Single
'Global CQ As Single


'Global Fevap(1) As Single
' 28/6/99
'Global Fevap(2) As Single
Global Fevap(3) As Single

' MODEL variables...
Global Mwater(200) As Single
Global Twater(200) As Single

Global Msteam(200) As Single
Global Tsteam(200) As Single
Global Psteam(200) As Single
Global Hsteam(200) As Single
Global Fsteam(200) As Single

Global qDry(200) As Single

' 26/5/99 added:
Global Fwater(200) As Single


Global CP_Steam As Single

Global hf As Single
Global hfg As Single
Global hg As Single


Global Tsteel(200) As Single
Global Qsteel(200) As Single


' 27/5/99
Global Qsteel_amb(200) As Single

Global Mair(2) As Single
Global Tair(2) As Single
Global Pair(2) As Single
Global Vair(2) As Single

'Global Vwater As Single

Global Pactual(2) As Single

'Global Rmain(2) As Single


Global CPW As Single
Global CPsteam As Single
Global CPmetal As Single

Global Psat As Single
Global Tsat As Single

Global Tamb_steam As Single
Global Pamb_steam As Single



'Global Fout1 As Single
'Global Fout_pump As Single
Global Fout_air As Single
Global Fout_steam As Single



' ==== GAS Constants ====

' Universal Gas Constant R0 = 8.3143 kJ/ (kmol deg K )
'
' Molecular Mass for AIR = MMair = 28.96 kg/kmol
' Gas constant for AIR , Rair = R0/MMair = 8.3143/28.96 = 0.2871  kJ / (kg deg K )
Global Rair As Single   ' 0.2871

' Molecular Mass for STEAM = MM_steam = 18.015 kg/kmol
' Gas constant for STEAM , Rsteam = R0/MM_steam = 8.3143/18.015 = 0.4615  kJ / (kg deg K )
Global Rsteam As Single  '0.4615

' =========================

Global valve_status(200) As Integer

Global STEP   ' In milli-seconds




' =========================







Sub STENTH(Qlatent, temp)

' This routine returns the Latent Heat of Vaporistion of water
' ( NOT the total enthalpy as the name STENTH suggests. )

    If (temp < 100#) Then
        
        ' Correction RWJ 14/4/99 (See Steam Tables  hfg = Latent Heat
        ' Qlatent = 0#
          Qlatent = -2.4202 * temp + 2499.02
         
    ElseIf (temp < 180#) Then
         Qlatent = -3.01 * temp + 2562.31
    ElseIf (temp < 260#) Then
         Qlatent = -4.397 * temp + 2815.82
    Else
         Qlatent = -6.728 * temp + 3418.48
    End If

End Sub
Sub Boiler_drum_sub(Pact As Single, Tattemp As Single, Fattemp As Single, Qattemp As Single, Mwat As Single, _
    Mwat_norm As Single, Twat As Single, Fwat_in As Single, Twat_in As Single, _
    Mstm As Single, Mstm_norm As Single, Tstm As Single, Mmetal As Single, Tmetal As Single, Qmetal As Single, _
    CQwat As Single, CQstm As Single, Cevap As Single, F_evap As Single, _
    Hstm As Single, qDryness As Single)
                
              
Dim Qwat As Single, Qstm As Single
Dim Qevap As Single


Dim H_stm As Single

Dim F_evap_set  As Single
Dim F_evap_dot As Single

Dim Tsat_act As Single
Dim hf_act As Single
Dim hfg_act As Single




'  10/6/99  Qattemp is TOO SMALL ( Tsteel(4) = 367 degC when Twater = 277 deg C )
'      0.000565 changed to 0.00565
'  10/6/99 ( Tsteel(4) = 278 degC when Twater = 277 deg C )   NOW OK
    Qattemp = (Tattemp - Twat) * 0.00565 * Mwat


     Qwat = (Tmetal - Twat) * CQwat * Mwat + Qattemp
     Qstm = (Tmetal - Tstm) * CQstm * Mstm

' Metal temperature
If Mmetal > 1# Then
     T_dot = ((Qmetal - Qwat - Qstm) / CPmetal) / Mmetal
     Tmetal = Tmetal + T_dot * STEP
End If



' BOILING or CONDENSING ?

' Find the Saturation Pressure of the water...
Call get_Psat(Psat, Twat)
' Find the enthalpy hf and hfg for the water...
Call get_hf_hfg(Psat, hf, hfg)

hg = hf + hfg


If Psat > Pact Then
    F_evap_set = (Psat - Pact) * Cevap * Mwat / Mwat_norm       ' BOILING
    
    ' 27/5/99 F_evap_set went to INFINITY at > 100 bar pressure SO...
    If (F_evap_set > 25#) Then F_evap_set = 25#
    
ElseIf Psat < Pact Then
    F_evap_set = (Psat - Pact) * Cevap * Mstm / Mstm_norm       ' CONDENSING
    
    ' 27/5/99 F_evap_set went to INFINITY at > 100 bar pressure SO...
    If (F_evap_set < -25#) Then F_evap_set = -25#
    
Else
        F_evap_set = 0#
End If

' 17/5/99  Fevap does not change instantaneously....
F_evap_dot = (F_evap_set - F_evap) * 0.2
F_evap = F_evap + F_evap_dot * STEP


    Qevap = 0#


' WATER temperature
If Mwat <= 0.001 Then
    Twat = Tmetal
    If F_evap > 0# Then
' We can't evaporate water which doesn't exist...
        Mwat = 0.001
        F_evap = 0#
    End If
Else
        Qevap = F_evap * hfg
        
        ' add water inlet from feed pumps Fwat_in
         T_dot = ((Qwat - Qevap) / CPW + Fwat_in * (Twat_in - Twat)) / (Mwat + 1#)
               
        Twat = Twat + T_dot * STEP
        
End If
    
   
' STEAM temperature
If Mstm <= 0.001 Then

    ' Tstm will follow Tmetal but "lag" behind it
    T_dot = (Tmetal - Tstm) * 0.1
    Tstm = Tstm + T_dot * STEP
    
    If F_evap < 0# Then
' We can't condense steam which doesn't exist...
        Mstm = 0.001
        F_evap = 0#
    End If
    
Else

    ' T_dot = ((Qstm + Fevap * hfg) / CPsteam + F1 * (T1 - Tstm)) / (Mstm + 1#)
    ' T_dot = ((Qstm / CPsteam + Qevap / CPW) + F1 * (T1 - Tstm)) / (Mstm + 1#)
    
    
     
      T_dot = (Qstm / CPsteam + F_evap * (Twat - Tstm)) / (Mstm + 1#)
              
              
      Tstm = Tstm + T_dot * STEP
      
End If


' WATER mass
M_dot = Fwat_in - F_evap
Mwat = Mwat + M_dot * STEP

' STEAM mass
M_dot = F_evap
Mstm = Mstm + M_dot * STEP



' OUTPUT conditions ( Temperature,dryness,Enthalpy )
If Mstm < 0.0015 Then

        Hstm = 0#
    qDryness = 0#
   
Else

    ' Find the Saturation Temperature ( Tsat_act ) at the actual pressure Pact...
    Call get_Tsat(Pact, Tsat_act)
    ' Find the enthalpy hf and hfg at the actual pressure Pact...
    Call get_hf_hfg(Pact, hf_act, hfg_act)
    ' Find the enthalpy H of the steam...
    ' Note: Assume DRY steam because this is a heater/condenser ( NOT a turbine )
    Call get_Hsteam(Pact, Tsat_act, H_stm, hf_act, hfg_act, Tstm, 1#, CP_Steam)
    
            Hstm = H_stm
        qDryness = 0.96
   
End If

If Mwat < 0.0015 Then

    hf = 0#
    hfg = 0#
    
End If


End Sub
Sub Heater_sub(Pact As Single, Tstm_IN As Single, Fstm_IN As Single, Mstm As Single, Tstm As Single, _
    Mmetal As Single, Tmetal As Single, Qmetal As Single, _
    CQstm As Single, Hstm As Single, qDryness As Single)


' 29/6/99 above variables as SINGLE




'   26/5/99

' Local variables...
'Dim Zmass As Single
'Dim Qwat As Single

Dim Qstm As Single
Dim H_stm As Single

'Dim Qevap As Single

'Dim F_evap_set  As Single
'Dim F_evap_dot As Single

Dim Tsat_act As Single
Dim hf_act As Single
Dim hfg_act As Single



' HEAT TRANSFER (kW)...
' 27/5/99   Qmetal is now an INPUT to this sub
'    Qmetal = (Tx - Tmetal) * CQmetal  ' Goes negative if cooling (eg Tx=Tamb_steam)

    
    ' Qwat = (Tmetal - Twat) * CQwat * Mwat
     
     Qstm = (Tmetal - Tstm) * CQstm * Mstm

' Metal temperature
'Zmass = Mmetal + Mwat + Mstm
If Mmetal > 1# Then
     T_dot = ((Qmetal - Qwat - Qstm) / CPmetal) / Mmetal
     Tmetal = Tmetal + T_dot * STEP
End If

' STEAM temperature
'If Mstm <= 0.001 Then
If Fstm_IN <= 0.001 Then

      ' Tstm will follow Tmetal but "lag" behind it
      T_dot = (Tmetal - Tstm) * 0.1
       Tstm = Tstm + T_dot * STEP
    
      '  T2 = Tmetal
         
   qDryness = 0#
    
Else
        ' Need to calculate CPsteam here
      T_dot = (Qstm / CPsteam + Fstm_IN * (Tstm_IN - Tstm)) / (Mstm + 1#)
       Tstm = Tstm + T_dot * STEP
      
      '  T2 = Tstm
      
   qDryness = 1#
     
End If


' OUTPUT conditions ( Enthalpy )
'If Mstm < 0.0015 Then
If Fstm_IN < 0.0015 Then

     Hstm = 0#
   
Else

    ' Find the Saturation Temperature ( Tsat_act ) at the actual pressure Pact...
    Call get_Tsat(Pact, Tsat_act)
    ' Find the enthalpy hf and hfg at the actual pressure Pact...
    Call get_hf_hfg(Pact, hf_act, hfg_act)
    ' Find the enthalpy H of the steam...
    ' Note: Assume DRY steam because this is a heater/condenser ( NOT a turbine )
    Call get_Hsteam(Pact, Tsat_act, H_stm, hf_act, hfg_act, Tstm, 1#, CP_Steam)
    
    Hstm = H_stm
    
End If



End Sub
Sub Heater_mix_sub(Pact As Single, Tstm_IN_a As Single, Fstm_IN_a As Single, _
    Tstm_IN_b As Single, Fstm_IN_b As Single, Mstm As Single, Tstm As Single, _
    Mmetal As Single, Tmetal As Single, Qmetal As Single, _
    CQstm As Single, Hstm As Single, qDryness As Single)



' 29/6/99   above variables as SINGLE



'   26/5/99

' Local variables...
'Dim Zmass As Single
'Dim Qwat As Single

Dim Qstm As Single
Dim H_stm As Single

'Dim Qevap As Single

'Dim F_evap_set  As Single
'Dim F_evap_dot As Single

Dim Tsat_act As Single
Dim hf_act As Single
Dim hfg_act As Single



' HEAT TRANSFER (kW)...
' 27/5/99   Qmetal is now an INPUT to this sub
'    Qmetal = (Tx - Tmetal) * CQmetal  ' Goes negative if cooling (eg Tx=Tamb_steam)

    
    ' Qwat = (Tmetal - Twat) * CQwat * Mwat
     
     Qstm = (Tmetal - Tstm) * CQstm * Mstm

' Metal temperature
'Zmass = Mmetal + Mwat + Mstm
If Mmetal > 1# Then
     T_dot = ((Qmetal - Qwat - Qstm) / CPmetal) / Mmetal
     Tmetal = Tmetal + T_dot * STEP
End If

' STEAM temperature
'If Mstm <= 0.001 Then
If (Fstm_IN_a + Fstm_IN_b) <= 0.001 Then

      ' Tstm will follow Tmetal but "lag" behind it
      T_dot = (Tmetal - Tstm) * 0.1
       Tstm = Tstm + T_dot * STEP
    
      '  T2 = Tmetal
         
   qDryness = 0#
    
Else
        ' Need to calculate CPsteam here
      T_dot = (Qstm / CPsteam + Fstm_IN_a * (Tstm_IN_a - Tstm) + Fstm_IN_b * (Tstm_IN_b - Tstm)) / (Mstm + 1#)
       Tstm = Tstm + T_dot * STEP
      
      '  T2 = Tstm
      
   qDryness = 1#
     
End If


' OUTPUT conditions ( Enthalpy )
'If Mstm < 0.0015 Then
'If Fstm_IN < 0.0015 Then
If (Fstm_IN_a + Fstm_IN_b) <= 0.001 Then

     Hstm = 0#
   
Else

    ' Find the Saturation Temperature ( Tsat_act ) at the actual pressure Pact...
    Call get_Tsat(Pact, Tsat_act)
    ' Find the enthalpy hf and hfg at the actual pressure Pact...
    Call get_hf_hfg(Pact, hf_act, hfg_act)
    ' Find the enthalpy H of the steam...
    ' Note: Assume DRY steam because this is a heater/condenser ( NOT a turbine )
    Call get_Hsteam(Pact, Tsat_act, H_stm, hf_act, hfg_act, Tstm, 1#, CP_Steam)
    
    Hstm = H_stm
    
End If



End Sub
Sub Turbine_sub(Pin, Pout, Turbo_K, Fsteam, Hin, Hout, Turbo_Torque, Turbo_power, Tout, qDry_out)
'              (IN , IN  ,    IN  ,    IN ,  IN,  OUT,          OUT,         OUT,  OUT,    OUT  )

Dim Shaft_K As Single       ' 0 to 1.0
Dim dEnthalpy As Single
Dim Tsat_out_equiv As Single
Dim hf_out As Single
Dim hfg_out As Single


' See Turbine characteristic ( Marine Steam engines and turbines Page 551 )
Shaft_K = 2# - shaft_rpm * 0.01     ' Varies from 2.0 at zero shaft speed  to 1.0 at full shaft speed

' Turbine Torque developed is a function of pressure difference...
' See Marine Steam Engines and Turbines (Butterworths) , page 538 ( Reduced Power operation )

' 1/7/99
'Turbo_Torque = (Pin - Pout) * Turbo_K * Shaft_K                                    ' kNm
 Turbo_Torque = (Pin - Pout) * Turbo_K * Shaft_K * Turbine_efficiency                ' kNm



' Power (kW) = Torque (kNm) * Angular Velocity ( radians/sec )
'Turbo_power = Turbo_Torque [kNm] * Shaft_speed [rads/sec]       ' kJ/sec OR kW
 Turbo_power = Turbo_Torque * shaft_rpm * 2# * 3.142 / 60#       ' kW

' Power = Enthalpy drop [kJ/kg] * Fsteam [kg/sec]           i.e  ' kJ/sec OR kW
' SO: Enthalpy drop [kJ/kg] = Power [kW] / Fsteam [kg/sec]
If Fsteam > 0.001 Then
        dEnthalpy = Turbo_power / Fsteam        '(kW) /(kg/s) =  (kJ/s) /(kg/s) =  kJ/kg
Else
        dEnthalpy = 0#         ' kJ/kg
End If

If (Hin - dEnthalpy > 0#) Then
    Hout = Hin - dEnthalpy
Else
    Hout = 0#
End If



Call get_Tsat(Pout, Tsat_out_equiv)
Call get_hf_hfg(Pout, hf_out, hfg_out)

'Call get_Tsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_steam)
Call get_Tsteam(Pout, Tsat_out_equiv, Hout, hf_out, hfg_out, Tout, qDry_out, CP_Steam)





End Sub
Sub Boiler_sub(Boil_numb As Integer)

' 14/6/99 For array storage...
If Boil_numb = 1 Then
    Boffset = 0
    Bnumb = 1
Else
    Boffset = 100
    Bnumb = 2
End If

'====================================================================================
        
' Pair is in kN/m^2...
  Pair(Bnumb) = Mair(Bnumb) * Rair * (Tair(Bnumb) + 273.15) / Vair(Bnumb)   '  from PV = mRT , P = mRT/V
' Convert Pair from kN/m^2 to BAR...( 100 kN/m^2 = 1 Bar )
  Pair(Bnumb) = Pair(Bnumb) * 0.01
  Tair(Bnumb) = Tsteam(Boffset + 1)
                    
' Psteam(1) in kNm...( NB Vsteam = Vair )
  Psteam(Boffset + 1) = Msteam(Boffset + 1) * Rsteam * (Tsteam(Boffset + 1) + 273.15) / Vair(Bnumb) '  from PV = mRT , P = mRT/V
' Psteam(1) in BAR...
  Psteam(Boffset + 1) = Psteam(Boffset + 1) * 0.01

          
' Dalton's Law of Partial Pressures...
  Pactual(Bnumb) = Pair(Bnumb) + Psteam(Boffset + 1)
        
        
' Need about 19,000 kW at present to generate 11.7 kg/sec at 60 Bar
           
' Attemperator/Super-Heater Mixing....
  Qsteel_amb(Boffset + 5) = (Tsteel(Boffset + 5) - Tamb_steam) * 0.4  ' Heat loss to ambient
  Qsteel(Boffset + 5) = -Qsteel_amb(Boffset + 5)                ' Heat loss to ambient
          
If Burner_ON(Bnumb) Then
            
                 
    ' BOILER DRUM HEAT...
    ' Need about 19,000 kW at present to generate 11.7 kg/sec at 60 Bar
          
    ' 10/6/99   Tsteel(1) approx = 278 deg C (at 60 Bar) ,  1300-278=1022
    Qsteel(Boffset + 1) = (1300# - Tsteel(Boffset + 1)) / 1022# * Vscroll_val(Bnumb) * 19000# / 20000# ' 0 to 19,000 kW
            
             
    ' FIRST SUPER-HEATER...CQiron = 3295 kW / ( 1300.0 - 410.0 ) deg C = 3.702
    ' 10/6/99   Tsteel(2) approx = 510 deg C (at 60 Bar) , 1300-510=790
    Qsteel(Boffset + 2) = (1300# - Tsteel(Boffset + 2)) / 790# * Vscroll_val(Bnumb) * 3295# * 1.09 / 20000#    ' 0  to 3295 kW
                 
                 
                 
    ' SECOND SUPER-HEATER...
    ' 10/6/99   Tsteel(3) approx = 610 deg C (at 60 Bar) , 1300-610=690
    Qsteel(Boffset + 3) = (1300# - Tsteel(Boffset + 3)) / 690# * Vscroll_val(Bnumb) * 3295# / 20000#     ' 0  to 3295 kW
                                                            
                                            
Else
                
    ' 10/6/99   Tsteel(1) approx = 278 deg C (at 60 Bar) , 1300-278=1022
    Qsteel(Boffset + 1) = (Tamb_steam - Tsteel(Boffset + 1)) / 1022# * 19000# ' 19,000 kW max
                
    ' 10/6/99   Tsteel(2) approx = 510 deg C (at 60 Bar) , 1300-510=790
    Qsteel(Boffset + 2) = (Tamb_steam - Tsteel(Boffset + 2)) / 790# * 3295# * 1.09  ' 0  to 3295 kW max
              
    ' 10/6/99   Tsteel(3) approx = 550 deg C (at 60 Bar) , 1300-550=750
    Qsteel(Boffset + 3) = (Tamb_steam - Tsteel(Boffset + 3)) / 690# * 3295#    ' 0  to 3295 kW max

                 
End If
                                                                                                                            
  
  
   ' Boiler_drum_sub(Pact, Tattemp, Fattemp, Qattemp, Mwat, Mwat_norm, Twat, Fwat_in, Twat_in, _
   '             Mstm, Mstm_norm, Tstm, Mmetal, Tmetal, Qmetal, _
   '            CQwat, CQstm, Cevap, F_evap, _
   '             Hstm, qDryness)
                
            
            
            
      Call Boiler_drum_sub(Pactual(Bnumb), Tsteam(Boffset + 4), Fsteam(Boffset + 2), Qsteel(Boffset + 4), _
            Mwater(Boffset + 1), 10000#, Twater(Boffset + 1), Fwater(Boffset + 10), Twater(Boffset + 10), _
            Msteam(Boffset + 1), 60#, Tsteam(Boffset + 1), 10000#, Tsteel(Boffset + 1), Qsteel(Boffset + 1), _
             1#, 1#, 100#, Fevap(Bnumb), _
            Hsteam(Boffset + 1), qDry(Boffset + 1))
            
                                                  
                            
    ' FIRST SUPER-HEATER...CQiron = 3295 kW / ( 1300.0 - 410.0 ) deg C = 3.702 (make CQ_steam same..)

    Call Heater_sub(Pactual(Bnumb), Tsteam(Boffset + 1), Fsteam(Boffset + 1), 10#, Tsteam(Boffset + 2), _
            500#, Tsteel(Boffset + 2), Qsteel(Boffset + 2), _
            3.702, Hsteam(Boffset + 2), qDry(Boffset + 2))
                                     
                               
    '  ATTEMPARATOR CONTROL VALVE  ( 506 to 514 degC )
    If Tsteam(Boffset + 5) < 506# Then
            V_attemp_pos_set(Bnumb) = 0#
    ElseIf Tsteam(Boffset + 5) > 514# Then
            V_attemp_pos_set(Bnumb) = 100#
    Else
            V_attemp_pos_set(Bnumb) = (Tsteam(Boffset + 5) - 506#) / 8# * 100#   ' Linear operation
    End If
        
    'ATTEMPARATOR CONTROL VALVE
            V_attemp_pos_dot = (V_attemp_pos_set(Bnumb) - Vattemp_POS(Bnumb)) * 0.2
            Vattemp_POS(Bnumb) = Vattemp_POS(Bnumb) + V_attemp_pos_dot * STEP
        
        
            Fsteam(Boffset + 2) = Fsteam(Boffset + 1) * Vattemp_POS(Bnumb) * 0.01
            Fsteam(Boffset + 3) = Fsteam(Boffset + 1) - Fsteam(Boffset + 2)
            
        
    ' SECOND SUPER-HEATER...CQiron = 3295 kW / ( 1300.0 - 538.0 ) deg C = 4.324 (make CQ_steam same..)

    Call Heater_sub(Pactual(Bnumb), Tsteam(Boffset + 2), Fsteam(Boffset + 3), 10#, Tsteam(Boffset + 3), _
            500#, Tsteel(Boffset + 3), Qsteel(Boffset + 3), _
            3.702, Hsteam(Boffset + 3), qDry(Boffset + 3))
                    
                                 
    '  ATTEMPARATOR
    ' SEE Boiler_drum_sub....     Qattemp = qsteel(4)

    Call Heater_sub(Pactual(Bnumb), Tsteam(Boffset + 2), Fsteam(Boffset + 2), 10#, Tsteam(Boffset + 4), _
            500#, Tsteel(Boffset + 4), -Qsteel(Boffset + 4), _
            3.702, Hsteam(Boffset + 4), qDry(Boffset + 4))
                    
                    
    '  ATTEMPARATOR mixes with the SECOND Stage Super-Heater

    Call Heater_mix_sub(Pactual(Bnumb), Tsteam(Boffset + 4), Fsteam(Boffset + 2), Tsteam(Boffset + 3), _
            Fsteam(Boffset + 3), 10#, Tsteam(Boffset + 5), 500#, Tsteel(Boffset + 5), Qsteel(Boffset + 5), _
            3.702, Hsteam(Boffset + 5), qDry(Boffset + 5))
                    
                    
                    
End Sub

Sub Condense_sub(Pact, Mwat, Mwat_norm, Twat, Fwat_out, _
                Mstm, Mstm_norm, Tstm_IN, Fstm_IN, Tstm, Mmetal, Tmetal, Qmetal, _
               CQwat, CQstm, Cevap, F_evap, _
                Hstm, qDryness)
                
                

' Local variables...
Dim Qwat As Single, Qstm As Single
Dim Qevap As Single


Dim H_stm As Single

Dim F_evap_set  As Single
Dim F_evap_dot As Single

Dim Tsat_act As Single
Dim hf_act As Single
Dim hfg_act As Single



     Qwat = (Tmetal - Twat) * CQwat * Mwat
     Qstm = (Tmetal - Tstm) * CQstm * Mstm


' Metal temperature
'Zmass = Mmetal + Mwat + Mstm
If Mmetal > 1# Then
     T_dot = ((Qmetal - Qwat - Qstm) / CPmetal) / Mmetal
     Tmetal = Tmetal + T_dot * STEP
End If



' BOILING or CONDENSING ?

' Find the Saturation Pressure of the water...
Call get_Psat(Psat, Twat)
' Find the enthalpy hf and hfg for the water...
Call get_hf_hfg(Psat, hf, hfg)

hg = hf + hfg


' 17/5/99
If Psat > Pact Then
  ' F_evap = (Psat - Pact) * Cevap                              ' BOILING
    F_evap_set = (Psat - Pact) * Cevap * Mwat / Mwat_norm       ' BOILING
    
    ' 27/5/99 F_evap_set went to INFINITY at > 100 bar pressure SO...
    If (F_evap_set > 25#) Then F_evap_set = 25#
    
ElseIf Psat < Pact Then
  ' F_evap = (Psat - Pact) * Cevap * 0.1 * Mstm                 ' CONDENSING
    F_evap_set = (Psat - Pact) * Cevap * Mstm / Mstm_norm       ' CONDENSING
    
    ' 27/5/99 F_evap_set went to INFINITY at > 100 bar pressure SO...
    If (F_evap_set < -25#) Then F_evap_set = -25#
    
Else
        F_evap_set = 0#
End If

' 17/5/99  Fevap does not change instantaneously....
F_evap_dot = (F_evap_set - F_evap) * 0.2
F_evap = F_evap + F_evap_dot * STEP


    Qevap = 0#


' WATER temperature
If Mwat <= 0.001 Then
    Twat = Tmetal
    If F_evap > 0# Then      ' We can't evaporate water which doesn't exist...
        Mwat = 0.001
        F_evap = 0#
    End If
Else
        Qevap = F_evap * hfg
        
     ' 28/6/99There is NO Fwat_in.....
     '   T_dot = ((Qwat - Qevap) / CPW + Fwat_in * (Twat_in - Twat)) / (Mwat + 1#)
         T_dot = ((Qwat - Qevap) / CPW) / (Mwat + 1#)
               
        Twat = Twat + T_dot * STEP
        
End If
    
   
' STEAM temperature
If Mstm <= 0.001 Then

    ' Not strictly true...
    ' Tstm = Tmetal
    ' Tstm will follow Tmetal but "lag" behind it
    T_dot = (Tmetal - Tstm) * 0.1
    Tstm = Tstm + T_dot * STEP
    
    If F_evap < 0# Then      ' We can't condense steam which doesn't exist...
        Mstm = 0.001
        F_evap = 0#
    End If
    
Else
    ' T_dot = ((Qstm + Fevap * hfg) / CPsteam + F1 * (T1 - Tstm)) / (Mstm + 1#)
    ' T_dot = ((Qstm / CPsteam + Qevap / CPW) + F1 * (T1 - Tstm)) / (Mstm + 1#)
    
    
    '  T_dot = (Qstm / CPsteam + F_evap * (Twat - Tstm) + F1 * (T1 - Tstm)) / (Mstm + 1#)
    ' 1/6/99 There is a flow of steam into the boiler (Fstm_IN)....
    ' T_dot = (Qstm / CPsteam + F_evap * (Twat - Tstm)) / (Mstm + 1#)
      T_dot = (Qstm / CPsteam + F_evap * (Twat - Tstm) + Fstm_IN * (Tstm_IN - Tstm)) / (Mstm + 1#)
              
              
      Tstm = Tstm + T_dot * STEP
      
End If


' WATER mass
' 26/5/99
'M_dot = -F_evap
M_dot = -Fwat_out - F_evap
Mwat = Mwat + M_dot * STEP

' STEAM mass
'M_dot = F_evap
M_dot = F_evap
Mstm = Mstm + M_dot * STEP


' OUTPUT conditions ( Temperature,dryness)
'If F1 > 0.001 Then
'    T2 = Tstm
'    qDryness = 1#
'Else
'    T2 = Tmetal
'    qDryness = 0#
'End If


' OUTPUT conditions ( Temperature,dryness,Enthalpy )
If Mstm < 0.0015 Then

        Hstm = 0#
    qDryness = 0#
   
Else

    ' Find the Saturation Temperature ( Tsat_act ) at the actual pressure Pact...
    Call get_Tsat(Pact, Tsat_act)
    ' Find the enthalpy hf and hfg at the actual pressure Pact...
    Call get_hf_hfg(Pact, hf_act, hfg_act)
    ' Find the enthalpy H of the steam...
    ' Note: Assume DRY steam because this is a heater/condenser ( NOT a turbine )
    Call get_Hsteam(Pact, Tsat_act, H_stm, hf_act, hfg_act, Tstm, 1#, CP_Steam)
    
            Hstm = H_stm
        qDryness = 0.96
   
End If

If Mwat < 0.0015 Then

    hf = 0#
    hfg = 0#
    
End If

End Sub

Sub boiler_drum(Boiler_numb As Integer, Boiler_offset As Integer)


Dim Qwat As Single, Qstm As Single
Dim Qevap As Single

' 1/6/99
'Dim Qattemp As Single

Dim H_stm As Single

Dim F_evap_set  As Single
Dim F_evap_dot As Single

Dim Tsat_act As Single
Dim hf_act As Single
Dim hfg_act As Single

Dim Cevap As Single



' Attemperation...
Qsteel(Boiler_offset + 4) = (Tsteam(Boiler_offset + 4) - Twater(Boiler_offset + 1)) * 0.00565 * Mwater(Boiler_offset + 1)

' CQwat = 1.0
' CQstm = 1.0
Qwat = (Tsteel(Boiler_offset + 1) - Twater(Boiler_offset + 1)) * Mwater(Boiler_offset + 1) + Qsteel(Boiler_offset + 4)

Qstm = (Tsteel(Boiler_offset + 1) - Tsteam(Boiler_offset + 1)) * Msteam(Boiler_offset + 1)


' Metal temperature
T_dot = ((Qsteel(Boiler_offset + 1) - Qwat - Qstm) / CPmetal) / 10000#
Tsteel(Boiler_offset + 1) = Tsteel(Boiler_offset + 1) + T_dot * STEP



' BOILING or CONDENSING ?

' Find the Saturation Pressure of the water...

Call get_Psat(Psat, Twater(Boiler_offset + 1))

' Find the enthalpy hf and hfg for the water...
Call get_hf_hfg(Psat, hf, hfg)

hg = hf + hfg


Cevap = 100#

If Psat > Pactual(Boiler_numb) Then

    F_evap_set = (Psat - Pactual(Boiler_numb)) * Cevap * Mwater(Boiler_offset + 1) / 10000#      ' BOILING
    
    ' 27/5/99 F_evap_set went to INFINITY at > 100 bar pressure SO...
    If (F_evap_set > 25#) Then F_evap_set = 25#
    
ElseIf Psat < Pactual(Boiler_numb) Then

    F_evap_set = (Psat - Pactual(Boiler_numb)) * Cevap * Msteam(Boiler_offset + 1) / 60#      ' CONDENSING
    
    ' 27/5/99 F_evap_set went to INFINITY at > 100 bar pressure SO...
    If (F_evap_set < -25#) Then F_evap_set = -25#
    
Else
        F_evap_set = 0#
End If

' 17/5/99  Fevap does not change instantaneously....
F_evap_dot = (F_evap_set - Fevap(Boiler_numb)) * 0.2
Fevap(Boiler_numb) = Fevap(Boiler_numb) + F_evap_dot * STEP


    Qevap = 0#


' WATER temperature
If Mwater(Boiler_offset + 1) <= 0.001 Then

    Twater(Boiler_offset + 1) = Tsteel(Boiler_offset + 1)
    If Fevap(Boiler_numb) > 0# Then      ' We can't evaporate water which doesn't exist...
        Mwater(Boiler_offset + 1) = 0.001
        Fevap(Boiler_numb) = 0#
    End If
    
Else
    Qevap = Fevap(Boiler_numb) * hfg
        
    T_dot = ((Qwat - Qevap) / CPW + Fwater(Boiler_offset + 10) * (Twater(Boiler_offset + 10) - Twater(Boiler_offset + 1))) / (Mwater(Boiler_offset + 1) + 1#)
               
    Twater(Boiler_offset + 1) = Twater(Boiler_offset + 1) + T_dot * STEP
        
End If
    
   
' STEAM temperature
If Msteam(Boiler_offset + 1) <= 0.001 Then

    ' Tsteam(Boiler_offset+1) will follow Tsteel(Boiler_offset+1) but "lag" behind it
    T_dot = (Tsteel(Boiler_offset + 1) - Tsteam(Boiler_offset + 1)) * 0.1
    Tsteam(Boiler_offset + 1) = Tsteam(Boiler_offset + 1) + T_dot * STEP
    
    If Fevap(Boiler_numb) < 0# Then      ' We can't condense steam which doesn't exist...
        Msteam(Boiler_offset + 1) = 0.001
        Fevap(Boiler_numb) = 0#
    End If
    
Else
     
    T_dot = (Qstm / CPsteam + Fevap(Boiler_numb) * (Twater(Boiler_offset + 1) - Tsteam(Boiler_offset + 1))) / (Msteam(Boiler_offset + 1) + 1#)
                            
    Tsteam(Boiler_offset + 1) = Tsteam(Boiler_offset + 1) + T_dot * STEP
      
End If


' WATER mass
M_dot = Fwater(Boiler_offset + 10) - Fevap(Boiler_numb)
Mwater(Boiler_offset + 1) = Mwater(Boiler_offset + 1) + M_dot * STEP

' STEAM mass
M_dot = Fevap(Boiler_numb)
Msteam(Boiler_offset + 1) = Msteam(Boiler_offset + 1) + M_dot * STEP



' OUTPUT conditions ( Temperature,dryness,Enthalpy )
If Msteam(Boiler_offset + 1) < 0.0015 Then

        Hsteam(Boiler_offset + 1) = 0#
    qDry(Boiler_offset + 1) = 0#
   
Else

    ' Find the Saturation Temperature ( Tsat_act ) at the actual pressure Pactual(Boiler_numb)...
    Call get_Tsat(Pactual(Boiler_numb), Tsat_act)
    
    ' Find the enthalpy hf and hfg at the actual pressure Pactual(Boiler_numb)...
    Call get_hf_hfg(Pactual(Boiler_numb), hf_act, hfg_act)
    
    ' Find the enthalpy H of the steam...
    ' Note: Assume DRY steam because this is a heater/condenser ( NOT a turbine )
    Call get_Hsteam(Pactual(Boiler_numb), Tsat_act, H_stm, hf_act, hfg_act, Tsteam(Boiler_offset + 1), 1#, CP_Steam)
    
            Hsteam(Boiler_offset + 1) = H_stm
        qDry(Boiler_offset + 1) = 0.96
   
End If

If Mwater(Boiler_offset + 1) < 0.0015 Then

    hf = 0#
    hfg = 0#
    
End If





End Sub




