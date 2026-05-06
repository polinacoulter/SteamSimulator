Attribute VB_Name = "Ballast_module"

' === 28/05/02 ===

Global TIME_STEP As Single

Global R_ballast(1 To 40) As Single
Global F_ballast(1 To 40) As Single
Global P_ballast(1 To 40) As Single
Global T_ballast_valve(1 To 40) As Single

Global Open_ballast_valve(1 To 40) As Boolean

Global ZWP(1 To 4) As Single
Global ZWS(1 To 4) As Single
'
Global B_pump_ON(1 To 2) As Boolean
'
'


Sub Ballast_Init()
Dim i As Integer

 
 TIME_STEP = 0.1
 
For i = 1 To 40

    R_ballast(i) = 0.01
    F_ballast(i) = 0#
    P_ballast(i) = 0#
    T_ballast_valve(i) = 0#
    Open_ballast_valve(i) = False
Next i

' Sea pressure (suction and discharge)
P_ballast(36) = 1#
 
For i = 1 To 4
    ZWP(i) = 0#
    ZWS(i) = 0#
Next i

    ZWP(4) = 0.142
    ZWP(3) = 0.113
    ZWP(2) = 0.151
    ZWP(1) = 0.111

    ZWS(4) = 0.151
    ZWS(3) = 0.123
    ZWS(2) = 0.166
    ZWS(1) = 0.097


For i = 1 To 2
    B_pump_ON(i) = False
Next i


End Sub
Sub Ballast()
Dim i As Integer
Dim PB_set As Single
Dim PB_dot As Single
Dim Z_dot(1 To 4) As Single
Dim F_ball_tank(1 To 4) As Single
Dim P_dot As Single


Dim F_to_tanks As Single


 
' Stbd Sea Suction Flow
P_ballast(30) = P_ballast(36)
F_ballast(8) = FLOW_One_Way(30, 2, 8, 8)

' Stbd Ring Main to Manifold Flow
F_ballast(10) = FLOW_One_Way(4, 2, 10, 10)
 
' Stbd Pump OUT to Over-board valve Flow
F_ballast(12) = FLOW_One_Way(6, 8, 12, 12)
 
' Stbd Pump OUT to re-circ valve Flow
F_ballast(13) = FLOW_One_Way(6, 7, 13, 13)
 
' re-circ valve to Stbd Main Flow
F_ballast(25) = FLOW_Two_Ways(7, 4, 25, 25)
 
 
' Port Sea Suction Flow
P_ballast(31) = P_ballast(36)
F_ballast(2) = FLOW_One_Way(31, 1, 2, 2)

' Port Ring Main to Manifold Flow
F_ballast(30) = FLOW_One_Way(3, 1, 30, 30)

' Port Pump OUT to Over-board valve Flow
F_ballast(5) = FLOW_One_Way(5, 8, 5, 5)

' Port Pump OUT to re-circ valve Flow
F_ballast(4) = FLOW_One_Way(5, 7, 4, 4)

' re-circ valve to Port Main Flow
F_ballast(23) = FLOW_Two_Ways(7, 3, 23, 23)


' Main Isolating Valve Near side
F_ballast(24) = FLOW_Two_Ways(3, 4, 24, 24)

' Main Isolating Valve Far side
F_ballast(22) = FLOW_Two_Ways(9, 10, 22, 22)

' Over-board valve to OVERBOARD Flow
P_ballast(32) = P_ballast(36)
F_ballast(26) = FLOW_One_Way(8, 32, 26, 26)


' === 29/05/02 4 VALVES ADDED ( 27,28,29,33 )===
' Port suction isolation
F_ballast(27) = FLOW_Two_Ways(3, 9, 27, 27)
' Stbd suction isolation
F_ballast(28) = FLOW_Two_Ways(4, 10, 28, 28)

' Common dis-charge to Port tanks
F_ballast(29) = FLOW_Two_Ways(7, 9, 29, 29)
' Common dis-charge to Stbd tanks
F_ballast(33) = FLOW_Two_Ways(7, 10, 33, 33)



' Ballast Pumps  ( 1 = Port , 2 = Stbd )

For i = 1 To 2

    If B_pump_ON(i) Then
        PB_set = 3#
    Else
        PB_set = 0#
    End If
    
     PB_dot = (PB_set - P_ballast(i + 36)) / 2#
     P_ballast(i + 36) = P_ballast(i + 36) + PB_dot * TIME_STEP
    
Next i

' Port
If P_ballast(1) < -0.98 Then        ' Gone TOO far into vacuum so...
     P_ballast(39) = 0#
Else
     P_ballast(39) = P_ballast(1) + P_ballast(37)
End If
Open_ballast_valve(39) = True       ' Pump DUMMY VALVE
F_ballast(39) = FLOW_One_Way(39, 5, 39, 39)


' Stbd
If P_ballast(2) < -0.98 Then        ' Gone TOO far into vacuum so...
     P_ballast(40) = 0#
Else
     P_ballast(40) = P_ballast(2) + P_ballast(38)
End If
Open_ballast_valve(40) = True       ' Pump DUMMY VALVE
F_ballast(40) = FLOW_One_Way(40, 6, 40, 40)


' Port Ballast Tanks
F_ballast(14) = FLOW_Two_Ways(9, 14, 14, 14)
P_ballast(14) = ZWP(4) * 1.1
If ZWP(4) < 0.005 Then              ' The Tank is EMPTY
    If F_ballast(14) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(14) = 0#
    End If
End If

F_ballast(15) = FLOW_Two_Ways(9, 15, 15, 15)
P_ballast(15) = ZWP(3) * 1.1
If ZWP(3) < 0.004 Then              ' The Tank is EMPTY
    If F_ballast(15) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(15) = 0#
    End If
End If

F_ballast(16) = FLOW_Two_Ways(9, 16, 16, 16)
P_ballast(16) = ZWP(2) * 1.1
If ZWP(2) < 0.006 Then              ' The Tank is EMPTY
    If F_ballast(16) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(16) = 0#
    End If
End If

F_ballast(17) = FLOW_Two_Ways(9, 17, 17, 17)
P_ballast(17) = ZWP(1) * 1.1
If ZWP(1) < 0.002 Then              ' The Tank is EMPTY
    If F_ballast(17) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(17) = 0#
    End If
End If


F_ball_tank(4) = F_ballast(14)
F_ball_tank(3) = F_ballast(15)
F_ball_tank(2) = F_ballast(16)
F_ball_tank(1) = F_ballast(17)

For i = 1 To 4
    If ZWP(i) > 1# Then
            ' We have filled the tank to 100% and are now filling the expansion pipe....
            Z_dot(i) = F_ball_tank(i) / 100#
    Else
        If i = 1 Then   '  Tank #1 is smaller
            Z_dot(i) = F_ball_tank(i) / 50000#
        Else
            Z_dot(i) = F_ball_tank(i) / 60000#
        End If
    End If
   
            ZWP(i) = ZWP(i) + Z_dot(i) * TIME_STEP
Next i


' Stbd Ballast Tanks
F_ballast(18) = FLOW_Two_Ways(10, 18, 18, 18)
P_ballast(18) = ZWS(4) * 1.1
If ZWS(4) < 0.005 Then              ' The Tank is EMPTY
    If F_ballast(18) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(18) = 0#
    End If
End If

F_ballast(19) = FLOW_Two_Ways(10, 19, 19, 19)
P_ballast(19) = ZWS(3) * 1.1
If ZWS(3) < 0.003 Then              ' The Tank is EMPTY
    If F_ballast(19) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(19) = 0#
    End If
End If

F_ballast(20) = FLOW_Two_Ways(10, 20, 20, 20)
P_ballast(20) = ZWS(2) * 1.1
If ZWS(2) < 0.005 Then              ' The Tank is EMPTY
    If F_ballast(20) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(20) = 0#
    End If
End If

F_ballast(21) = FLOW_Two_Ways(10, 21, 21, 21)
P_ballast(21) = ZWS(1) * 1.1
If ZWS(1) < 0.002 Then              ' The Tank is EMPTY
    If F_ballast(21) < 0# Then      ' We are SUCKING from an EMPTY TANK so...
       F_ballast(21) = 0#
    End If
End If

F_ball_tank(4) = F_ballast(18)
F_ball_tank(3) = F_ballast(19)
F_ball_tank(2) = F_ballast(20)
F_ball_tank(1) = F_ballast(21)


For i = 1 To 4
    If ZWS(i) > 1# Then
            ' We have filled the tank to 100% and are now filling the expansion pipe....
            Z_dot(i) = F_ball_tank(i) / 100#
    Else
        If i = 1 Then   '  Tank #1 is smaller
            Z_dot(i) = F_ball_tank(i) / 50000#
        Else
            Z_dot(i) = F_ball_tank(i) / 60000#
        End If
    End If
   
            ZWS(i) = ZWS(i) + Z_dot(i) * TIME_STEP
Next i


' Ballast Pressure points (nodes)
P_dot = (F_ballast(2) + F_ballast(30) - F_ballast(39)) / 100#
P_ballast(1) = P_ballast(1) + P_dot * TIME_STEP

P_dot = (F_ballast(8) + F_ballast(10) - F_ballast(40)) / 100#
P_ballast(2) = P_ballast(2) + P_dot * TIME_STEP

P_dot = (F_ballast(23) - F_ballast(30) - F_ballast(27) - F_ballast(24)) / 100#
P_ballast(3) = P_ballast(3) + P_dot * TIME_STEP

P_dot = (F_ballast(25) - F_ballast(10) - F_ballast(28) + F_ballast(24)) / 100#
P_ballast(4) = P_ballast(4) + P_dot * TIME_STEP

P_dot = (F_ballast(39) - F_ballast(5) - F_ballast(4)) / 100#
P_ballast(5) = P_ballast(5) + P_dot * TIME_STEP

P_dot = (F_ballast(40) - F_ballast(13) - F_ballast(12)) / 100#
P_ballast(6) = P_ballast(6) + P_dot * TIME_STEP

P_dot = (F_ballast(4) + F_ballast(13) - F_ballast(23) - F_ballast(25) - F_ballast(29) - F_ballast(33)) / 100#
P_ballast(7) = P_ballast(7) + P_dot * TIME_STEP

P_dot = (F_ballast(5) + F_ballast(12) - F_ballast(26)) / 100#
P_ballast(8) = P_ballast(8) + P_dot * TIME_STEP


F_to_tanks = F_ballast(14) + F_ballast(15) + F_ballast(16) + F_ballast(17)
P_dot = (F_ballast(27) + F_ballast(29) - F_to_tanks - F_ballast(22)) / 100#
P_ballast(9) = P_ballast(9) + P_dot * TIME_STEP


F_to_tanks = F_ballast(18) + F_ballast(19) + F_ballast(20) + F_ballast(21)
P_dot = (F_ballast(28) + F_ballast(33) - F_to_tanks + F_ballast(22)) / 100#
P_ballast(10) = P_ballast(10) + P_dot * TIME_STEP



End Sub
Function FLOW_One_Way(IN_no As Integer, OUT_no As Integer, V_no As Integer, R_no As Integer) As Single

If Open_ballast_valve(V_no) Then
    If P_ballast(IN_no) > P_ballast(OUT_no) Then
        FLOW_One_Way = (P_ballast(IN_no) - P_ballast(OUT_no)) / R_ballast(R_no)
    Else
        FLOW_One_Way = 0#
    End If
Else
        FLOW_One_Way = 0#
End If

End Function
Function FLOW_Two_Ways(IN_no As Integer, OUT_no As Integer, V_no As Integer, R_no As Integer) As Single

If Open_ballast_valve(V_no) Then
        FLOW_Two_Ways = (P_ballast(IN_no) - P_ballast(OUT_no)) / R_ballast(R_no)
Else
        FLOW_Two_Ways = 0#
End If

End Function

