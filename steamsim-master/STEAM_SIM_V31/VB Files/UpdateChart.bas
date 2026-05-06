Attribute VB_Name = "UpdateChart"
' used to return enthalpy values
Option Explicit

Public Sub UpdateGraph()
Dim i As Integer
Dim j As Integer
Dim Ploted As Integer

 'calculate entropy values
 
 CalculateEntropy
 
 CalculateZero
  
 'convert to correct units
 If Not (SIUnits) Then
   SItoBS
 End If
 
 'reset data arrays in graph component
 Mollier.Graph1.DataReset = 1 'reset graph data
 Mollier.Graph1.DataReset = 8 'reset xposdata
 Mollier.Graph1.DataReset = 23 'rest overlay data
 
 'read new data values. Testing each value for range
 For i = 0 To 5
  j = i + 1
  If (Entropy(i) >= MINENTROPY) And (Entropy(i) <= MAXENTROPY) Then
   If (Enthalpy(i) >= MINENTHALPY) And (Enthalpy(i) <= MAXENTHALPY) Then
    'set graph values
    With Mollier.Graph1
     .ThisSet = j
     .ThisPoint = 1
     .XPosData = Entropy(i)  ' x co-ordinate
     .GraphData = Enthalpy(i) ' y co-cordinate
     .ThisPoint = j
     .OverlayXPosData = Entropy(i)
     .OverlayGraphData = Enthalpy(i)
    End With
    
    Ploted = Ploted + 1
   End If
  End If
 Next
  
  ' update grid values
  UpdateGrid
   'redraw graph
  If Ploted <> 6 Then
   Mollier.Graph1.OverlayGraph = 0
  Else
   Mollier.Graph1.OverlayGraph = 1
  End If
 Mollier.Graph1.DrawMode = 3
 End Sub

Private Sub UpdateGrid()
  Dim i As Integer
   For i = 0 To 5
      With Mollier.MSFlexGrid1
          .TextMatrix(GRID_ROW_PRESSURE, i + 1) = Format(Pressure(i), "##0.000")
          .TextMatrix(GRID_ROW_ENTHALPY, i + 1) = Format(Enthalpy(i), "##0")
          
          ' 1/7/99
            If i = 4 Or i = 5 Then
                .TextMatrix(GRID_ROW_TEMPERATURE, i + 1) = Format(Temperature(i), "##0.0")
            Else
                .TextMatrix(GRID_ROW_TEMPERATURE, i + 1) = Format(Temperature(i), "##0")
            End If
            
          .TextMatrix(GRID_ROW_ENTROPY, i + 1) = Format(Entropy(i), "##0.000")
          .TextMatrix(GRID_ROW_DRYNESS, i + 1) = Format(Dryness(i), "##0.000")
      End With
   Next
End Sub

Public Sub ResetGrid()
  Dim i As Integer
   For i = 0 To 5
      With Mollier.MSFlexGrid1
          Pressure(i) = 0
          Entropy(i) = 0
          Enthalpy(i) = 0
          Temperature(i) = 0
          Dryness(i) = 0
      End With
   Next
End Sub

Private Sub CalculateZero()
Dim Press As Single
Dim T1_steam As Single
Dim H As Single

Dim Dry As Single
Dim T1_sat As Single
Dim sf As Single
Dim sfg As Single
Dim hf As Single
Dim hfg As Single
Dim x As Single
  
 If Enthalpy(1) <> 0 Then ' allowes all entries in table to equal zero at start
  Press = 0.047  ' pressure at point zero
  T1_steam = Temperature(5) ' copy temp
  Dry = Dryness(1) ' copy dryness

  Call get_Tsat(Press, T1_sat)
  ' get sf and sfg
  Call get_sf_sfg(Press, sf, sfg)
  Call get_hf_hfg(Press, hf, hfg)
  'get enthropy
  'Call get_Hsteam(Press, T1_sat, s, sf, sfg, T1_steam, Dry, CP_Steam)
  
  'set value of entropy
  Enthalpy(0) = hf + (((Entropy(1) - sf) / sfg) * hfg)
  Entropy(0) = Entropy(1)
  Pressure(0) = Press
  Dryness(0) = Dry
  Temperature(0) = T1_steam
 End If
End Sub

 Private Sub CalculateEntropy()
Dim Index As Integer
Dim Press As Single
Dim T1_steam As Single
Dim s As Single
Dim CP_Steam As Single

Dim Dry As Single
Dim T1_sat As Single
Dim sf As Single
Dim sfg As Single

Dim hf_out As Single
Dim hfg_out As Single
Dim Tout As Single
Dim qDry_out As Single
Dim H_enthalpy As Single

 ' calculate entropy
 ' 1/7/99
'For index = 0 To 5
 For Index = 1 To 5
 
  Press = Pressure(Index)  ' copy pressure
  
    ' 1/7/99
   ' T1_steam = Temperature(index) ' copy temp
  
  
    ' 1/7/99 To try to cure LP Bleed Entropy jumping ( synch problem with model  )
    '  Works OK....
    Call get_Tsat(Press, T1_sat)
    Call get_hf_hfg(Press, hf_out, hfg_out)
    H_enthalpy = Enthalpy(Index)
    
    'Call get_Tsteam(P1_steam, T1_sat, H1_steam, hf_1, hfg_1, T1_steam, q1_dryness, CP_steam)
     Call get_Tsteam(Press, T1_sat, H_enthalpy, hf_out, hfg_out, Tout, qDry_out, CP_Steam)
     T1_steam = Tout ' copy temp

  
  Dry = Dryness(Index) ' copy dryness

  Call get_Tsat(Press, T1_sat)
  ' get sf and sfg
  Call get_sf_sfg(Press, sf, sfg)
  'get enthropy
  Call get_ssteam(Press, T1_sat, s, sf, sfg, T1_steam, Dry, CP_Steam)
  'set value of entropy
  
 ' Dry_s = Dry
 ' Dry_h = Dryness(5)
  
 ' Mollier.List1.Clear
 ' Mollier.List1.AddItem Format(Dry_s, "#.000")
 ' Mollier.List1.AddItem Format(Dry_h, "#.000")
  
  Dryness(Index) = Dry
  
  Entropy(Index) = s
  
 Next
 
End Sub

Private Sub SItoBS()
'convert si units to bs units
Dim i As Integer
  For i = 0 To 5
    Enthalpy(i) = Enthalpy(i) * 0.4299
    Entropy(i) = Entropy(i) * 0.2388
    
    ' 20/7/99
    Temperature(i) = Temperature(i) * 1.8 + 32#  ' deg C to deg F
    Pressure(i) = Pressure(i) * 14.502      ' Bar to psi
    
    
  Next
End Sub
'not used
Private Sub BStoSI()
' convert bs units to si units
Dim i As Integer
  For i = 0 To 5
    Enthalpy(i) = Enthalpy(i) / 0.4299
    Entropy(i) = Entropy(i) / 0.2388
  Next
End Sub

