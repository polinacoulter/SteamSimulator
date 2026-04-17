Attribute VB_Name = "SET_Model_Constants"
Sub SET_Constants()
Dim i As Integer
Dim IN_string As String
Dim Real_number As Single

' === 23/08/02 =======
On Error GoTo ErrH

FileNo = FreeFile
Open "c:\STEAM_SIM\text\CONSTANTS.txt" For Input As #FileNo

    Input #FileNo, IN_string
    Shaft_Inertia_multiplier = Val(IN_string)
    i = 1
    Call set_CONSTANT(i, Shaft_Inertia_multiplier)
    
    Input #FileNo, IN_string
    Ship_Inertia_multiplier = Val(IN_string)
    i = 2
    Call set_CONSTANT(i, Ship_Inertia_multiplier)
    
    Input #FileNo, IN_string
    Pitch_multiplier = Val(IN_string)
    i = 3
    Call set_CONSTANT(i, Pitch_multiplier)
    
    ' === 24/09/02 17.056 ===
    Input #FileNo, IN_string
    White_Smoke_multiplier = Val(IN_string)
    i = 4
    Call set_CONSTANT(i, White_Smoke_multiplier)
    
    
    ' === 12/12/02 19.605 ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 5
    Call set_CONSTANT(i, Real_number)
    
    ' === 12/12/02 19.605 ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 6
    Call set_CONSTANT(i, Real_number)
    
    
    ' ===16/12/02 CONSTANT(7) = Ship Resistive Multiplier===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 7
    Call set_CONSTANT(i, Real_number)
    
    ' ===17/12/02 CONSTANT(8) = Ship Resistive Multiplier ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 8
    Call set_CONSTANT(i, Real_number)
    
    
    ' ===17/12/02 CONSTANT(9) = ADT Tank Size Multiplier ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 9
    Call set_CONSTANT(i, Real_number)
   
    ' ===31/03/03 CONSTANT(10) = FO Low Pressure Boiler trip point     (350.0 psi) ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 10
    Call set_CONSTANT(i, Real_number)
    ' ===31/03/03 CONSTANT(11) = FO Low Pressure stand-by cut in point (390.0 psi) ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 11
    Call set_CONSTANT(i, Real_number)
  
  
    ' ===15/07/09 CONSTANT(12) = Stbd Boiler Burner # 1 Diesel Oil Burn Rate Multiplier  (0.1 to 1.0) ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 12
    Call set_CONSTANT(i, Real_number)
  
    ' ===15/07/09 CONSTANT(13) = Port Boiler Burner # 1 Diesel Oil Burn Rate Multiplier  (0.1 to 1.0) ===
    Input #FileNo, IN_string
    Real_number = Val(IN_string)
    i = 13
    Call set_CONSTANT(i, Real_number)
  
  
  
Close #FileNo

Exit Sub

ErrH:
resp = MsgBox("There was an error reading c:\STEAM_SIM\text\CONSTANTS.txt" & Chr(10) & _
       "You must ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
       Stop

End Sub

