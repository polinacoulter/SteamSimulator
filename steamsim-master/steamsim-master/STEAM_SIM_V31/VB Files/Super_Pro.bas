Attribute VB_Name = "Super_Pro"
Sub Dongle_check()


' === 17/01/03 ADDED this subroutine....===
Dim mPercentage As Integer

mPercentage = 20


    '===6/8/02 new dongle===
    'Dim Dongle_Response As String
    
    
    Dim Dongle_Password As String
    Dim j As Integer
    Dim Crash_it As Double
    
    Randomize Timer
    j = Rnd * 100
    
    If j <= mPercentage Then 'we pay attention to the response of the dongle
        Dim i As Integer
        
        
        '===5/8/02===
        'Randomize Timer
        'I = Rnd * (UBound(DonglePassword, 1) - LBound(DonglePassword, 1)) 'select one random code from the list
        
        'OPEN_SENT 'send the code selected to the dongle and save its response
        'Dongle_Response = Format$(SSQUERY(ByVal DonglePassword(I), ByVal Len(DonglePassword(I))))
        'CLOSE_SENT
        
        
        
        '===5/8/02 testing superpro dongle===
        Dim Idr As Integer
        Dim CopyResponse As String
    
            Randomize Timer
            i = Int(Rnd(1) * 200) + 1     'select one random code from the list
            
            
            ' === 03/02/03 TESTING  Query = 1E1C , Response = AB11 ===
            'i = 14
            
            ' Get the query value and expected response.
            QueryStr = GetQueryData(i)
            ExpectedResponse = GetResponseData(i)

            QueryLength = (Len(QueryTable(i).Query)) / 2

            ' Query the key
            
           ' === 03/01/03 ===
           'Result% = RNBOsproQuery(MyAPIPacket, maxim100_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
           
           ' === 11/09/03 ===
           'Result% = RNBOsproQuery(MyAPIPacket, maxim700_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
            Result% = RNBOsproQuery(MyAPIPacket, maxim600_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
            
            Dongle_Response = ResponseStr
                                
            
            Dongle_Reading(i) = ""
            For Idr = 0 To QueryLength - 1
                If Len(Hex(ResponseStr.data(Idr))) < 2 Then
                    CopyResponse = "0" & Hex(ResponseStr.data(Idr))
                Else
                    CopyResponse = Hex(ResponseStr.data(Idr))
                End If
                Dongle_Reading(i) = Dongle_Reading(i) & CopyResponse
            Next Idr
            
            
            
            '============== 31/10/02 timeout stuff ===================
            
            Result% = RNBOsproRead(MyAPIPacket, Timeout_period_address, Timeout_period)
            
            If Timeout_period <> 0 Then
            
                System_current_date_LONG = CDate(Date)
            
                If System_current_date_LONG > 30000 Then
                    System_current_date_LONG = System_current_date_LONG - 30000
                End If
                
                System_current_date_INT = System_current_date_LONG
                            
                If System_current_date_LONG < SSPRO_current_date Or System_current_date_LONG - SSPRO_start_date > Timeout_period Then
                    
                    ' Decrement the counter to de-activate Maxim100...(Now maxim700...)
                   ' === 03/02/03 ===
                   'Result% = RNBOsproDecrement(MyAPIPacket, SSPRO_Write_Password, maxim100_MemAddress - 1)
                    
                    ' === 11/09/03 ===
                    'Result% = RNBOsproDecrement(MyAPIPacket, SSPRO_Write_Password, maxim700_MemAddress - 1)
                     Result% = RNBOsproDecrement(MyAPIPacket, SSPRO_Write_Password, maxim600_MemAddress - 1)
                    
                    ' Set the Timeout_period to zero (i.e. make it a normal key)
                    Result% = RNBOsproWrite(MyAPIPacket, SSPRO_Write_Password, Timeout_period_address, 0, 0)
                    
                    
                    '=== 03/02/03 ===
                    TRIAL_HAS_ENDED = True
                    Trial_frm.Show
                
                Else
                
                    If Write_count = 0 Then
                    
                            
                       ' === 04/11/02 DO NOT WRITE to the EEPROM in the dongle un-necessarily =============
                        Result% = RNBOsproRead(MyAPIPacket, SSPRO_current_date_address, SSPRO_current_date)
                        If SSPRO_current_date < System_current_date_INT Then
                        
                            ' Up-date the SSPRO_current_date....
                            SSPRO_access_code = 0
                            Result% = RNBOsproWrite(MyAPIPacket, SSPRO_Write_Password, SSPRO_current_date_address, System_current_date_INT, SSPRO_access_code)
                            
                            Write_count = 1200
                            
                        End If
                       ' ==================================================================================
                        
                    Else
                            Write_count = Write_count - 1
                    End If
                    
                    
                
                End If
            End If
            
            '============== 31/10/02 timeout stuff end ===============
            
            
            
        '===5/8/02 end===
        
        
        
        
         ' testing j = 1 / 0
         
        '=== 25/04/02 ===
        'If Dongle_Response <> DongleResponse(i) Then
        '    RaiseEvent DongleMissing
        'End If
        
        '===5/8/02===
        'Dongle_Reading(I) = Dongle_Response
                    
        
    Else 'we do not pay attention to the dongle response
        'Dim Nbr_Char As Integer
        'Dongle_Password = ""
        'Do
        '    Randomize Timer
        '    Nbr_Char = Rnd * 64 'make a random number of character between 3 and 64 characters to be sure that the dongle will send a response
        'Loop While Nbr_Char < 3
       '
        'For i% = 0 To Nbr_Char Step 1
        '    Randomize Timer
        '    Dongle_Password = Dongle_Password & Chr$(33 + Rnd * 93) 'make the string with printable ascii chraracters
        'Next i%
        
        '===5/8/02 superpro dongle testing===
        'OPEN_SENT ' Open the dongle port
        'Format$ (SSQUERY(ByVal Dongle_Password, ByVal Len(Dongle_Password)))
        'CLOSE_SENT ' Close the dongle port
        
        
        Randomize Timer
        'i = Int(Rnd(1) * 200) + 1     'select one random code from the list
        

        'Make random words for this pass

        QueryLength = Int(10 * Rnd) + 2 'make a random number between 2 and 11, eqn is ((Upper - Lower +1)*rnd+lower


        ' === 03/02/03 ===
       'For j = 1 To QueryLength
        For j = 0 To QueryLength - 1
            QueryStr.data(j) = Int(256 * Rnd)
        Next j
        
        
        
        
        'QueryStr = GetQueryData(i)
        
       ' === 03/02/03 ===
       'Result% = RNBOsproQuery(MyAPIPacket, maxim100_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
        
        ' === 11/09/03 ===
       'Result% = RNBOsproQuery(MyAPIPacket, maxim700_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
        Result% = RNBOsproQuery(MyAPIPacket, maxim600_MemAddress, QueryStr, ResponseStr, unused, QueryLength)
        
        
        
    End If

End Sub
