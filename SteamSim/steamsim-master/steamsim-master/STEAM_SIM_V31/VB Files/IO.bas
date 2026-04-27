Attribute VB_Name = "IO"
Option Explicit
'Global CardStatus As Integer


Public MyIOIOReadyStateHandler As IOIOReadyStateHandler


Private Function ProfibusDisabled() As Boolean
  ProfibusDisabled = DISABLE_PROFIBUS
End Function


Public Sub ErrorMessage(Status As Integer)
Dim msg As String
  Select Case Status
    Case 32
      msg = "Function Paramaters Incorrect"
    Case 36
      msg = "Invalid WinBloc Address"
    Case 45
      msg = "Applicom dialogue Software Not Resident in Memory" & vbCrLf & "Please Run PCINITIO"
    Case 47
      msg = "Invalid Profi Card Number"
    Case 95
      msg = "Applicom Driver Inacessable"
  End Select
  MsgBox msg, vbCritical, "Error"
End Sub
Private Function Writedata(TabByte() As Byte, CLOCK As Timer, Card_no As Integer, Block_NO As Integer, Nb As Integer) As Boolean
'writes data to blocks
Dim Ret_value As Boolean
Dim Status As Integer, i As Integer, Offset As Integer
If ProfibusDisabled Then
  Writedata = True
  Exit Function
End If
Offset = 0
' write tabbyte array to buffer
  Ret_value = IO_WriteQByte(Card_no, Block_NO, Offset, Nb, TabByte(0), Status)
  
  If Ret_value = False Then ' Testing for errors
    CLOCK.Enabled = False
    'MsgBox "Error Writing Too Profibus, Status  :" & Status, vbCritical, "Error"
    ErrorMessage Status
    Writedata = False ' flag error
    Exit Function
  End If
  Writedata = True
End Function
Private Function WriteRefresh(CLOCK As Timer, Card_no As Integer) As Boolean
Dim Ret_value As Boolean
Dim Status As Integer
If ProfibusDisabled Then
  WriteRefresh = True
  Exit Function
End If
' refresh buffer
  Ret_value = IO_RefreshOutput(Card_no, Status)
  If Ret_value = False Then  ' test if refresh occured
    CLOCK.Enabled = False
    'MsgBox "Error Refreshing Profibus, Status :" & Status, vbCritical, "Error"
    ErrorMessage Status
    WriteRefresh = False
    Exit Function
  End If
  WriteRefresh = True
End Function


Public Function Set_D_Output(CLOCK As Timer, Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
'prepares data to be written to 32do
Dim Block_NO  As Integer, Bit As Integer, TAIndex As Integer, base As Integer
Dim Nb As Integer
Dim index As Integer, i As Integer
Dim Bits(1 To 8) As Byte
Dim TabByte(3) As Byte
Dim error As Boolean
'Static illuminate As Boolean
'Static illuminate2 As Boolean
If ProfibusDisabled Then
  Set_D_Output = True
  Exit Function
End If
Set_D_Output = False
' Prepare profibus for writing too
Nb = 4  ' number of bytes to be written



' only for testing
'Start_addr = 35
' end test

While Start_addr <= End_addr 'D_OUTPUT_END_ADDR
  TAIndex = 0
  If Card = WINCARD_A Then
      Illuminate = Not Illuminate
  Else
      Illuminate2 = Not Illuminate2
  End If
  While TAIndex < 4
    'Calculated base address of block in do array
    If Card = WINCARD_A Then
      base = (TAIndex * 16) + (Start_addr - D_OUTPUT_START_ADDR) * NO_D_OUTPUT_CHANNELS  'offset to first channel of block

    Else
      base = DO_OFFSET + (TAIndex * 16) + (Start_addr - D_OUTPUT_START_ADDR_B) * NO_D_OUTPUT_CHANNELS  'offset to first channel of block
      
    End If
    Bit = 1 ' start at bit one
    While Bit < 9
      index = (base + Bit * 2) - 2  ' positon in digital output array
      If D_OUTPUT(index) Then  ' test if light is flashing
        If Illuminate Then
          ' switch off lamp
          Bits(Bit) = &H0 ' set bit to zero. no need to shift as i*0 = 0
        Else
          ' switch on lamp
          Bits(Bit) = &H1 * 2 ^ (Bit - 1) 'set left most bit and left shift by bit
        End If
      Else
        If D_OUTPUT(index + 1) Then 'test if light is constant
          Bits(Bit) = &H1 * 2 ^ (Bit - 1) 'set left most bit and left shift by bit
        Else
        'light is off
          Bits(Bit) = &H0 ' set bit to zero. no need to shift as i*0 = 0
        End If
      End If
    
      Bit = Bit + 1 ' increment bit position
    Wend
    'pack bit into a byte
    For i = 1 To 8
      TabByte(TAIndex) = TabByte(TAIndex) Or Bits(i)
    Next
    TAIndex = TAIndex + 1 ' increment tabbyte array position
  Wend
  ' write to buffer

  error = Writedata(TabByte, CLOCK, Card, Start_addr, Nb)
  
  'test for error
  If error = False Then Exit Function
  'reset tab byte array
  For i = 0 To 3
    TabByte(i) = 0
  Next
  Start_addr = Start_addr + 1
Wend

'write data to blocks
error = WriteRefresh(CLOCK, Card)

'test for error
If error = False Then Exit Function
'reverse lamp status
'illuminate = Not illuminate
Set_D_Output = True
End Function

Public Function Set_A_Output(CLOCK As Timer, Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
'prepares data to be written to 4AO
Dim TAIndex As Integer
Dim base As Integer, Nb As Integer, counter As Integer, i As Integer
Dim TabByte(8) As Byte
Dim error As Boolean
If ProfibusDisabled Then
  Set_A_Output = True
  Exit Function
End If
Nb = 8  ' number of bytes to be written

Set_A_Output = False
' testing
'Start_addr = 15
'end testing

While Start_addr <= End_addr  'Block_NO_End
  counter = 0
  For TAIndex = 0 To 7 Step 2
    
    If Card = WINCARD_A Then
      base = (Start_addr - A_OUTPUT_START_ADDR) * NO_A_OUTPUT_CHANNELS ' base index of array
    Else
      base = AO_OFFSET + (Start_addr - A_OUTPUT_START_ADDR_B) * NO_A_OUTPUT_CHANNELS ' base index of array
    End If
    
    TabByte(TAIndex) = A_OUTPUT(base + counter)
    ' 4AO use 12 bit d to a converter. least sig 4 bits  to
    If A_OUTPUT(base + counter) > 192 Then
      TabByte(TAIndex + 1) = 255 'full
    Else
      If A_OUTPUT(base + counter) > 64 Then
        TabByte(TAIndex + 1) = 128 ' half
      Else
        TabByte(TAIndex + 1) = 0 ' zero
      End If
    End If
    counter = counter + 1
  Next
    
  'write to buffer
  error = Writedata(TabByte, CLOCK, Card, Start_addr, Nb)
  
  If error = False Then Exit Function
  
  'reset tab byte array
  For i = 0 To 3
    TabByte(i) = 0
  Next
  Start_addr = Start_addr + 1
Wend
4
  'write data to blocks
  error = WriteRefresh(CLOCK, Card)

  If error = False Then Exit Function
  Set_A_Output = True
End Function

Public Function ReadData(TabByte() As Byte, CLOCK As Timer, Card As Integer, Block As Integer, Nb As Integer) As Boolean
'reads data from blocks
Dim Ret_value As Boolean
Dim Status As Integer
If ProfibusDisabled Then
  ReadData = True
  Exit Function
End If

  'refresh inputs buffer
  Ret_value = IO_RefreshInput(Card, Status)
  If Ret_value = False Then
    CLOCK.Enabled = False
    'MsgBox "Error Refreshing Profibus, Status  :" & Status, vbCritical, "Error"
    ErrorMessage Status
    ReadData = False
    Exit Function
  End If

  ' read values from buffer
  Ret_value = IO_ReadIByte(Card, Block, 0, Nb, TabByte(0), Status)
  ' check for errors
  If Ret_value = False Then
    CLOCK.Enabled = False
    'MsgBox "Error Reading Profibus, Status  :" & Status, vbCritical, "Error"
    ErrorMessage Status
    ReadData = False
    Exit Function
  End If
  ReadData = True
End Function

Public Function Get_A_Input(CLOCK As Timer, Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
'Processes data read from 4AI blocks
Dim TAIndex As Integer
Dim base As Integer, Nb As Integer, counter As Integer
Dim TabByte(8) As Byte
Dim error As Boolean
If ProfibusDisabled Then
  Get_A_Input = True
  Exit Function
End If
  Nb = 8  'number of bytes read from block
  
Get_A_Input = False
'testing only
'Start_addr = 50
' end test
  While Start_addr <= End_addr 'Block_NO_End
    counter = 0
    'read input buffer
    error = ReadData(TabByte, CLOCK, Card, Start_addr, Nb)
    If error = False Then Exit Function
    For TAIndex = 0 To 7 Step 2
      If Card = WINCARD_A Then
        base = (Start_addr - A_INPUT_START_ADDR) * NO_A_INPUT_CHANNELS ' base index of array
      Else
        base = AI_OFFSET + (Start_addr - A_INPUT_START_ADDR_B) * NO_A_INPUT_CHANNELS ' base index of array
      End If
      Dim checkString As String
      Dim index As Integer
      index = base + counter
      checkString = "ain" & index & ","
      If (InStr(profiIgnoreList, checkString) = 0) Then             'If the value exists in the string, don't write to the array
        A_INPUT(base + counter) = TabByte(TAIndex) 'copy values
      End If
      counter = counter + 1
    Next
    Start_addr = Start_addr + 1
  Wend
Get_A_Input = True
End Function

Public Function Get_D_input(CLOCK As Timer, Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
'Processess data read from 32di blocks
Dim TAIndex As Integer
Dim base As Integer, Nb As Integer, Bit As Integer
Dim TabByte(4) As Byte
Dim error As Boolean
If ProfibusDisabled Then
  Get_D_input = True
  Exit Function
End If
 Get_D_input = False
  Nb = 4  ' number of bytes to read
'testing only
'Start_addr = 54
'end test
   While Start_addr <= End_addr 'Block_NO_End
    error = ReadData(TabByte, CLOCK, Card, Start_addr, Nb) 'read data from buffer
    If error = False Then Exit Function
    If Card = WINCARD_A Then
      base = (Start_addr - D_INPUT_START_ADDR) * NO_D_INPUT_CHANNELS ' address in array of first channel on a block
    Else
      base = DI_OFFSET + (Start_addr - D_INPUT_START_ADDR_B) * NO_D_INPUT_CHANNELS ' address in array of first channel on a block
    End If
    TAIndex = 0 ' index for read data array
    While TAIndex < 4
      For Bit = 0 To 7
        If (TabByte(TAIndex) And 2 ^ Bit) Then
          D_INPUT(base + (8 * TAIndex) + Bit) = True
        Else
          D_INPUT(base + (8 * TAIndex) + Bit) = False
        End If
      Next
      TAIndex = TAIndex + 1
    Wend
    Start_addr = Start_addr + 1
  Wend
  Get_D_input = True
End Function

Private Sub Request_IOIO_Data()
    ' Requires a reference to Microsoft XML library

    ' Drop any previous request unconditionally. The old gate (If ioioXmlHttp Is Nothing)
    ' could permanently latch closed if a callback ever failed to nil the handle, which
    ' would silently freeze polling. Abandoning the prior request is safe over localhost.
    On Error Resume Next
    Set ioioXmlHttp = Nothing
    On Error GoTo 0

    Set ioioXmlHttp = New MSXML2.XMLHTTP

    If MyIOIOReadyStateHandler Is Nothing Then          ' Create only one instance of the wrapper class.
        Set MyIOIOReadyStateHandler = New IOIOReadyStateHandler
    End If

    ' MSXML2.XMLHTTP shares the WinINet cache (same one IE uses). Without a unique
    ' URL per request plus no-cache headers, identical GETs after the first 200 are
    ' served from cache and never reach the network — polling silently freezes.
    ' A short ?_=N counter alone wasn't enough on XP — WinINet sometimes ignores
    ' simple query strings for cache keying. Combining Timer (seconds-since-midnight
    ' as a Single, e.g. 45123.234) with a per-call counter gives every URL
    ' millisecond-level entropy that no cache layer can collapse.
    Static cacheBuster As Long
    cacheBuster = cacheBuster + 1

    Dim cacheBusterUrl As String
    cacheBusterUrl = IOIO_STATUS_URL & "?t=" & CStr(Timer) & "&n=" & CStr(cacheBuster)

    On Error Resume Next
    ioioXmlHttp.Open "GET", cacheBusterUrl, True
    ioioXmlHttp.setRequestHeader "Cache-Control", "no-cache"
    ioioXmlHttp.setRequestHeader "Pragma", "no-cache"
    ioioXmlHttp.OnReadyStateChange = MyIOIOReadyStateHandler        ' Assign the wrapper class object to onreadystatechange.
    ioioXmlHttp.send
    On Error GoTo 0

End Sub

Public Sub Send_IOIO_Inputs()
    ' One-shot push of the snapshot's expected input state to the Python server.
    ' Called from the snapshot-load path so the server's STATE["ain"]/STATE["din"]
    ' reflect "what hardware should look like" at the moment the snapshot was taken.
    ' Subsequent slider drags in the browser then create real, meaningful mismatches
    ' on the SnapCompare page.
    On Error Resume Next
    Set ioioInHttp = Nothing
    On Error GoTo 0

    ' Pull the expected values from Fortran into the Reload arrays.
    Call get_AI_Reloads(VB_AI_Reload(0))
    Call get_DI_Reloads(VB_DI_Reload(0))

    Set ioioInHttp = New MSXML2.XMLHTTP

    Dim ainParts() As String
    ReDim ainParts(0 To UBound(VB_AI_Reload))
    Dim i As Long
    For i = 0 To UBound(VB_AI_Reload)
        ainParts(i) = CStr(VB_AI_Reload(i))
    Next i

    Dim dinParts() As String
    ReDim dinParts(0 To UBound(VB_DI_Reload))
    For i = 0 To UBound(VB_DI_Reload)
        dinParts(i) = CStr(VB_DI_Reload(i))
    Next i

    Dim body As String
    body = "ain=" & Join(ainParts, ",") & "&din=" & Join(dinParts, ",")

    On Error Resume Next
    ioioInHttp.Open "POST", IOIO_INPUTS_URL, True
    ioioInHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    ioioInHttp.send body
    On Error GoTo 0

End Sub

Private Sub Send_IOIO_Outputs()
    ' Push A_OUTPUT and D_OUTPUT to the Python server. Fire-and-forget async POST: we
    ' don't care about the response, and abandoning a previous in-flight request is
    ' safe — same self-healing pattern as Request_IOIO_Data.
    On Error Resume Next
    Set ioioOutHttp = Nothing
    On Error GoTo 0

    Set ioioOutHttp = New MSXML2.XMLHTTP

    Dim aoutParts() As String
    ReDim aoutParts(0 To UBound(A_OUTPUT))
    Dim i As Long
    For i = 0 To UBound(A_OUTPUT)
        aoutParts(i) = CStr(A_OUTPUT(i))
    Next i

    Dim doutParts() As String
    ReDim doutParts(0 To UBound(D_OUTPUT))
    For i = 0 To UBound(D_OUTPUT)
        doutParts(i) = CStr(D_OUTPUT(i))
    Next i

    Dim body As String
    body = "aout=" & Join(aoutParts, ",") & "&dout=" & Join(doutParts, ",")

    On Error Resume Next
    ioioOutHttp.Open "POST", IOIO_OUTPUTS_URL, True
    ioioOutHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    ioioOutHttp.send body
    On Error GoTo 0

End Sub

Private Sub CheckInputsToIgnore()
If (Not profiIgnoreListRead) Then
    profiIgnoreListRead = True
    Dim MyLine As String
    Open "C:\Steam_Sim\AnalogInputsToIgnore.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, MyLine
            Dim index As Integer
            Dim commaPosition As Integer
            commaPosition = InStr(MyLine, ",")
            If (commaPosition = 0) Then
                'No Comma means we are of format ain18 or ain123
                index = CInt(Val(Mid$(MyLine, 4)))
            Else
                index = CInt(Val(Mid$(MyLine, 4, commaPosition - 4)))
            End If
            
            profiIgnoreList = profiIgnoreList & "ain" & index & ","
    
            MsgBox MyLine
        Loop
    Close #1
End If
End Sub

Public Function ReadAllCards(CLOCK As Timer) As Boolean
'reads all applicom cards and returns true if successfull else false
Dim error As Boolean
If ProfibusDisabled Then
  Request_IOIO_Data
  Send_IOIO_Outputs
  ReadAllCards = True
  Exit Function
End If
ReadAllCards = False
'do
'CardStatus = 0
If Set_D_Output(CLOCK, WINCARD_A, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR) = False Then Exit Function
'CardStatus = 1
If Set_D_Output(CLOCK, WINCARD_B, D_OUTPUT_START_ADDR_B, D_OUTPUT_END_ADDR_B) = False Then Exit Function
'CardStatus = 2
'ao
'GoTo dibit
If Set_A_Output(CLOCK, WINCARD_A, A_OUTPUT_START_ADDR, A_OUTPUT_END_ADDR) = False Then Exit Function
'CardStatus = 21
If Set_A_Output(CLOCK, WINCARD_B, A_OUTPUT_START_ADDR_B, A_OUTPUT_END_ADDR_B) = False Then Exit Function
'CardStatus = 22
'dibit:
'di
If Get_D_input(CLOCK, WINCARD_A, D_INPUT_START_ADDR, D_INPUT_END_ADDR) = False Then Exit Function
'CardStatus = 31
If Get_D_input(CLOCK, WINCARD_B, D_INPUT_START_ADDR_B, D_INPUT_END_ADDR_B) = False Then Exit Function
'CardStatus = 32

'This is the new IOIO-based connections for the handful of analog inputs
Request_IOIO_Data

'ai
If Get_A_Input(CLOCK, WINCARD_A, A_INPUT_START_ADDR, A_INPUT_END_ADDR) = False Then Exit Function
'CardStatus = 41
If Get_A_Input(CLOCK, WINCARD_B, A_INPUT_START_ADDR_B, A_INPUT_END_ADDR_B) = False Then Exit Function
'CardStatus = 42
ReadAllCards = True

End Function
