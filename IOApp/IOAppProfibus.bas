Attribute VB_Name = "IOAppProfibus"
Option Explicit

' Profibus read/write helpers, copied and adapted from the simulator's IO.bas.
' Stripped of:
'   - the simulator-only "If ProfibusDisabled" early-out (this app is exactly
'     the case where Profibus IS enabled),
'   - the HTTP routines (Request_IOIO_Data, Send_IOIO_Outputs) - those live
'     in IOAppNet.bas and operate in the inverse direction,
'   - the orchestration (ReadAllCards) - the form's Timer1_Timer drives the
'     loop directly with per-card gating.
'
' On hardware error, helpers MsgBox via ErrorMessage (same UX the simulator
' had) and return False. The form's timer logs and skips that card on the
' next tick.

Public Sub ErrorMessage(Status As Integer)
    Dim msg As String
    Select Case Status
        Case 32
            msg = "Function Parameters Incorrect"
        Case 36
            msg = "Invalid WinBloc Address"
        Case 45
            msg = "Applicom dialogue Software Not Resident in Memory" & vbCrLf & "Please Run PCINITIO"
        Case 47
            msg = "Invalid Profi Card Number"
        Case 95
            msg = "Applicom Driver Inaccessible"
        Case Else
            msg = "Profibus error, status=" & Status
    End Select
    MsgBox msg, vbCritical, "I/O App Error"
End Sub

Private Function Writedata(TabByte() As Byte, Card_no As Integer, Block_NO As Integer, Nb As Integer) As Boolean
    Dim Ret_value As Boolean
    Dim Status As Integer
    Dim Offset As Integer

    Offset = 0
    Ret_value = IO_WriteQByte(Card_no, Block_NO, Offset, Nb, TabByte(0), Status)

    If Ret_value = False Then
        ErrorMessage Status
        Writedata = False
        Exit Function
    End If
    Writedata = True
End Function

Private Function WriteRefresh(Card_no As Integer) As Boolean
    Dim Ret_value As Boolean
    Dim Status As Integer

    Ret_value = IO_RefreshOutput(Card_no, Status)
    If Ret_value = False Then
        ErrorMessage Status
        WriteRefresh = False
        Exit Function
    End If
    WriteRefresh = True
End Function

Public Function Set_D_Output(Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
    Dim Block_NO As Integer, Bit As Integer, TAIndex As Integer, base As Integer
    Dim Nb As Integer
    Dim Index As Integer, i As Integer
    Dim Bits(1 To 8) As Byte
    Dim TabByte(3) As Byte
    Dim ok As Boolean

    Set_D_Output = False
    Nb = 4

    While Start_addr <= End_addr
        TAIndex = 0
        If Card = WINCARD_A Then
            Illuminate = Not Illuminate
        Else
            Illuminate2 = Not Illuminate2
        End If
        While TAIndex < 4
            If Card = WINCARD_A Then
                base = (TAIndex * 16) + (Start_addr - D_OUTPUT_START_ADDR) * NO_D_OUTPUT_CHANNELS
            Else
                base = DO_OFFSET + (TAIndex * 16) + (Start_addr - D_OUTPUT_START_ADDR_B) * NO_D_OUTPUT_CHANNELS
            End If
            Bit = 1
            While Bit < 9
                Index = (base + Bit * 2) - 2
                If D_OUTPUT(Index) Then
                    ' even index = flash
                    If Illuminate Then
                        Bits(Bit) = &H0
                    Else
                        Bits(Bit) = &H1 * 2 ^ (Bit - 1)
                    End If
                Else
                    If D_OUTPUT(Index + 1) Then
                        ' odd index = steady
                        Bits(Bit) = &H1 * 2 ^ (Bit - 1)
                    Else
                        Bits(Bit) = &H0
                    End If
                End If
                Bit = Bit + 1
            Wend
            For i = 1 To 8
                TabByte(TAIndex) = TabByte(TAIndex) Or Bits(i)
            Next
            TAIndex = TAIndex + 1
        Wend

        ok = Writedata(TabByte, Card, Start_addr, Nb)
        If ok = False Then Exit Function

        For i = 0 To 3
            TabByte(i) = 0
        Next
        Start_addr = Start_addr + 1
    Wend

    ok = WriteRefresh(Card)
    If ok = False Then Exit Function
    Set_D_Output = True
End Function

Public Function Set_A_Output(Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
    Dim TAIndex As Integer
    Dim base As Integer, Nb As Integer, counter As Integer, i As Integer
    Dim TabByte(8) As Byte
    Dim ok As Boolean

    Nb = 8
    Set_A_Output = False

    While Start_addr <= End_addr
        counter = 0
        For TAIndex = 0 To 7 Step 2
            If Card = WINCARD_A Then
                base = (Start_addr - A_OUTPUT_START_ADDR) * NO_A_OUTPUT_CHANNELS
            Else
                base = AO_OFFSET + (Start_addr - A_OUTPUT_START_ADDR_B) * NO_A_OUTPUT_CHANNELS
            End If

            TabByte(TAIndex) = A_OUTPUT(base + counter)
            ' 12-bit DAC, least-significant 4 bits
            If A_OUTPUT(base + counter) > 192 Then
                TabByte(TAIndex + 1) = 255
            ElseIf A_OUTPUT(base + counter) > 64 Then
                TabByte(TAIndex + 1) = 128
            Else
                TabByte(TAIndex + 1) = 0
            End If
            counter = counter + 1
        Next

        ok = Writedata(TabByte, Card, Start_addr, Nb)
        If ok = False Then Exit Function

        For i = 0 To 3
            TabByte(i) = 0
        Next
        Start_addr = Start_addr + 1
    Wend

    ok = WriteRefresh(Card)
    If ok = False Then Exit Function
    Set_A_Output = True
End Function

Private Function ReadData(TabByte() As Byte, Card As Integer, Block As Integer, Nb As Integer) As Boolean
    Dim Ret_value As Boolean
    Dim Status As Integer

    Ret_value = IO_RefreshInput(Card, Status)
    If Ret_value = False Then
        ErrorMessage Status
        ReadData = False
        Exit Function
    End If

    Ret_value = IO_ReadIByte(Card, Block, 0, Nb, TabByte(0), Status)
    If Ret_value = False Then
        ErrorMessage Status
        ReadData = False
        Exit Function
    End If
    ReadData = True
End Function

Public Function Get_A_Input(Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
    Dim TAIndex As Integer
    Dim base As Integer, Nb As Integer, counter As Integer
    Dim TabByte(8) As Byte
    Dim ok As Boolean

    Nb = 8
    Get_A_Input = False

    While Start_addr <= End_addr
        counter = 0
        ok = ReadData(TabByte, Card, Start_addr, Nb)
        If ok = False Then Exit Function
        For TAIndex = 0 To 7 Step 2
            If Card = WINCARD_A Then
                base = (Start_addr - A_INPUT_START_ADDR) * NO_A_INPUT_CHANNELS
            Else
                base = AI_OFFSET + (Start_addr - A_INPUT_START_ADDR_B) * NO_A_INPUT_CHANNELS
            End If
            A_INPUT(base + counter) = TabByte(TAIndex)
            counter = counter + 1
        Next
        Start_addr = Start_addr + 1
    Wend
    Get_A_Input = True
End Function

Public Function Get_D_input(Card As Integer, Start_addr As Integer, End_addr As Integer) As Boolean
    Dim TAIndex As Integer
    Dim base As Integer, Nb As Integer, Bit As Integer
    Dim TabByte(4) As Byte
    Dim ok As Boolean

    Nb = 4
    Get_D_input = False

    While Start_addr <= End_addr
        ok = ReadData(TabByte, Card, Start_addr, Nb)
        If ok = False Then Exit Function
        If Card = WINCARD_A Then
            base = (Start_addr - D_INPUT_START_ADDR) * NO_D_INPUT_CHANNELS
        Else
            base = DI_OFFSET + (Start_addr - D_INPUT_START_ADDR_B) * NO_D_INPUT_CHANNELS
        End If
        TAIndex = 0
        While TAIndex < 4
            For Bit = 0 To 7
                If (TabByte(TAIndex) And 2 ^ Bit) Then
                    D_INPUT(base + (8 * TAIndex) + Bit) = 1
                Else
                    D_INPUT(base + (8 * TAIndex) + Bit) = 0
                End If
            Next
            TAIndex = TAIndex + 1
        Wend
        Start_addr = Start_addr + 1
    Wend
    Get_D_input = True
End Function
