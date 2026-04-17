Attribute VB_Name = "HardwareTests"
Option Explicit


Public Function GetChannelNum(BlockNum As Integer, Card As Integer, Start_addr_A As Integer, _
                              End_addr_A As Integer, Start_addr As Integer) As Integer
Dim temp As Integer

  If Card = WINCARD_A Then
    temp = ((BlockNum - Start_addr_A)) * NO_D_OUTPUT_CHANNELS
  Else  'wincard B
    temp = DO_OFFSET + ((BlockNum - Start_addr) * NO_D_OUTPUT_CHANNELS)
  End If
  If temp < 0 Then temp = 0
  GetChannelNum = temp
End Function

Public Function GetBlockNum(Channel As Integer, Card As Integer, Module As String) As Integer
If Module = "DO" Then
  If Card = WINCARD_A Then
    GetBlockNum = (Channel \ NO_D_OUTPUT_CHANNELS) + D_OUTPUT_START_ADDR
  Else  'wincard_B
    GetBlockNum = ((Channel - DO_OFFSET) \ NO_D_OUTPUT_CHANNELS) + D_OUTPUT_START_ADDR_B
  End If
Else
  If Card = WINCARD_A Then
    GetBlockNum = (Channel \ NO_D_INPUT_CHANNELS) + D_INPUT_START_ADDR
  Else  'wincard_B
    GetBlockNum = ((Channel - DI_OFFSET) \ NO_D_INPUT_CHANNELS) + D_INPUT_START_ADDR_B
  End If
End If
End Function

Public Sub DITest(Card As Integer)
'loops through digital inputs, reporting the
'active switch
Dim Index As Integer, max As Integer, i As Integer
Dim DIStr As String
Dim DiStatus As String

'Dim DIDataBase As New ADODB.Recordset
'Dim SQL As String
If Card = 1 Then
  Index = 0
  max = DI_OFFSET - 1
Else
  If Card = 2 Then
    Index = DI_OFFSET
    max = REALADDR_DI 'NO_D_INPUT - 1
  Else 'both cards
    Index = 0
    max = REALADDR_DI
  End If
End If
        
'ReDim Changes(index To max)
  While Index <= max
    'If (D_INPUT(index) <> Changes(index)) And D_INPUT(index) Then
    If (D_INPUT(Index) <> Changes(Index)) Then
      FrmHardwareTest.TBBlockNumDI = Format(GetBlockNum(Index, Card, "DI"), "000")
      FrmHardwareTest.TBChannelNumDI = Format(Index, "000")
      '11/11/99 no database
      'DIDataBase.MoveFirst
      'DIDataBase.Find "[i/o ref] =" & index
      'If Not (DIDataBase.EOF) Then
        'FrmHardwareTest.TBSwitchDes = DIDataBase(1)
      'Else
        'FrmHardwareTest.TBSwitchDes = "Unknown Switch"
      'End If
      '17/11/99 Call GetHardIOText("DI", Index, FrmHardwareTest.TBSwitchDes.text)
      Call GetHardIOText("DI", Index, DIStr)
      If D_INPUT(Index) Then
        DiStatus = "TRUE"
      Else
        DiStatus = "FALSE"
      End If
      
      FrmHardwareTest.TBSwitchDes.text = DiStatus & Space(2) & DIStr
      
    End If
    Changes(Index) = D_INPUT(Index)
    Index = Index + 1
  Wend
End Sub

Public Sub AITest(Card As Integer)
'loops through analogue inputs, reporting values
Dim Index As Integer, max As Integer
Dim SQL As String
Dim AIStr As String * 40
Dim IOText As String
Dim TempStr As String
Dim i As Integer

If Not AITestFirstPass Then 'get AI text in
    For Index = 0 To 255
        Select Case Index
        Case 0
            AIStr = "Throttle Local Pointer Position"
            'this is a cheat because direct access file can't hold data in record 0
            'so can't hold text for ch 0. All others OK.
     
        'Case 1 To 19, 40 To 41, 44 To 46
         Case 1 To 19, 40 To 41, 44 To 47
           Call GetHardIOText("AI", Index, IOText)
            IOText = RTrim(IOText)
            AIStr = Space(40)
            If Len(IOText) <= 40 Then
                Mid$(AIStr, 1, Len(IOText)) = IOText
            Else
                AIStr = Mid$(IOText, 1, 40)
            End If
        End Select
        
        AITextString(Index) = AIStr
        AITestFirstPass = True
        
    Next
End If

If FrmHardwareTest.Check1.Value Then    'show as %
    FrmHardwareTest.Label10.Caption = "Value %"
Else
    FrmHardwareTest.Label10.Caption = "Value 0-255"
End If


Index = 0
max = REALADDR_AI



' === 12/12/01 14.107 Display FLICKERS because the list is Cleared EVERY PASS  so ...===
'    FrmHardwareTest.List1.Clear ' clear list box
    
  
While Index <= max
  
        Select Case Index
    
        '  Do NOT skip the Synchro-scope
        '  Case 0 To 19, 40 To 41, 44 To 46
        Case 0 To 19, 40 To 41, 44 To 47
     
            If Index = 46 Then
                Index = Index
            End If
            
            AIStr = AITextString(Index)

            ' 15/11/99 Added Test statement
            A_value = A_INPUT(Index)
            old_A_value = OLD_A_input(Index)
        
            TempStr = AIStr & Space(4) & Format(Index, "000") & Space(6)
       
            If (Abs(old_A_value - A_value) > 5) Then
            
             ' === 12/12/01 14.107  Display flickers so do this...===
             
                   AI_update_required = True
                
                ' === 12/12/01 ===
                '  If FrmHardwareTest.Check1.Value Then    'show as %
                '      FrmHardwareTest.List1.AddItem TempStr & Format((OLD_A_input(Index) * 100 / 255), "000") & Space(6) & Format((A_INPUT(Index) * 100 / 255), "000")
                '  Else
                '      FrmHardwareTest.List1.AddItem TempStr & Format(OLD_A_input(Index), "000") & Space(6) & Format(A_INPUT(Index), "000")
                '  End If
                               
            Else
            
                ' === 12/12/01 14.107 ===
                ' If FrmHardwareTest.Check1.Value Then 'show as %
                '     FrmHardwareTest.List1.AddItem TempStr & Format((A_INPUT(Index) * 100 / 255), "000")
                ' Else
                '     FrmHardwareTest.List1.AddItem TempStr & Format(A_INPUT(Index), "000")
                ' End If
                
            End If
            
        End Select

        Index = Index + 1
    
Wend

' === 12/12/01 14.107 Display flickers so do this...===
If AI_update_required Then

   AI_update_required = False
   
    FrmHardwareTest.List1.clear ' clear list box
    
    ' === 12/12/01 While and Wend DIDN'T WORK so let's do it proporly...===
    'While Index <= max
     For Index = 0 To max
     
        Select Case Index
    
        '  Do NOT skip the Synchro-scope
        '  Case 0 To 19, 40 To 41, 44 To 46
        Case 0 To 19, 40 To 41, 44 To 47
     
            If Index = 46 Then
                Index = Index
            End If
            
            AIStr = AITextString(Index)

            ' 15/11/99 Added Test statement
            A_value = A_INPUT(Index)
            old_A_value = OLD_A_input(Index)
        
            TempStr = AIStr & Space(4) & Format(Index, "000") & Space(6)
       
            If (Abs(old_A_value - A_value) > 5) Then
                                                        
                    If FrmHardwareTest.Check1.Value Then    'show as %
                        FrmHardwareTest.List1.AddItem TempStr & Format((OLD_A_input(Index) * 100 / 255), "000") & Space(6) & Format((A_INPUT(Index) * 100 / 255), "000")
                    Else
                        FrmHardwareTest.List1.AddItem TempStr & Format(OLD_A_input(Index), "000") & Space(6) & Format(A_INPUT(Index), "000")
                    End If
                    
                    ' === 12/12/01 14.107 ADDED: ===
                    OLD_A_input(Index) = A_INPUT(Index)

                    
            Else
            
                If FrmHardwareTest.Check1.Value Then 'show as %
                    FrmHardwareTest.List1.AddItem TempStr & Format((A_INPUT(Index) * 100 / 255), "000")
                Else
                    FrmHardwareTest.List1.AddItem TempStr & Format(A_INPUT(Index), "000")
                End If
                
            End If
            
        End Select

    '    Index = Index + 1
    'Wend
    
    Next Index
    
    
End If
   
  
End Sub

