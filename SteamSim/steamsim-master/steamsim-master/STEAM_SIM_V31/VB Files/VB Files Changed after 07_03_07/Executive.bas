Attribute VB_Name = "Exec"


' === 14/03/01 ===
Global This_PC_Name As String

' === 13/03/01 ===
' There are "VIRTUAL PUMPS" between DI's 3000 and 3199 ( See Breaker X-ref V15.xls )
Global VIRTUAL_PUMP_TEXT(200) As String


'14/12/00
Global old_d_run As Byte
Global Explosion_count As Integer


'11/12/00
Global I_AM_LLANELLI As Boolean


'08/09/00
Global ResetAllAlarms As Boolean
'Global StrSpeed As String


' JYA 12/6/98 Dongle variables and DLL declarations
' Beginning
    
 ' 18/6/98
  ' Global Const Dongle_Enable_Flag As Byte = False ' IMPORTANT: Set to True before creating the .exe file
  ' 9/2/99
  ' Global Const Dongle_Enable_Flag As Byte = True  ' IMPORTANT: Set to True before creating the .exe file
  
   Global AITestFirstPass
   Global AITextString(0 To 255) As String * 50
    Global HardDIs(200) As Integer
    Global HardAIs(100) As Integer
    Global DIReloadText(200) As String * 100
    Global AIReloadText(200) As String * 100
    Global DIDifferences As Boolean
    Global AIDifferences As Boolean
    Global DIDiffIndex As Integer
    Global AIDiffIndex As Integer
    Global OldDIDiffIndex As Integer
    Global OldAIDiffIndex As Integer
    
    Global SetPointMult(0 To 3) As Integer
    
    Global Dongle_Enable_Flag '  **** NOW SET in frm_intro.frm  ****
    
    
    Global Const Dongle_Password = "malopropism" ' Password sent to the dongle
    '18/6/98 Global Const Dongle_Password_Response = "62624" ' Correct reply from the dongle from the above password
    Global Const Dongle_Password_Response As Integer = -2912 ' Correct reply from the dongle from the above password
 
 ' === 21/12/01 ===
    Global Dongle_Integer As Integer
    
' === 09/11/03 For the OLD BASBH Dongle... ===
   'Global Dongle_Response As Integer
    Global Dongle_Response_Integer As Integer
    
    Global Exit_sharply As Boolean
    
    
    
 ' 23/6/98 FORCE to C:\HAVEN to allow operation from VB5 development mode
 '   Declare Function OPEN_SENT Lib "Sp32w.dll" () As Integer
 '   Declare Function CLOSE_SENT Lib "Sp32w.dll" () As Integer
 '   Declare Function SSQUERY Lib "Sp32w.dll" (ByVal querystr$, ByVal querylen%) As Integer
    
    Declare Function OPEN_SENT Lib "c:\HAVEN\Sp32w.dll" () As Integer
    Declare Function CLOSE_SENT Lib "c:\HAVEN\Sp32w.dll" () As Integer
    Declare Function SSQUERY Lib "c:\HAVEN\Sp32w.dll" (ByVal QueryStr$, ByVal querylen%) As Integer



' End

' 24/2/98 For now
Global FaultEntryflag As Byte

'Misc
Global ClearText As String
Global next_form As Integer
Global InstructorName As String
Global ClassName As String
Global CourseName As String
Global StudentName As String
Global Sound_test_on As Integer
'27/11/99 Global sndVolume(0 To 5) As Integer ' Volumes for the sounds
Global sndVolume(0 To 1, 31) As Integer ' Volumes for the sounds on PCs 0 and 1
Global sndFreq(0 To 1, 31) As Double ' Frequencies for the sounds on PCs 0 and 1
Global SndEnabled(0 To 1, 31) As Boolean
Global DummyPCNumber As Integer

Global ComOpen As Byte
Global ComOpen_LSS As Byte

' 10/4/98 RWJ is using this in MAIN.BAS.....
'Global reset_flag As Integer

'Don't Know yet
Global hardwareAddress As Integer ' Base address of the Amplicon Board
Global hardwareIRQ As Integer ' IRQ of the Amplicon Board

'Logging
'Global log_strings(30) As String
Global LogFileNumber As Integer ' Filenumber of the log file
Global LogFileName As String ' And corresponding file name
Global LogFileOpen As Byte 'flag to remember if logging is on
Global LogLevel As Integer ' The level of logging (1..maybe 3?)
Global ViewLogFileNumber As Integer
Global LogString As String
Global CurrentLogString As String
Global LogViewOn As Byte

'23/11/99
Global LogTitleStr As String

Global Const LogTitle As Integer = 0
Global Const LogActions As Integer = 1
Global Const LogFaults As Integer = 2
Global Const LogChangeAlarmSP As Integer = 3
Global Const LogPCDIs As Integer = 4
Global Const LogAlarms As Integer = 5
Global Const LogVars As Integer = 6
Global Const LogSnaps As Integer = 7
'25/10/99
Global Const LogHardDIs As Integer = 8
Global Const LogHardAIs As Integer = 9
Global Const PCAIs As Integer = 10
Global LoggingFirstPass As Boolean
Global DIDummy(0 To 3071) As Boolean
Global AIDummy(0 To 255) As Byte


Global OldAlarmTime As Date


'SI / ANSI
Global LastUnitType As Boolean
Global SI_units As Boolean

' 19/10/99
Global Advanced_mode As Boolean


'Language

Global useEnglish As Byte
Global Language_Str(2000) As String
Global Main_strings(65) As String
Global LanguagePrefix As String * 1

'IO Test
Global DiONString As String
Global DiOFFString As String
Global DiString As String
Global DiStatus As String
Global DINumberChanged As Byte
Global ONEvent As Byte
Global AItest_strings(40) As String
Global DItest_strings(30) As String

'Faults
Global Fault_Types(64) As String
Global Faultext(1000) As String
Global ThisFaultnumber As Integer
Global FaultXRef(2000) As Integer
Global NumberOfFaults As Integer


' 14/10/99
Global FaultUpdateRequired As Boolean
Global OldNumberOffaults As Integer


Global Fault(2000) As Byte
Global FaultScheduleIndex As Integer
Global FaultSchedule(100, 3) As Integer ' 1st element is fault nmumber, second is time in mins before activatuion,third is time left
Global ModelTimerCounter
Global fault_time(2000) As Date
Global FaultGroup(100, 2) As Integer 'stores the first and last line number relating to the fault text for each fault group.
Global BiggestFaultNumber As Integer
Option Compare Text

'Actions
Global Action_Types(64) As String
Global ActionText(1000) As String
Global ThisActionNumber As Integer
Global ActionXRef(1400) As Integer
Global NumberOfActions As Integer
Global Action(2, 1400) As Byte
Global ActionGroup(100, 2) As Integer 'stores the first and last line number relating to the fault text for each fault group.
Global BiggestActionNumber As Integer
Global Action_time(1400) As String * 16
Global ActionEntryFlag As Byte
Global ClearedByStudent As Byte


'Plot Variables
Global Var_Types(64) As String
Global VarText(1000) As String
Global ThisVarNumber As Integer
Global VarXRef(1400) As Integer
Global NumberOfVars As Integer
Global VarGroup(100, 2) As Integer 'stores the first and last line number relating to the fault text for each fault group.
Global BiggestVarNumber As Integer
Global Var_time(8) As String * 16
Global VarEntryFlag As Byte
' now in exec.bas Global LogVars As Integer
Global PlotVar(8) As Integer
Global VarChanges As Byte


Global plot_Reals(16) As Single
Global ThickLine(8) As Integer
Global OldThickLine(8) As Integer
Global LegendText(8) As String
Global GridStyle As Integer
Global UnloadPlot As Byte


'Engine Dynamic Plots
Global Const pi As Single = 3.1416
Global PVLookUp(1 To 16) As Single
Global PCLookUp(1 To 16) As Single
Global OldY(0 To 16) As Single
Global Xm(0 To 31) As Single
Global Ym(0 To 31) As Single


Global PistonSpeedOn As Byte
Global Stroke_CrankOn As Byte
Global P_VOn As Byte
Global P_CrankOn As Byte
Global DrawCardOn As Byte
Global dP_CrankOn As Byte
Global dHeat_Crank As Byte
Global NOxOn As Byte


Global angle As Integer
Global throw As Single
Global bore As Single
Global stroke As Single
Global PistonOffset As Single
'Global stroke As Single
Global CylLength As Single
Global LengthConRod As Single
Global Position As Single
Global engrevslider As Integer
Global engrevs As Single
Global PistSpeed(360) As Single
Global Pos1 As Single
Global Pos2 As Single
Global OneDegTime As Single
Global resp As Integer
Global AngInRad As Single
Global r As Single
Global ArcSin_X As Single
Global L As Single
Global Lc As Single

Global Points(0 To 10000) As Single
Global Vc1 As Single   'Initial Vol, compression stroke
Global Vc As Single    'Calculated Vol, compression stroke
Global Vc2 As Single   'Final Vol, compression stroke

Global Vx1 As Single   'Initial Vol, exhaust stroke
Global Vx As Single    'Calculated Vol, exhaust stroke
Global Vx2 As Single   'Final Vol, exhaust stroke

Global Pc1 As Single   'Initial Pressure, compression stroke
Global Pc As Single    'Calculated Pressure, compression stroke
Global Pc2 As Single   'Final Pressure, compression stroke

Global Px1 As Single   'Initial Pressure, exhaust stroke
Global Px As Single    'Calculated Pressure, exhaust stroke
Global Px2 As Single   'Final Pressure, exhaust stroke

Global K1 As Single     'calculated Constant (compression) = P*(V**Gamma)
Global K2 As Single     'calculated Constant (power)= P*(V**Gamma)

Global C_STEPs As Integer    'number of STEPs compression stroke - this is configurable

Global Gamma As Single     'nominally 1.4
Global NoSTEPsPerUnitVol As Integer
Global start As Integer
Global finish As Integer
Global STEPs As Integer
Global MaxPCompression As Single
Global LastAngle As Integer
Global LastPressure As Single
Global MaxEngRevs As Single
Global NVc As Single
Global MaxVol As Single
Global MinVol As Single
Global SweptVol As Single
Global CompRatio As Single
Global UnitOfVol As Single
Global MaxSpeed As Single
Global maxPressure As Single
Global ScavAirPressure As Single
Global InjectionAngle As Single
Global IgnitionDelay As Single
Global IgnitionAngle As Single
Global IgnitionVar As Single
Global MaxPAngle As Single
Global ExhVVOpenAngle As Single
Global ExhVVDecayRate As Single
Global FuelInjPeriod As Single
Global CetaneVal As Single
Global LinerTemp As Single
Global CylWallTemp As Single
Global PistonCrownTemp As Single
Global DiamInjParticle As Single
Global SprayAngle As Single
Global ScavEfficiency As Single
Global NOx As Single
Global NomLinerTemp As Single
Global NomCylWallTemp As Single
Global NomPistonCrownTemp As Single

' 19/1/99
Global T_exhaust As Single
Global Diam_constant As Single

Global SpeedMult As Single
Global XPos(1000) As Single
Global LastUnit As Single
Global NumberOfPoints As Integer
Global NumberOfNewPoints As Single
Global YaxisOffset As Single

Global EditPVLU As Byte

Global MinPressProportion As Single
Global ScavInletRiseRate As Single
Global ConstScavPAngle As Single
Global ScavPDecayRate As Single
Global AirValveOpenVol As Single
Global AmbientAirPressure As Single

Global CylNumber As Integer
Global ScavT As Single
Global ScavH As Single
Global Fudge1 As Single
Global CylVariation As Single

'Global counter As Integer
Global ix As Single
Global incr As Single
Global AirValveOpenAngle As Integer
Global TwoStroke As Byte




'Network & Groups
Global Number_of_PCs As Integer
Global PC_Group(0 To 31) As Integer
Global my_host_id(0 To 31) As Integer
Global ImageBeenDragged As Byte
Global OldOnLineStatus As Byte
Global NewOnLineStatus As Byte
Global OnLineAllowed As Byte
Global OffLineAllowed As Byte
Global OldOnLineCount As Integer


' === 06/03/07 ===
'Global HostSockTxConnected(16) As Byte
Global HostSockTxConnected(32) As Byte

Global ClientSockTxConnected As Byte
Global Max_Host_sock_Tx As Integer ' sockets 0..Max_Host_sock_Tx
Global Max_Host_sock_Rx As Integer
Global Max_Client_Sock_Tx As Integer
Global Max_Client_Sock_Rx As Integer




'Snapshot
Global TakeSnapFlag As Byte
Global ElecExec As Byte
Global Const Rec_Length = 200


'TCP/IP & Intro Page
Global FirstTcpipInitCounter As Long
Global FirstTcpipInit As Byte
Global Const Max_Socket = 1

' 13/3/98
Global client_counter(Max_Socket) As Integer

'Global Const Port_0 = 2000
'Global Const Port_1 = 2001

Global Const Array_Size = 200
' 17/3/98
Global Const Real_Array_Size = 1000


Global RxConnected As Byte
Global IntroTimerCounter As Integer
Global IntroCounter As Integer
Global AcceptButtonPressed As Byte
Global HostAvailable As Byte
Global HOST As String
Global Const HOST_A = "1.1.1.2"
Global Const HOST_B = "1.1.1.3"

'20/12/01 Won't work anymore ?
'Global Const HostAddrPC00 = "1.1.1.0"
'Global Const HostAddrPC01 = "1.1.1.1"
'Global Const HostAddrPC02 = "1.1.1.2"
'Global Const HostAddrPC03 = "1.1.1.3"
'Global Const HostAddrPC04 = "1.1.1.4"
'Global Const HostAddrPC05 = "1.1.1.5"

' === 25/03/02 ===
'Global Const HostAddrPC00 = "192.168.1.100"
'Global Const HostAddrPC01 = "192.168.1.1"
'Global Const HostAddrPC02 = "192.168.1.2"
'Global Const HostAddrPC03 = "192.168.1.3"
'Global Const HostAddrPC04 = "192.168.1.4"
'Global Const HostAddrPC05 = "192.168.1.5"

Global HostAddrPC00 As String
Global HostAddrPC01 As String
Global HostAddrPC02 As String
Global HostAddrPC03 As String
Global HostAddrPC04 As String
Global HostAddrPC05 As String

' === 06/03/07 ===
Global HostAddrPC06 As String
Global HostAddrPC07 As String
Global HostAddrPC08 As String
Global HostAddrPC09 As String
Global HostAddrPC10 As String
Global HostAddrPC11 As String
Global HostAddrPC12 As String
Global HostAddrPC13 As String
Global HostAddrPC14 As String
Global HostAddrPC15 As String


' === 25/03/02 END ===



Global Const Host_Tx_Base = 2000
'Global Const SocketTxBasePortAddr = 2100
Global Const Host_Rx_Base = 2100
Global ExecTimerCounter As Double

'General Configuration
Global PTT As Boolean




' 19/3/98
'Global Rx_PC_reals(200) As Single
Global Rx_PC_reals(2000) As Single

Global MessageString As String
Global HostId As String
Global TempStr1 As String
Global TempStr2 As String
' 23/3/98
Global TempStr3 As String
Global TempStr4 As String

'Password
Global instructor_password As String
Global PasswordSet As Byte

'15/06/07 add wav for Gen Alarm sound
Global Gen_Alarm_Bell_flag As Boolean







Sub SaveDefaultSoundMixdata()
Dim FileNo As Integer
Dim i As Integer

On Error GoTo ErrH

FileNo = FreeFile(0)

Open "c:\STEAM_SIM\Sounds\SoundConfig.dat" For Output As #FileNo

        
'04/04/07 must copy slider to variable first - before storing

'For i = 1 To 20
'    sndFreq(0, i) = SoundMix.scrSpeed0(i).Value * 100
'    sndVolume(0, i) = SoundMix.scrVolume0(i).Value
'Next
        
       
'10/12/00 For i = 1 To 8
For i = 1 To 31
    Print #FileNo, sndFreq(0, i)
    Print #FileNo, sndFreq(1, i)
    Print #FileNo, sndVolume(0, i) 'save as a positive number and invert when loading
    Print #FileNo, sndVolume(1, i)
Next

'Input #FileNo, PC_name

'This_PC_number = Val(Mid$(PC_name, Len(PC_name) - 1, 2))


Close #FileNo

Exit Sub

ErrH:
resp = MsgBox("There was an error writing to C:\Steam_Sim\SoundConfig.dat" & Chr(10) & _
       "Ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
Close #FileNo
End Sub

Sub LoadDefaultSoundMixData()
Dim FileNo As Integer
Dim i As Integer
Dim MyStr As String


On Error GoTo ErrH

FileNo = FreeFile(0)

Open "c:\STEAM_SIM\Sounds\SoundConfig.dat" For Input As #FileNo
        
'10/12/00 For i = 1 To 8
For i = 1 To 31
    Input #FileNo, MyStr
    sndFreq(0, i) = Val(MyStr)
    Input #FileNo, MyStr
    sndFreq(1, i) = Val(MyStr)
    Input #FileNo, MyStr
    sndVolume(0, i) = Val(MyStr)
    Input #FileNo, MyStr
    sndVolume(1, i) = Val(MyStr)
Next

Close #FileNo
'Call InitializeSliders

Exit Sub

ErrH:

Debug.Print "Error = " & error$

resp = MsgBox("There was an error reading C:\Steam_Sim\SoundConfig.dat" & Chr(10) & _
       "Ensure this file is correct before proceeding", vbCritical + vbOKOnly, "Initialization Error")
Close #FileNo
End Sub
Sub LogDIAIs()
Dim i As Integer
Dim DITextStr As String
Dim AITextStr As String
Dim DiStatus As String * 6
Dim AI_Old As Integer
Dim AI_New As Integer
Dim AiDiff As Integer


'On Error GoTo LogDIAIErr

If Not LoggingFirstPass Then 'set dummy array = actual for first pass

    For i = 1 To UBound(DIDummy)
        DIDummy(i) = D_INPUT(i)
    Next
    For i = 0 To UBound(A_INPUT)
        AIDummy(i) = A_INPUT(i)
    Next
    LoggingFirstPass = True
End If


'Now check for any Digital changes and put them in di_changes array.
DiCount = 0
For i = 1 To UBound(DIDummy)
    'If DIDummy(i) <> D_INPUT(i) Then
    If DIDummy(i) <> D_INPUT(i) And D_INPUT(i) Then
      DiCount = DiCount + 1
      DI_Changes(DiCount) = i             'store the di number
      DI_States(DiCount) = D_INPUT(i)     'store the di status
      'DIDummy(i) = D_INPUT(i)
    End If
    DIDummy(i) = D_INPUT(i)
Next
    
If DiCount > 0 Then
    For i = 1 To DiCount
        '11/11/99 no database
        'DIDataBase.MoveFirst
        'DIDataBase.Find "[i/o ref] =" & DI_Changes(i)
        If DI_States(i) Then
            DiStatus = "ON"
        Else
            DiStatus = "OFF"
        End If
        
        'Get the text
        '11/11/99 no database
        'If Not (DIDataBase.EOF) Then
            'DITextStr = DIDataBase(1) & "       DI(" & DI_Changes(i) & ")"
        'Else
            'DITextStr = "Unknown Switch at IO Ref " & DI_Changes(i)
        'End If
        Call GetHardIOText("DI", DI_Changes(i), DITextStr)
        DITextStr = DITextStr & "       DI(" & DI_Changes(i) & ")"
        Call Logging(DITextStr, LogHardDIs)
    Next
End If

'Now check for any Analog changes and put them in ai_changes array.

AiCount = 0
'22/11/99  For i = 0 To UBound(A_INPUT)
For i = 0 To 46
    
    'Select Case i
    'Case 0 To 19, 40 To 41, 44 To 46
      If AIChanging(i) Then
        AiOld(i) = AiNew(i)     'if its still changing, store last read value as ref
      Else
        AiOld(i) = AIDummy(i)   'otherwise store orig value
      End If
      AiNew(i) = A_INPUT(i)

      AiDiff = AiNew(i) - AiOld(i)
      If Abs(AiDiff) > 4 Then  ' pot is changing
        AIChanged(i) = True
        AIChanging(i) = True
        'AiCount = AiCount + 1
        'AI_Changes(AiCount) = i             'store the ai number changing
      'AIDummy(i) = A_INPUT(i)
      Else
        AIChanging(i) = False
      End If
  'End Select
  If i = 0 Then
    i = 0
  End If
  

Next

'if AiCount > 0 Then
    'For i = 1 To AiCount
    For i = 0 To UBound(A_INPUT)
          If i = 0 Then
             i = 0
       End If

        
        If AIChanged(i) And Not AIChanging(i) Then 'then it has just stopped changing so log it
            '11/11/99 no database
            'AIDataBase.MoveFirst
            'AIDataBase.Find "[i/o ref] =" & i
        
            'Get the text
            'If Not (AIDataBase.EOF) Then
                'AITextStr = AIDataBase(1) & " changed from " & Format(AIDummy(i) * 0.392, "##0") _
                '& "% to " & Format(AiNew(i) * 0.392, "##0") & "%    AI(" & i & ")"
            'Else
                'AITextStr = "Unknown Analog Input Change on Channel # " & i
            'End If
            
            Call GetHardIOText("AI", i, AITextStr)
            
            ' === 07/06/02 ===
            If i = 0 Then
                ' It is the THROTTLE POT which is scaled -10 to +10
                If AiNew(i) > 133 Then
                    AITextStr = AITextStr _
                    & Format((AiNew(i) - 127) * 10 / 127, "##0.0") & " Ahead"
                ElseIf AiNew(i) > 120 Then
                    AITextStr = AITextStr _
                    & "    Stop"
                Else
                    AITextStr = AITextStr _
                    & Format(Abs(127 - AiNew(i)) * 10 / 127, "##0.0") & " Astern"
                End If
                
            ElseIf i = 1 Then
                ' It is the TELEGRAPH POT which is scaled FULL AHEAD to FULL ASTERN
                ' This is set in GetHardIOText so do NOWT...
            Else
            
                AITextStr = AITextStr & " changed from " & Format(AIDummy(i) * 0.392, "##0") _
                & "% to " & Format(AiNew(i) * 0.392, "##0") & "%    AI(" & i & ")"
                
            End If
             
                
            Call Logging(AITextStr, LogHardAIs)
            
            AIChanged(i) = False
            
            AIDummy(i) = AiNew(i)
            
        End If
    Next
'End If


Exit Sub



LogDIAIErr:

Call Logging("Error finding Switch Changes", LogHardDIs)

   
    



End Sub


Sub Logging(LogString As String, LogType As Integer)
Dim TimeStr As String


'Global Const LogTitle As Integer = 0
'Global Const LogActions As Integer = 1
'Global Const LogFaults As Integer = 2
'Global Const LogChangeAlarmSP As Integer = 3
'Global Const LogPCDIs As Integer = 4
'Global Const LogAlarms As Integer = 5
'Global Const LogVars As Integer = 6

'24/2/98 RWJ  added for now so I can make an EXE file

Select Case LogType

Case 0
    TimeStr = "" 'don't add time for title
Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
    TimeStr = Format(Time, "hh:mm:ss") & "   "
    '4/4/00 LogString = TimeStr & LogString & Chr(10)
    'LogString = TimeStr & LogString

 
End Select
 
'23/11/99 LogString = TimeStr & LogString & Chr(10)


If I_AM_HOST Or Not In_a_group Then
    If LogFileOpen Then
        '4/4/00 add CR for print but NOT for message as it screws up transmission
        '4/4/00 this routine used by client & host. Only add time & CR if HOST
        LogString = TimeStr & LogString & Chr(10)

        Print #LogFileNumber, LogString
        '23/11/99 try putting new entries at top CurrentLogString = CurrentLogString & LogString  'maintain mirror of log
        CurrentLogString = LogString & CurrentLogString    'maintain mirror of log
        If LogViewOn Then 'update log being viewed
            'frm_LogView.RichTextBox1.text = frm_LogView.RichTextBox1.text & LogString
            frm_LogView.RichTextBox1.text = LogTitleStr & CurrentLogString
        End If
    
    End If

Else
   ' Send the Log String to the HOST PC via a SOCKET, for DIs, Actions and Set Points only   '
    '4/4/00 MessageString = "#SET_PCDI#" & Format(nPCDI, "00000") & "#"

    If LogType = LogPCDIs Or LogType = LogActions Or LogType = LogChangeAlarmSP Then
    
        ' === 06/06/02 ===
        LogString = Replace(LogString, "#", "$")

        MessageString = "#SET_LOG#" & LogString & "#"
        
        Call frm_Host.SendMessage(MessageString)
        Debug.Print "Sending Log Message = " & MessageString
    End If
End If






End Sub



Sub GetLanguageStrings(ModuleName As String, ErrorFlag As Byte)
Dim i As Integer
Dim Fnum As Integer
Dim MyChar As String * 1
Dim MyWord As String

' 20/1/99
Dim jj As Integer

ErrorFlag = False
'On Error GoTo ErrorHandler

If useEnglish Then
    'Filepathname = App.Path & "\text\E_" & ModuleName    Use this version once App.Path is sensible
    '7/10/99 Filepathname = "C:\HAVEN\text\E_" & ModuleName
     Filepathname = "c:\steam_sim\text\E_" & ModuleName

     Debug.Print "about to open: " & Filepathname & "  OK"

Else
    'Filepathname = App.Path & "\text\" & LanguagePrefix & "_" & ModuleName Use this version once App.Path is sensible
     '7/10/99 Filepathname = "c:\HAVEN\text\" & LanguagePrefix & "_" & ModuleName
      Filepathname = "c:\steam_sim\text\" & LanguagePrefix & "_" & ModuleName

End If
'Open Filepathname For Input As #1
'Open Filepathname For Input As #Fnum

'Debug.Print "opened " & Filepathname & "  OK"

i = 1
'25/2/98 Problem with Input # is it will not read commas etc and so text gets corrupted
'Do While Not EOF(Fnum)
   ' Input #Fnum, Language_Str(i)
   ' i = i + 1
'Loop
'Close #Fnum

Fnum = FreeFile(0)
Open Filepathname For Input As #Fnum
i = 1
Do While Not EOF(Fnum)
    MyChar = " "
    MyWord = " "
    
    Do While Asc(MyChar) <> 10 And Not EOF(Fnum) ' Loop until CR
        MyChar = Input(1, #Fnum)   ' Get one character.
        'Debug.Print "MyChar = " & MyChar
        If (Asc(MyChar) > 31) Then MyWord = MyWord & MyChar
    
    Loop
    
    'Debug.Print " MyWord = " & MyWord ' Print to Debug window.
    Language_Str(i) = Trim(MyWord)

    'Language_Str(i) = MyWord
    i = i + 1
Loop

' ==== 20/1/99 ====
For jj = i To 2000
   Language_Str(jj) = ""
Next jj
' =================


Close #Fnum
'Debug.Print "closed " & Filepathname & "  OK"

Exit Sub


errorhandler:

Response = MsgBox("Error Reading The Text File: " & Filepathname & Chr(10) & Chr(13) _
& "Correct Error and Reload Program", vbOK + vbCritical, "Error Message")
Close #1
'Debug.Print "closed " & Filepathname & " In ERROR HANDLER"
ErrorFlag = True

Exit Sub


End Sub
Sub GetIOData(ErrorFlag As Byte)
Dim i As Integer
Dim j As Integer
Dim Fnum As Integer
Dim MyChar As String * 1
Dim MyWord As String
Dim MyNumber As Integer
Dim checkit As Boolean
Dim ValveNo As Integer


Dim DI_number As Integer
Dim DO_number As Integer

Dim Break_point_number As Integer

Dim VirtualValveText As String
Dim Fnum2 As Integer
Dim MyString As String




Dim VV_NEW_line(1 To 14) As Integer

' 20/1/99
Dim jj As Integer

ErrorFlag = False
'14/9/99
'Filepathname = "C:\STEAM_SIM\text\" & ModuleName
Filepathname = "C:\STEAM_SIM\text\Breakers.dat"

Fnum = FreeFile(0)
Open Filepathname For Input As #Fnum

i = 1
   
     Do While Not EOF(Fnum) ' Loop until CR
        For j = 1 To 18
          Input #Fnum, MyNumber
          breakers(i, j) = MyNumber
        Next
        i = i + 1
     Loop

Close #Fnum

' 4/9/99   Send breakers( )  to Fortran DLL...

Call set_Breakers(breakers(1, 1))



'14/9/99 - valves
'Filepathname = "C:\STEAM_SIM\text\valves.dat"
Filepathname = "C:\STEAM_SIM\text\ValvesAndText.dat"

Fnum = FreeFile(0)
Open Filepathname For Input As #Fnum
'jbf
Fnum2 = FreeFile(0)
Open "C:\steam_sim\text\VirtualValveText.dat" For Random Access Write As #Fnum2 Len = Rec_Length
     
        'For record_num = 1 To 20
           ' Mid(InfoString, 2, Len(str(record_num))) = str(record_num)
            'Put #Fnum2, record_num, InfoString
        'Next
'Close #Fnum2


' 13/10/99
For i = 1 To UBound(VB_Valves)
  For j = 1 To 14
          VB_Valves(i, j) = 0
  Next j
Next i

'  14/10/99   Dummy DI's ( for 600 valves ) range from 3200 to (3200 + 1200) = 4400
DI_number = 3200
'  14/10/99   Dummy D)'s ( for 600 valves ) range from 3300 to (3300 + 2400) = 5700
DO_number = 3300



i = 1
   
     Do While Not EOF(Fnum) ' Loop until CR
     
        checkit = False
        
        For j = 1 To 14
          Input #Fnum, MyNumber
            
            ' 13/10/99
                VV_NEW_line(j) = MyNumber
                
        Next
        'jbf
        Input #Fnum, MyString 'read text after all data
        
        
        ValveNo = VV_NEW_line(1)
        
        ' === 07/03/07 TESTING ===
        If ValveNo = 487 Then
            i = 1 ' BREAK
        End If
        ' === 07/03/07 TESTING ===
       
        If ValveNo > 0 Then
        
          If VB_Valves(ValveNo, 1) = ValveNo Then
          
            checkit = False
            For j = 1 To 14
                If VB_Valves(ValveNo, j) <> VV_NEW_line(j) Then
                    checkit = True
                End If
            Next
            If checkit = True Then
                'Debug.Print "Incompatible VV data for Valve # " & VV_NEW_line(1)
              Response = MsgBox("Incompatible VV data for Valve # " & VV_NEW_line(1), vbOK + vbCritical, "Error Message")

            End If
            
          Else
          
            If (ValveNo < 1 Or ValveNo > 600) Then
              Response = MsgBox("Error ValvNo<0 or > 600 ", vbOK + vbCritical, "Error Message")
            Else
            
            
              '
              '  GoTo DI_DO_test:
              
              
                  ' 14/10/99
                  If (VV_NEW_line(2) > 3199) Then
                                                    
                            ' MUST BE A VIRTUAL VALVE so assign DUMMY addresses:
                            
                              VB_Valves(ValveNo, 1) = VV_NEW_line(1)
                          
                      'If (DI_number = 4405) Then
                      '   xxxx = 1
                      'End If
                      ' === 20/08/02   See TELEG.F90 which uses DI 4401 to 4403 ===
                       If (DI_number > 4390) And (DI_number < 4410) Then
                            DI_number = 4410
                       End If
                       
                          
                              DI_number = DI_number + 1
                              VB_Valves(ValveNo, 2) = DI_number
                              Put #Fnum2, DI_number, MyString & " OPEN"
                              
                   'If (DI_number = 4405) Then
                   '      xxxx = 1
                   ' End If
                             
                              DI_number = DI_number + 1
                              VB_Valves(ValveNo, 3) = DI_number
                              Put #Fnum2, DI_number, MyString & " CLOSE"
                              
                    'If (DI_number = 4405) Then
                    '     xxxx = 1
                    'End If

                              
                              DI_number = DI_number + 1
                              VB_Valves(ValveNo, 4) = DI_number
                              Put #Fnum2, DI_number, MyString & " AUTO"
                              
                              
                    ' If (DI_number = 4405) Then
                    '    xxxx = 1
                    ' End If

                        ' === 20/08/02   See TELEG.F90 which uses DO 5702 to 5709
                        ' and also AVR_AUTO_LAMP(1) to (4) which use DO 5711 to 5719  ===
                       If (DO_number > 5690) And (DO_number < 5740) Then
                            DO_number = 5740
                       End If
                          
                              DO_number = DO_number + 2
                              VB_Valves(ValveNo, 5) = DO_number
                              
                              DO_number = DO_number + 2
                              VB_Valves(ValveNo, 6) = DO_number
                              
                              DO_number = DO_number + 2
                              VB_Valves(ValveNo, 7) = DO_number
                              
                              
                          For j = 8 To 14
                      
                              VB_Valves(ValveNo, j) = VV_NEW_line(j)
                                                   
                          Next
                          
                      If (VV_NEW_line(1) = 212) Then   ' VV 212 = Burner #1 STBD HAND valve
                          ' 14/10/99 Insert a BREAK-POINT on the next line to de-bug
                        Break_point_number = DO_number
                      End If
                      
                           
                     ' === 20/08/02 ===
                     'If (DI_number > 4400) Then    ' ( for 600 valves )
                      If (DI_number > 4999) Then    ' ( for 600 valves Note:  3*600 + 3200 = 5000 )
                      
                        Response = MsgBox(" DI > 4400 ", vbOK + vbCritical, "ERROR")
                        If Response = vbOK Then X = 1# / 0#
                      End If
                      
                     ' === 20/08/02 ===
                     'If (DO_number > 5700) Then   ' ( for 600 valves )
                      If (DO_number > 6899) Then   ' ( for 600 valves Note:  6*600 + 3300 = 6900 )
                        Response = MsgBox(" DO > 5700 ", vbOK + vbCritical, "ERROR")
                        If Response = vbOK Then X = 1# / 0#
                      End If

                  Else
                  
                      For j = 1 To 14
                      
                        ' ORIGINAL PROGRAM:
                          VB_Valves(ValveNo, j) = VV_NEW_line(j)
                                                   
                      Next
                      
                  End If
                 
                  GoTo skip_DI_DO_test:
                      
DI_DO_test:
                      
                      
                            VB_Valves(ValveNo, 1) = VV_NEW_line(1)
                            VB_Valves(ValveNo, 12) = VV_NEW_line(12)
                            VB_Valves(ValveNo, 13) = VV_NEW_line(13)
                    
                      
              ' TEST the Line to see if the DI has already been assigned...
                
                  For j = 2 To 4
                  
                    If (VV_NEW_line(j) > 0 And VV_NEW_line(j) < 5000) Then
                      If (VV_NEW_line(j) = DI_used(VV_NEW_line(j))) Then
                        Response = MsgBox(" DUPLICATE DI " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      Else
                            VB_Valves(ValveNo, j) = VV_NEW_line(j)
                          DI_used(VV_NEW_line(j)) = VV_NEW_line(j)
                      End If
                    Else
                      If (VV_NEW_line(j) <> 0) Then
                        Response = MsgBox(" DI out of bounds " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      End If
                    End If
                    
                  Next j
                  
                  For j = 8 To 9
                  
                    If (VV_NEW_line(j) > 0 And VV_NEW_line(j) < 5000) Then
                      If (VV_NEW_line(j) = DI_used(VV_NEW_line(j))) Then
                        Response = MsgBox(" DUPLICATE DI " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      Else
                            VB_Valves(ValveNo, j) = VV_NEW_line(j)
                          DI_used(VV_NEW_line(j)) = VV_NEW_line(j)
                      End If
                    Else
                      If (VV_NEW_line(j) <> 0) Then
                        Response = MsgBox(" DI out of bounds " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      End If
                    End If
                    
                  Next j
                  
               ' TEST the Line to see if the DI has already been assigned...
                
                  For j = 5 To 7
                  
                    If (VV_NEW_line(j) > 0 And VV_NEW_line(j) < 7000) Then
                      If (VV_NEW_line(j) = DO_used(VV_NEW_line(j))) Then
                        Response = MsgBox(" DUPLICATE DO " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      Else
                            VB_Valves(ValveNo, j) = VV_NEW_line(j)
                          DO_used(VV_NEW_line(j)) = VV_NEW_line(j)
                      End If
                    Else
                      If (VV_NEW_line(j) <> 0) Then
                        Response = MsgBox(" DO out of bounds " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      End If
                    End If
                    
                  Next j
       
                   For j = 11 To 11
                  
                    If (VV_NEW_line(j) > 0 And VV_NEW_line(j) < 7000) Then
                      If (VV_NEW_line(j) = DO_used(VV_NEW_line(j))) Then
                        Response = MsgBox(" DUPLICATE DO " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      Else
                            VB_Valves(ValveNo, j) = VV_NEW_line(j)
                          DO_used(VV_NEW_line(j)) = VV_NEW_line(j)
                      End If
                    Else
                      If (VV_NEW_line(j) <> 0) Then
                        Response = MsgBox(" DO out of bounds " & VV_NEW_line(j), vbOK + vbCritical, "Error Message")
                      End If
                    End If
                    
                  Next j
                  
                  
skip_DI_DO_test:
              
            End If
              
          End If
          
        End If
        

                             
'   ORIGINALLY
'           VB_Valves(i, j) = MyNumber
          
                              
        
        i = i + 1
     Loop

Close #Fnum
Close #Fnum2

Call set_valves(VB_Valves(1, 1))


Exit Sub


errorhandler:

Response = MsgBox("Error Reading The Text File: " & Filepathname & Chr(10) & Chr(13) _
& "Correct Error and Reload Program", vbOK + vbCritical, "Error Message")
Close #1
'Debug.Print "closed " & Filepathname & " In ERROR HANDLER"
ErrorFlag = True

Exit Sub


End Sub
Sub GetHardIOText(IOType As String, IONum As Integer, IOText As String)

Dim fp As Integer
Dim InfoString As String * 100
Dim info As String

On Error GoTo Errhandler

fp = FreeFile(0)

If IOType = "AI" Then
   Open "C:\Steam_Sim\Text\AIText.dat" For Random Access Read As #fp Len = 100
ElseIf IOType = "AO" Then
   Open "C:\Steam_Sim\Text\AOText.dat" For Random Access Read As #fp Len = 100
ElseIf IOType = "DI" Then


  ' === 06/06/02 ===
  ' Open "C:\Steam_Sim\Text\DIText.dat" For Random Access Read As #fp Len = 100
    Open "C:\Steam_Sim\Text\DIText_2002.dat" For Random Access Read As #fp Len = 100
   
   
ElseIf IOType = "DO" Then
   Open "C:\Steam_Sim\Text\DOText.dat" For Random Access Read As #fp Len = 100
Else
    GoTo Errhandler
End If

' === 07/06/02 ===
If IOType = "AI" Then

    If IONum = 0 Then
    
        ' === 07/06/02 AI(0) is the MECD THROTTLE POT ===
        ' === Record 0 of the AIText does not exist so MUST DO THIS INSTEAD..  =============
        IOText = "MECD throttle "
        
    ElseIf IONum = 1 Then
    
        ' === 07/06/02 AI(1) is the MECD TELEGRAPH POT ===
        
        If VB_More_integers(2) > EOT_instructor(1) Then
            IOText = "MECD Telegraph Full Ahead "
        ElseIf VB_More_integers(2) > EOT_instructor(2) Then
            IOText = "MECD Telegraph Half Ahead "
        ElseIf VB_More_integers(2) > EOT_instructor(3) Then
            IOText = "MECD Telegraph Slow Ahead "
        ElseIf VB_More_integers(2) > EOT_instructor(4) Then
            IOText = "MECD Telegraph Dead Slow Ahead "
        ElseIf VB_More_integers(2) > EOT_instructor(5) Then
            IOText = "MECD Telegraph Stand-by "
        ElseIf VB_More_integers(2) > EOT_instructor(6) Then
            IOText = "MECD Telegraph Stop "
        ElseIf VB_More_integers(2) > EOT_instructor(7) Then
            IOText = "MECD Telegraph FWE "
        ElseIf VB_More_integers(2) > EOT_instructor(8) Then
            IOText = "MECD Telegraph Dead Slow Astern "
        ElseIf VB_More_integers(2) > EOT_instructor(9) Then
            IOText = "MECD Telegraph Slow Astern "
        ElseIf VB_More_integers(2) > EOT_instructor(10) Then
            IOText = "MECD Telegraph Half Astern "
        Else
            IOText = "MECD Telegraph Full Astern "
        End If

    Else
    
        Get #fp, IONum, InfoString
        IOText = Mid$(InfoString, 1, 100)
        
    End If
    
Else
        Get #fp, IONum, InfoString
        IOText = Mid$(InfoString, 1, 100)
End If

' === 06/06/02 PATCH because the .xls spreadsheet HAS TOTAL MISSED OUT DI's 948 to 955 ===
' so added this patch instead....
If IOType = "DI" Then

    If IONum = 948 Then
        IOText = "Frd Power Panel MCCB Open"
    ElseIf IONum = 950 Then
        IOText = "Frd Power Panel MCCB Reset"
        
    ElseIf IONum = 949 Then
        IOText = "Cargo Panel # 1 MCCB Close"
    ElseIf IONum = 951 Then
        IOText = "Cargo Panel # 1 MCCB Open"
    ElseIf IONum = 952 Then
        IOText = "Cargo Panel # 1 MCCB Reset"
    
    ElseIf IONum = 953 Then
        IOText = "Cargo Panel # 2 MCCB Close"
    ElseIf IONum = 954 Then
        IOText = "Cargo Panel # 2 MCCB Open"
    ElseIf IONum = 955 Then
        IOText = "Cargo Panel # 2 MCCB Reset"
        
    End If
        
    
    ' === Also Texas has a SECOND Ballast Pump
    If TARGET_this_sim = TEXAS Then
    
        If IONum = 912 Then
            IOText = "Ballast Pump # 1 MCCB Close"
        ElseIf IONum = 913 Then
             IOText = "Ballast Pump # 1 MCCB Open"
        ElseIf IONum = 914 Then
             IOText = "Ballast Pump # 1 MCCB Reset"
        End If
        
        If IONum = 993 Then
            IOText = "Ballast Pump # 2 MCCB Close"
        ElseIf IONum = 994 Then
             IOText = "Ballast Pump # 2 MCCB Open"
        ElseIf IONum = 995 Then
             IOText = "Ballast Pump # 2 MCCB Reset"
        End If
        
    End If
    
End If
' ==========================================================================================

Close #fp

Exit Sub

Errhandler:
IOText = "Error Reading Text for " & IOType & "(" & IONum & ")"

'12/12/00 'can't reproduce it
Close #fp

End Sub
Sub FillPVLookUp()

PVLookUp(0) = 50

End Sub
Sub FILL_VIRTUAL_PUMP_TEXT()
Dim i As Integer

' === 13/03/01 ===
' NOTE: There are "VIRTUAL PUMPS" between
'    DI's 3000 and 3199 ( See Breaker X-ref V15.xls )

'Global VIRTUAL_PUMP_TEXT(200) As String

For i = 0 To 200
  VIRTUAL_PUMP_TEXT(i) = "DI(" & Format(i, "##0") & ") is not assignied"
Next i

VIRTUAL_PUMP_TEXT(0) = "FO Transfer Pump # 1 STOP "
VIRTUAL_PUMP_TEXT(1) = "FO Transfer Pump # 1 START"
VIRTUAL_PUMP_TEXT(2) = "FO Transfer Pump # 1 STBY "

VIRTUAL_PUMP_TEXT(9) = "Emergency Feed Pump STOP "
VIRTUAL_PUMP_TEXT(10) = "Emergency Feed Pump START"
VIRTUAL_PUMP_TEXT(11) = "Emergency Feed Pump STBY "

VIRTUAL_PUMP_TEXT(21) = "FO Transfer Pump # 2 STOP "
VIRTUAL_PUMP_TEXT(22) = "FO Transfer Pump # 2 START"
VIRTUAL_PUMP_TEXT(23) = "FO Transfer Pump # 2 STBY "

VIRTUAL_PUMP_TEXT(33) = "LO Purifier STOP "
VIRTUAL_PUMP_TEXT(34) = "LO Purifier START"
VIRTUAL_PUMP_TEXT(35) = "LO Purifier AUTO "

' === 30/05/02 ===
'VIRTUAL_PUMP_TEXT(36) = "Ballast pump STOP "
'VIRTUAL_PUMP_TEXT(37) = "Ballast pump START"
'VIRTUAL_PUMP_TEXT(38) = "Ballast pump STBY "
VIRTUAL_PUMP_TEXT(36) = "Ballast pump # 1 STOP "
VIRTUAL_PUMP_TEXT(37) = "Ballast pump # 1 START"
VIRTUAL_PUMP_TEXT(38) = "Ballast pump # 1 STBY "

VIRTUAL_PUMP_TEXT(42) = "Control Air Compressor STOP "
VIRTUAL_PUMP_TEXT(43) = "Control Air Compressor START"
VIRTUAL_PUMP_TEXT(44) = "Control Air Compressor STBY "

VIRTUAL_PUMP_TEXT(45) = "ADT pump # 1 STOP "
VIRTUAL_PUMP_TEXT(46) = "ADT pump # 1 START"
VIRTUAL_PUMP_TEXT(47) = "ADT pump # 1 STBY "

VIRTUAL_PUMP_TEXT(48) = "ADT pump # 2 STOP "
VIRTUAL_PUMP_TEXT(49) = "ADT pump # 2 START"
VIRTUAL_PUMP_TEXT(50) = "ADT pump # 2 STBY "

VIRTUAL_PUMP_TEXT(51) = "GS pump STOP "
VIRTUAL_PUMP_TEXT(52) = "GS pump START"
VIRTUAL_PUMP_TEXT(53) = "GS pump STBY "

VIRTUAL_PUMP_TEXT(78) = "FWG SW pump STOP "
VIRTUAL_PUMP_TEXT(79) = "FWG SW pump START"
VIRTUAL_PUMP_TEXT(80) = "FWG SW pump STBY "

VIRTUAL_PUMP_TEXT(81) = "FWG distillate pump STOP "
VIRTUAL_PUMP_TEXT(82) = "FWG distillate pump START"
VIRTUAL_PUMP_TEXT(83) = "FWG distillate pump STBY "

' === 30/05/02 ADDED: ===
VIRTUAL_PUMP_TEXT(87) = "Ballast pump # 2 STOP "
VIRTUAL_PUMP_TEXT(88) = "Ballast pump # 2 START"
VIRTUAL_PUMP_TEXT(89) = "Ballast pump # 2 STBY "


VIRTUAL_PUMP_TEXT(126) = "Service Air Compressor STOP "
VIRTUAL_PUMP_TEXT(127) = "Service Air Compressor START"
VIRTUAL_PUMP_TEXT(128) = "Service Air Compressor STBY "

End Sub
