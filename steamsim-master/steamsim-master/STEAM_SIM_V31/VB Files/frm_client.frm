VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm frm_client 
   BackColor       =   &H8000000C&
   Caption         =   "Client Sockets"
   ClientHeight    =   6690
   ClientLeft      =   1095
   ClientTop       =   2085
   ClientWidth     =   7995
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin CSSocketWrench.Socket Socket1 
      Index           =   0
      Left            =   300
      Top             =   1200
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   -1  'True
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin CSSocketWrench.Socket Socket2 
      Left            =   300
      Top             =   1680
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   660
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   6345
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   609
      SimpleText      =   "Hello"
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Unknown"
            TextSave        =   "Unknown"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2884
            Text            =   "Simulator Not Started "
            TextSave        =   "Simulator Not Started "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MouseIcon       =   "frm_client.frx":0000
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_snapshot_load 
         Caption         =   "&Load Snapshot"
      End
      Begin VB.Menu mnu_snapshot_save 
         Caption         =   "&Save Snapshot"
      End
      Begin VB.Menu mnu_hyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_data_record 
         Caption         =   "Start R&Ecord"
      End
      Begin VB.Menu mnu_data_replay 
         Caption         =   "Start Re&Play"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_hyphen4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_log_actions 
         Caption         =   "L&oG"
         Begin VB.Menu mnu_log_view 
            Caption         =   "&View"
         End
         Begin VB.Menu mnu_log_print 
            Caption         =   "&Print"
         End
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&Xit"
      End
   End
   Begin VB.Menu mnu_simulator 
      Caption         =   "&Simulator"
      Begin VB.Menu mnu_sim_run 
         Caption         =   "&Run"
         Begin VB.Menu Initialise 
            Caption         =   "&Initialse"
         End
         Begin VB.Menu mnu_sim_resume 
            Caption         =   "&Resume"
         End
      End
      Begin VB.Menu mnu_sim_halt 
         Caption         =   "&Halt"
      End
      Begin VB.Menu mnu_sim_reset 
         Caption         =   "Re&Set"
      End
      Begin VB.Menu mnu_hyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_fault_entry 
         Caption         =   "&Faults"
         Begin VB.Menu mnu_faluts_entry 
            Caption         =   "&Entry/Clearance"
         End
         Begin VB.Menu mnu_faults_queued 
            Caption         =   "&Queued Fault Display"
         End
      End
      Begin VB.Menu mnu_instruct_params 
         Caption         =   "&Instructor Parameters"
      End
      Begin VB.Menu mnu_next_fault_control 
         Caption         =   "&Next Fault Control"
      End
      Begin VB.Menu mnu_mobile_fault_control 
         Caption         =   "&Mobile Fault Control Unit"
      End
      Begin VB.Menu mnu_hyphen9 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_hw_test 
         Caption         =   "&Hardware Test"
         Begin VB.Menu DItest 
            Caption         =   "&DI test"
         End
         Begin VB.Menu DOtest 
            Caption         =   "&DO test"
            Begin VB.Menu BreakerTest 
               Caption         =   "&Breaker test"
            End
            Begin VB.Menu LampTest 
               Caption         =   "&Lamp test"
            End
            Begin VB.Menu DOIndividual 
               Caption         =   "&Individual"
            End
         End
         Begin VB.Menu AItest 
            Caption         =   "&AI  test"
         End
         Begin VB.Menu AnalogOutputTest 
            Caption         =   "&AO tests"
         End
      End
   End
   Begin VB.Menu mnu_graphics 
      Caption         =   "&Graphics"
      Begin VB.Menu mnu_g_rr 
         Caption         =   "&Real/Reactive Power"
      End
      Begin VB.Menu mnu_g_cvp 
         Caption         =   "&Current/Voltage Phasor"
      End
      Begin VB.Menu mnu_g_synch_phasors 
         Caption         =   "&Synchronising Phasor"
      End
      Begin VB.Menu LoadShare 
         Caption         =   "&Load Share"
      End
      Begin VB.Menu GenCharacteristics 
         Caption         =   "&Generator Characteristics"
      End
      Begin VB.Menu mnu_offline 
         Caption         =   "&Off-Line"
         Begin VB.Menu mnu_g_offline_rr 
            Caption         =   "&Real/Reactive Power"
         End
         Begin VB.Menu mnu_g_offline_cvp 
            Caption         =   "&Current/Voltage Phasor"
         End
      End
      Begin VB.Menu mnu_g_status 
         Caption         =   "&Status"
      End
      Begin VB.Menu mnu_power_distribution 
         Caption         =   "Power &Distribution"
      End
      Begin VB.Menu mnu_Panel_A 
         Caption         =   "&Panel  A"
      End
      Begin VB.Menu mnu_panel_B 
         Caption         =   "&Panel  B"
      End
   End
   Begin VB.Menu mnu_tcpip 
      Caption         =   ""
      Begin VB.Menu mnu_tcpip_array 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "frm_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LastSocket As Integer


' An array of 2500 singles can be reliably transmitted on one machine in a loop using 32k buffers
' Above 2500 errors begin to appear
' When going through one machine, data speed is aprox 1000 items per second

' When using two machines, transmitting the data takes less than half a second


' ==== 26/3/97 RWJ moved to main.bas (made GLOBAL CONST etc =======
' Number of elements in incoming array
'Const array_size = 200

' The Ports are used for the communiction sockets - 2000..2005 are known to be useable.
'Const PORT_0 = 2000
'Const PORT_1 = 2001

' Maximum nomber of sockets receiving data from Host (should be the
' same as in the Host program )
'Const Max_Socket = 1
' ============================================================

' One counter for each array of incoming data
Dim counter(Max_Socket) As Integer

Sub HH_TPFLTCHECK(IFNUM As Integer, I_TNFXR As Integer)

' 12/3/97
Dim i As Integer
Dim SPFAULT(8) As Integer
Dim SpecialFault As Integer
Dim ivalue As Integer

'Dim MessageString  As String

Dim fault_no As String
Dim freq_val As String
Dim volt_val As String

'  RWJ To  SET a fault:
   fault_no = "0"
   freq_val = "1"
   volt_val = "1"

'    MessageString = Fault_string & fault_no & "#" & freq_val & "#" & volt_val & "#"
'    frm_exec.Socket2.SendLen = Len(MessageString)
'    frm_exec.Socket2.SendData = MessageString
 
' 12/3/97   NOTE: This code was taken from TPFC1.F90

'  see : Call HH_TPFLTCHECK(IFNUM, I_TNFXR)

'
'
'           DATA SPFAULT /11,14,17,20,71,74,77,80/
            SPFAULT(1) = 11
            SPFAULT(2) = 14
            SPFAULT(3) = 17
            SPFAULT(4) = 20
            SPFAULT(5) = 71
            SPFAULT(6) = 74
            SPFAULT(7) = 77
            SPFAULT(8) = 80
         
         
'C special faults require additional instructor input eg Hz or V
'
'
'c 7 / 1 / 97
'c           SAVE /I/,/ILINE/,/IPAGE/,/COL/,/ROW/,/CVALUE/,/FDATA/,/INPAGE/,
'c     +          /ERROR/,/IFDATA/,/SPECIALFAULT/,/PROMPT/
'c
'
'              Col = 0
'              Row = 22
'
'C decode chars into integer, check entry is valid, check if special fault
'C requiring extra data input


'c 17/1/97     CALL CHARDECODE(CVALUE,IFNUM,ERROR)
'            IFNUM = IVALUE
'            ERROR = .false.
'c 14 / 2 / 97
'                   write(6,18) IFNUM,TNFXR(IFNUM)
' 18                format(' IFNUM = ',I6,'   TNFXR(IFNUM) = ',I6)
'

'              IF(.NOT.ERROR.AND.IFNUM.LT.900.AND.TNFXR(IFNUM).NE.0) THEN
'                  SPECIALFAULT = .FALSE.
'                  DO 20 I = 1,8
'                    IF(IFNUM.EQ.SPFAULT(I)) SPECIALFAULT = .TRUE.
'20                CONTINUE
              If (IFNUM < 900 And I_TNFXR > 0) Then
                   SpecialFault = False
                  For i = 1 To 8
                     If (IFNUM = SPFAULT(i)) Then SpecialFault = True
                  Next i
'
'
'c 14 / 2 / 97
'                   write(6,19) IFNUM,SPECIALFAULT
' 19                format(' IFNUM = ',I6,'  SPECIALFAULT = ',L4)
'
'
'                  If (SPECIALFAULT) Then
                   If (SpecialFault) Then
'                     IF(IFNUM.LT.21) THEN
                      If (IFNUM < 21) Then
'
'C communicate with VDU unless entry is from hand held unit..
'
'                        IF(VDU_TYPE.NE.2) THEN
' 30                        WRITE(6,40)ESC,POSITION,COL,ROW,ESC,CLEAREOL
' 40                        FORMAT(4A,'Hz ? .... ',2A,$)
'
'c 17 / 1 / 97
'c                           READ(5,50)CVALUE
'  50                        FORMAT(3A)
'
'c 17/1/97                  IF(CVALUE.EQ.' ') RETURN
'
'c 17/1/97                  CALL CHARDECODE(CVALUE,IVALUE,ERROR)
'                         IVALUE = Ifreq
'
'                           IF(.NOT.ERROR.AND.IVALUE.GE.50.AND.IVALUE.LE.65) THEN
'                               FFREQS(IFNUM) = IVALUE
'                           Else
'                               WRITE(6,60)BLEEP
' 60                            FORMAT(1A,$)
'c 20/1/97                      GOTO 30
'                           End If
'
'                        Else
'
'c 14 / 2 / 97
'                          write(6,*) ' Using Hand Held terminal for Frequency entry '
'
line_31:
'c 20 / 1 / 97
'c                         The Hand Held Terminal is beind used...
'c31                        WRITE(6,41)HHESC,HH_HOMED,HHESC,HHCLEAR
' 31                        WRITE(7,41)HHESC,HH_HOMED,HHESC,HHCLEAR
                            MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeD) & Chr$(HHEsc) & Chr$(HHClear)

' 41                        FORMAT(4A,'Hz ? .... ',$)
                            MSComm1.Output = "Hz ? .... "

'
'c 20/1/97                  READ(5,50)CVALUE
'                           READ(7,50)CVALUE
                            Call HH_read(3, ENTRY)

'                           IF(CVALUE.EQ.' ') RETURN
                            If (Mid$(ENTRY, 1, 1) = Chr$(13)) Then GoTo end_this_sub:
'
'                           Call CHARDECODE(CVALUE, IVALUE, ERROR)
                            ivalue = Val(ENTRY)
                            
'                           IF(.NOT.ERROR.AND.IVALUE.GE.50.AND.IVALUE.LE.65) THEN
                            If (ivalue >= 50 And ivalue <= 65) Then
                            
'                               FFREQS(IFNUM) = IVALUE
' NOTE: Need to set             FFREQS(IFNUM)
'                                   MessageString = Fault_string & fault_no & "#" & freq_val & "#" & volt_val & "#"
                                    MessageString = Fault_string & Str$(IFNUM) & "#" & ENTRY & "#" & "0" & "#"
                                    frm_exec.Socket2.SendLen = Len(MessageString)
                                    frm_exec.Socket2.SendData = MessageString



                            Else
                            
'c                              WRITE(6,61)HHESC,HH_HOMED,HHESC,HHCLEAR
'                               WRITE(7,61)HHESC,HH_HOMED,HHESC,HHCLEAR
                                MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeD) & Chr$(HHEsc) & Chr$(HHClear)
                                
' 61                            FORMAT(4A,'55-65 Hz only')
                                MSComm1.Output = "55-65 Hz only"
'
'                               Call SLEEPqq(2000)
                                Call Sleep(2000)
'
'                               GoTo 31
                                GoTo line_31:
                                
                            End If
'
'                        End If
'
                     Else
'
'                        IF(VDU_TYPE.NE.2) THEN
'
' 70                        WRITE(6,80)ESC,POSITION,COL,ROW,ESC,CLEAREOL
' 80                        FORMAT(4A,'Volts ? .... ',2A,$)
'
'c 17 / 1 / 97
'c                           READ(5,50)CVALUE
'c                           IF(CVALUE.EQ.' ') RETURN
'
'c                           CALL CHARDECODE(CVALUE,IVALUE,ERROR)
'                         IVALUE = Ivolts
'
'                           IF(.NOT.ERROR.AND.IVALUE.GE.410.AND.IVALUE.LE.475) THEN
'                              FVOLTS(IFNUM) = IVALUE
'                           Else
'                              WRITE(6,60)BLEEP
'c 20/1/97                     GOTO 70
'                           End If
'
'                        Else
'
'
'c 14 / 2 / 97
'                          write(6,*) ' Using Hand Held terminal for Volts entry '
'
line_71:
'c 20/1/97  Hand-held Terminal changed from unit 6 to unit 7
'c71                        WRITE(6,81)HHESC,HH_HOMED,HHESC,HHCLEAR
' 71                        WRITE(7,81)HHESC,HH_HOMED,HHESC,HHCLEAR
                            MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeD) & Chr$(HHEsc) & Chr$(HHClear)

' 81                        FORMAT(4A,'Volts ?... ',$)
                            MSComm1.Output = "Volts ?... "

'                           READ(7,50)CVALUE
                            Call HH_read(3, ENTRY)

'                           IF(CVALUE.EQ.' ') RETURN
                            If (Mid$(ENTRY, 1, 1) = Chr$(13)) Then GoTo end_this_sub:
'
'                           Call CHARDECODE(CVALUE, IVALUE, ERROR)
                            ivalue = Val(ENTRY)

'                           IF(.NOT.ERROR.AND.IVALUE.GE.410.AND.IVALUE.LE.475) THEN
                            If (ivalue >= 410 And ivalue <= 475) Then

'                              FVOLTS(IFNUM) = IVALUE
' NOTE: Need to SET            FVOLTS(IFNUM)
                                    MessageString = Fault_string & Str$(IFNUM) & "#" & "0" & "#" & ENTRY & "#"
                                    frm_exec.Socket2.SendLen = Len(MessageString)
                                    frm_exec.Socket2.SendData = MessageString

                            Else

'                              WRITE(7,82)HHESC,HH_HOMED,HHESC,HHCLEAR
                               MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeD) & Chr$(HHEsc) & Chr$(HHClear)
' 82                           FORMAT(4A,'420-475 V only!',$)
                               MSComm1.Output = "420-475 V only!"
'
'                              Call SLEEPqq(2000)
                               Call Sleep(2000)
'
'                              GoTo 71
                               GoTo line_71:
'
                            End If

                        End If
                        
                     Else      ' ( NOT a specialfault )
                     
' 18/3/97 SET A FAULT:
                                    MessageString = Fault_string & Str$(IFNUM) & "#" & "0" & "#" & "0" & "#"
                                    frm_exec.Socket2.SendLen = Len(MessageString)
                                    frm_exec.Socket2.SendData = MessageString
                                    
                     End If        '  Endif of the IF SPECIALFAULT
                     
                  End If
                  
' ====================================

'   NOTE: WHAT about all the rest of this code TNFAULTDATA(TNFXR(IFNUM),1) etc etc   ??????

' ====================================
                  
                  

'C now set fault in array and assign value if special fault - but do this only
'C if fault has NOT been entered as a scenario fault ( inthis case VDU_TYPE = 1 )



'                  IF(VDU_TYPE.NE.1) THEN
'                     FAULT(IFNUM) = .TRUE.
' 17/3/97 SET A FAULT:
' 18/3/97 remove:
'                                    MessageString = Fault_string & Str$(IFNUM) & "#" & "0" & "#" & "0" & "#"
'                                    frm_exec.Socket2.SendLen = Len(MessageString)
'                                    frm_exec.Socket2.SendData = MessageString


'                     TNFAULTDATA(TNFXR(IFNUM), 1) = HOURS
'                     TNFAULTDATA(TNFXR(IFNUM), 2) = MINS
'                     TNFAULTDATA(TNFXR(IFNUM), 3) = SECS
'                     IF(IFNUM.EQ.11) FSHORE1PCP = IVALUE
'                     IF(IFNUM.EQ.14) FSHORE2PCP = IVALUE
'                     IF(IFNUM.EQ.71) VSHORE1PCP = IVALUE
'                     IF(IFNUM.EQ.74) VSHORE2PCP = IVALUE
'                     ERROR = .FALSE.
'c 14 / 2 / 97
'                          write(6,*) ' RETURN from TPFC1.F90 '

'                     Return
'
'                  Else

'C if a scenario fault then fill fault queue with fault number + value if
'C special fault

'                     IF(FAULTCOUNT(2).GT.10) THEN
'                         WRITE(6,90)ESC,POSITION,COL,ROW,ESC,CLEAREOL,BLEEP
' 90                      FORMAT(7A,'Fault Queue is Full',$)
'
'                         Call SLEEPqq(2000)

'                     Else
'                        FAULTQUEUE(FAULTCOUNT(2)) = IFNUM
'                        IF(IFNUM.EQ.11) FAULTQVALUE(FAULTCOUNT(2)) = IVALUE
'                        IF(IFNUM.EQ.14) FAULTQVALUE(FAULTCOUNT(2)) = IVALUE
'                        IF(IFNUM.EQ.71) FAULTQVALUE(FAULTCOUNT(2)) = IVALUE
'                        IF(IFNUM.EQ.74) FAULTQVALUE(FAULTCOUNT(2)) = IVALUE
'                        FAULTCOUNT(2) = FAULTCOUNT(2) + 1
'                     End If
'
'                  End If
'
'
'              Else
'
'c 14 / 2 / 97
'                   write(6,*) ' Using Hand Held terminal for Error message '
'
'                  IF(VDU_TYPE.NE.2) THEN
'                     WRITE(6,60)BLEEP
'                     ERROR = .TRUE.
'                  Else
'                     WRITE(6,91)HHESC,HH_HOMEU,HHESC,HHCLEAR
' 91                  FORMAT(4A,'Error',/,'Retry ... ',$)
'
'                     Call SLEEPqq(2000)
'
'                  End If
'
'              End If
 

end_this_sub:


End Sub


Sub HH_read(Nchars As Integer, Instring As String)
Dim dummy
Dim i As Integer, j As Integer
' 11/3/97
Dim incount As Integer
Dim inchar(10) As String * 1

' 10/3/97
If Not MSComm1.PortOpen Then
  MSComm1.PortOpen = True
End If
  
  
For i = 1 To Nchars
    Do
        dummy = DoEvents()
'   Loop Until MSComm1.InBufferCount >= Nchars
    Loop Until MSComm1.InBufferCount >= 1
    
' Read the serial port.
'   Instring = MSComm1.Input
    inchar(i) = MSComm1.Input
' Echo the character back to the Hand Held VDU
    MSComm1.Output = inchar(i)
'    Instring = Instring + MSComm1.Input
    If (i = Nchars Or inchar(i) = Chr$(13)) Then GoTo exit_read:
Next i
    
exit_read:

   Instring = ""
For j = 1 To i
   Instring = Instring + inchar(j)
Next j
    
'    MSComm1.Output = Instring
    
' TEST ONLY:
'text1(0).Text = ""
'For i = 1 To Len(Instring)
'    text1(0).Text = text1(0).Text & Hex(Asc(Mid$(Instring, i, 1))) & "H "
'Next i

End Sub


Sub LogEvents()
Dim AIStart As Integer
Dim AIEnd As Integer
Dim i As Integer
Dim j As Integer
Dim InfoString As String * 100
Dim DInumb As Integer

'Debug.Print "entered logevents"

T42RUN = PC_INTS(90)
PCPRUN = PC_INTS(91)
SCPRUN = PC_INTS(92)

AIStart = 1
AIEnd = 11
If (PCP_Executive And (T42RUN Or PCPRUN)) Or (Not PCP_Executive And SCPRUN) Then
    If PCP_Executive Then
        AIStart = 1
        If SPLITMODE Then
            AIEnd = 5
        Else
            AIEnd = 11
        End If
    Else
        AIEnd = 11
        If SPLITMODE Then
            AIStart = 6
        Else
            AIStart = 1
        End If
    End If

    If Not DI_Test Then                                                                        ' don't log during a DI hardware test
        For i = 0 To 20                                                                             'check for any DI changes
            DInumb = PC_INTS(119 + i)
            If DInumb > 0 Then
                Call Logging.AddEvent("DI", IOTEXT(DInumb))              'log the event by adding to the list
                'Debug.Print "added di " & DInumb & " to list"
                MessageString = "#CLEAR_DI_CHANGES#" & Str$(DInumb) & "#"
                Call SendMessage(MessageString)
                Call Sleep(100)
            End If
        Next
        'MessageString = "#CLEAR_DI_CHANGES#" & Str$("3071") & "#"
       'Call SendMessage(MessageString)
        'If DIchanges(0) > 0 Then                                                          'clear down the list of changes because thy are all logged now
            'MessageString = "#CLEAR_DI_CHANGES#"
            'Call SendMessage(MessageString)
        'End If
        
   
    End If

    'Clear the info string first
    For i = 1 To 100
        Mid(InfoString, i, 1) = " "
    Next
    For i = AIStart To AIEnd
        
        AI(AIXREF(i)) = PC_INTS(107 + i)
     
    Next



    If Not AI_test Then                                                                         'don't log AI changes during AI test
        For i = AIStart To AIEnd
            Select Case i
            Case 1, 2, 3, 4, 6, 7, 9, 10                                                    'don't log synchro changes - exclude numbers 5,8, and 11
            'Debug.Print "AI(AIXREF(i) = " & AI(AIXREF(i))
            'Debug.Print "OldAI(AIXREF(i)) =" & OldAI(AIXREF(i))
            If Abs(AI(AIXREF(i)) - OldAI(AIXREF(i))) >= 7 Then          'log if AI changes by 4 or more (approx 3%)
            'If AI(AIXREF(i)) - OldAI(AIXREF(i)) >= 4 Or AI(AIXREF(i)) - OldAI(AIXREF(i)) <= -4 Then          'log if AI changes by 4 or more (approx 3%)

                'InfoString = Date & Time
                Mid(InfoString, 1, 40) = LTrim(AITEXT(AIXREF(i)))
                Mid(InfoString, 41, 13) = "  Adjusted to "
                Mid(InfoString, 55, 3) = Str(AI(AIXREF(i)) * 100 / 255)
                Mid(InfoString, 60, 1) = "%"
                OldAI(AIXREF(i)) = AI(AIXREF(i))                                    'reset old AI to pick up future changes
                Call Logging.AddEvent("AI", InfoString)             'log the event by adding to the list
            
            End If
            End Select
        Next
    End If
End If
End Sub

Sub mfu_control()

'   6/5/97  added
Dim offset As Integer

Dim fault_no As String
Dim freq_val As String
Dim volt_val As String

Dim gen_no As String
Dim ovpu_no As String
Dim ovpu As Integer

Dim i As Integer
Dim msg As Integer

Dim FCHARS As String * 3
Dim GCHARS As String * 15

Dim IFNUM As Integer
Dim ERROR As Integer
Dim FAULT_IFNUM As Integer, FLT_IFNUM As Byte
Dim Igen As Integer

Dim status As String * 10
Dim ThisFaultnumber As Integer
Dim FaultOK As Integer
Dim FaultAlreadyEntered As Integer

Dim Igeny As Integer

   fault_no = "0"
   freq_val = "1"
   volt_val = "1"

'8/5/97 handle split mode too
If SPLITMODE Then
        If PCP_Executive Then
                   offset = 0
        Else
                   offset = 6
        End If
Else
                   offset = 0
End If



msg = MsgBox("Transfer control to the Hand Held Terminal ?", 1, " Hand Held Terminal")
If msg = 2 Then Exit Sub                 'Cancel is pressed

ERROR = False

 ' Open the port first time around note this will be same for scp or pcp
If Not MSComm1.PortOpen Then
    MSComm1.PortOpen = True
End If
            
'Put the light on
MSComm1.Output = Chr$(HHEsc) & Chr$(HHLightOn)


'set cursor top left + display menu:
line_70:
MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
MSComm1.Output = "MENU: F G N O  R" & Chr$(10) & Chr$(13) & "enter code : "
           
'read up to 4 chars
Call HH_read(4, ENTRY)

'Fault Options, display.  there are 4 options for faults :  entry,clear indiv faults, clear all and display indiv faults
If Mid$(ENTRY, 1, 1) = "F" Then
line_110:
        
    MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
    MSComm1.Output = "FAULT OPTIONS: " & Chr$(10) & Chr$(13) & "F,C,A or D "

    Call HH_read(4, ENTRY)
     If Mid$(ENTRY, 1, 1) = "F" Then                                    'Request fault number if F is entered followed by CR
                 If Mid$(ENTRY, 2, 1) = Chr$(13) Then
                        MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                        MSComm1.Output = "Fault No ? "
                        Call HH_read(4, FCHARS)                         'Read Fault Number as a string
                        If (Mid$(FCHARS, 1, 1) = Chr$(13)) Then GoTo line_70:
                 Else
                        FCHARS = Mid$(ENTRY, 2, 3)                   'Decode numbers after letter F to give fault number as a string
                 End If
                 
line_180:
                    ThisFaultnumber = Val(FCHARS)
                    Call CheckFault(ThisFaultnumber, FaultOK, FaultAlreadyEntered)
                    Debug.Print "called faultcheck, Fault # " & ThisFaultnumber & "FaultOK = " & FaultOK
                    Debug.Print "FaultAlreadyEntered =" & FaultAlreadyEntered
                    
                    If Not FaultOK Or FaultAlreadyEntered Then
                         If ThisFaultnumber <> 0 Then
                            MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                            MSComm1.Output = "ERROR:  Not Avail" & Chr$(10) & Chr$(13) & "Next fault ? .. "
                            Call Sleep(2000)
                        End If
                        GoTo line_110
                    Else
                        If Not PCP_Executive And SPLITMODE Then                       'If SCP in splitmode, add an offset so model refers to SCP
                                                                                                                                'This is either 6 or 8 depending on type of fault
                                If (ThisFaultnumber >= 350 And ThisFaultnumber <= 357) Or (ThisFaultnumber >= 470 And ThisFaultnumber <= 496) _
                                                                       Or (ThisFaultnumber >= 150 And ThisFaultnumber <= 176) Then
                                     offset = 8
                                Else
                                    offset = 6
                                End If
                        Else
                                offset = 0
                        End If
                        
                    Select Case ThisFaultnumber                                                        'Prompt for Volts or Freq for Special Faults
                        
    ' 8/5/97    Case 11, 17
                     Case 11, 14
                       
                        MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                        MSComm1.Output = "Freq (Hz) ? .."
                        Do
                            Call HH_read(3, freq_val)
                            If Val(freq_val) >= 55 And Val(freq_val) <= 65 Then
                                Exit Do
                            Else
                                If Val(freq_val) = 0 Then GoTo line_180
                                MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                                MSComm1.Output = "ERROR: Enter  " & Chr$(10) & Chr$(13) & "55-65 Hz ..."
                            End If
                        Loop
                        
                    Case 71, 74
                        
                        MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                        MSComm1.Output = "Voltage ?.."
                        Do
                            Call HH_read(3, volt_val)
                            If Val(volt_val) > 409 And Val(volt_val) < 476 Then
                                Exit Do
                            Else
                                If Val(volt_val) = 0 Then GoTo line_180
                                MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                                MSComm1.Output = "ERROR: Enter " & Chr$(10) & Chr$(13) & "410-475 V  ..."
                            End If
                        Loop
                       
                    End Select
                    
                    ThisFaultnumber = ThisFaultnumber + offset        ' Add Offset before calling DLL
                    MessageString = Fault_string & Str$(ThisFaultnumber) & "#" & freq_val & "#" & volt_val & "#" ' Enter Vault plus any values
                    Call SendMessage(MessageString)
                    Debug.Print "setting fault number : " & ThisFaultnumber
                    'frm_exec.Socket2.SendLen = Len(MessageString)
                    'frm_exec.Socket2.SendData = MessageString
                End If
                'Ask for next fault as a string
                MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                MSComm1.Output = "Next fault .."
                    
                Call HH_read(4, FCHARS)
                GoTo line_180:
   
'Clear Single fault
    ElseIf Mid$(ENTRY, 1, 1) = "C" Then
              
                   If Mid$(ENTRY, 2, 1) = Chr$(13) Then
                            MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                            MSComm1.Output = "Clear No ? "
                            Call HH_read(3, FCHARS)                                 'read number to be cleared as a separate string
                   Else
                            FCHARS = Mid$(ENTRY, 2, 3)                            'read number to be cleared as a string after "C"
                             If (Mid$(FCHARS, 1, 1) = Chr$(13)) Then GoTo line_70:
                   End If
                  
line_240:
                    If (Mid$(FCHARS, 1, 1) = Chr$(13)) Then GoTo line_110:
                    ThisFaultnumber = Val(FCHARS)
                    If ThisFaultnumber = 0 Then GoTo line_110
                    Call CheckFault(ThisFaultnumber, FaultOK, FaultAlreadyEntered)
                   ' 13/5/97 dont check if Already Entered as this requires frm_faults to be loaded If Not FaultOK Or Not FaultAlreadyEntered Then
                        
                    If Not FaultOK Then

                         MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                         MSComm1.Output = "ERROR:  Not " & Chr$(10) & Chr$(13) & "Cleared, Next .."
                         Call HH_read(3, FCHARS)
                         'Call Sleep(2000)
                         GoTo line_240
                    Else
                        If Not PCP_Executive And SPLITMODE Then                       'If SCP in splitmode, add an offset so model refers to SCP
                                                                                                                                'This is either 6 or 8 depending on type of fault
                                If (ThisFaultnumber >= 350 And ThisFaultnumber <= 357) Or (ThisFaultnumber >= 470 And ThisFaultnumber <= 496) _
                                                                       Or (ThisFaultnumber >= 150 And ThisFaultnumber <= 176) Then
                                     offset = 8
                                Else
                                    offset = 6
                                End If
                        Else
                                offset = 0
                        End If
                        ThisFaultnumber = ThisFaultnumber + offset        ' Add Offset before calling DLL
                         MessageString = Fault_C1_string & Str$(ThisFaultnumber) & "#" & "0" & "#" & "0" & "#"
                         'frm_exec.Socket2.SendLen = Len(MessageString)
                         'frm_exec.Socket2.SendData = MessageString
                          Debug.Print "Clearing fault number : " & ThisFaultnumber
                         Call SendMessage(MessageString)
                         MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                         MSComm1.Output = "Cleared " & Format$(FCHARS, "###") & " ok " & Chr$(10) & Chr$(13) & "Next .."
                         Call HH_read(3, FCHARS)
                         GoTo line_240:
                    End If
                    
'Clear All faults
    ElseIf Mid$(ENTRY, 1, 1) = "A" Then
                                    
                MessageString = Fault_CA_string & Str$(IFNUM) & "#" & "0" & "#" & "0" & "#"
                frm_exec.Socket2.SendLen = Len(MessageString)
                frm_exec.Socket2.SendData = MessageString
                MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                MSComm1.Output = "All Faults" & Chr$(10) & Chr$(13) & "     Cleared OK"
     
                Call Sleep(3000)
                  
                GoTo line_110:
                
'Display a fault text
    ElseIf Mid$(ENTRY, 1, 1) = "D" Then
               
                   If Mid$(ENTRY, 2, 1) = Chr$(13) Then

                       MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                       MSComm1.Output = "Display Fault " & Chr$(10) & Chr$(13) & "Number ? "
                       Call HH_read(3, FCHARS)
                  Else
                       FCHARS = Mid$(ENTRY, 2, 3)
                  End If
 
line_290:
                   If (Mid$(FCHARS, 1, 1) = Chr$(13)) Then GoTo line_70:            'Back to main menu if CR
                   
                    ThisFaultnumber = Val(FCHARS)
                    If ThisFaultnumber = 0 Then GoTo line_70
                    Call CheckFault(ThisFaultnumber, FaultOK, FaultAlreadyEntered)
                    If Not FaultOK Then
                         MSComm1.Output = "ERROR:  Not Available" & Chr$(10) & "Next ....."
                         Call Sleep(2000)
                         Call HH_read(3, FCHARS)
                         GoTo line_290
                    Else
                        MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                        MSComm1.Output = Mid$(FAULTEXT(FaultXRef(ThisFaultnumber)), 7, 32)
                    End If
                    Call Sleep(3000)                'wait 3 secs and output a second line if necessary
                    If (Mid$(FAULTEXT(FaultXRef(ThisFaultnumber)), 39, 11) <> "") Then

                          MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                          MSComm1.Output = Mid$(FaultXRef(ThisFaultnumber), 38, 12)
                          Call Sleep(3000)
                    End If
    End If
ElseIf (Mid$(ENTRY, 1, 1) = "G") Then

                If (Mid$(ENTRY, 2, 1) = Chr$(13)) Then
               
line_400:
                     MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                     MSComm1.Output = "Gen start/stop " & Chr$(10) & Chr$(13) & "enter 1-6 .."
                     Call HH_read(1, FCHARS)
                 Else
                     FCHARS = Mid$(ENTRY, 2, 1)
                 End If
                  If (Mid$(FCHARS, 1, 1) = Chr$(13)) Then GoTo line_70:
                 
                  Igen = Val(Mid$(FCHARS, 1, 1))
                   If (Igen < 7 And Igen > 0) Then
                                         Igeny = Igen
                            If Not PCP_Executive And SPLITMODE Then Igen = Igen + 6

                            gen_no = Str$(Igen)
                            ovpu_no = "0"

 ' 8/5/97             If On_line_gen_on(Igen) Then
                           If On_line_gen_on(Igeny) Then
                                MessageString = INSTRP_string & ovpu_no & "#" & gen_no & "#STOP#" & Str$(IPCP_loading) & "#"
                                status = "Stopped"
                            Else
                                MessageString = INSTRP_string & ovpu_no & "#" & gen_no & "#START#" & Str$(IPCP_loading) & "#"
                                status = "Started"
                            End If
                            frm_exec.Socket2.SendLen = Len(MessageString)
                            frm_exec.Socket2.SendData = MessageString
                            If (Igen = 1) Then GCHARS = "GENERATOR M1"
                            If (Igen = 2) Then GCHARS = "SHORE SUPPLY 1"
                            If (Igen = 3) Then GCHARS = "GENERATOR M2"
                            If (Igen = 4) Then GCHARS = "GENERATOR J2"
                            If (Igen = 5) Then GCHARS = "SHORE SUPPLY 2"
                            If (Igen = 6) Then GCHARS = "GENERATOR J1"
                   
                            MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
                            MSComm1.Output = GCHARS & Chr$(10) & Chr$(13) & status

                            Call Sleep(3000)
                    
                 Else
                 
                            GoTo line_400:
                     
                 End If

' reset all OVPU's if a 'O' is enterd  ....


ElseIf (Mid$(ENTRY, 1, 1) = "O") Then
           
        For i = 1 To 4
        
'             If i = 1 Then ovpu_no = "1"
'             If i = 1 Then ovpu_no = "3"
'             If i = 1 Then ovpu_no = "4"
'             If i = 1 Then ovpu_no = "6"
             If i = 1 Then ovpu = 1
             If i = 2 Then ovpu = 3
             If i = 3 Then ovpu = 4
             If i = 4 Then ovpu = 6
             ovpu = ovpu + offset
             ovpu_no = Str(ovpu)
             
'  7/5/97
                      gen_no = "0"
' 7/5/97
'          MessageString = INSTRP_string & "1" & "#" & "0" & "#STOP#" & Str$(IPCP_loading) & "#"
           MessageString = INSTRP_string & ovpu_no & "#" & gen_no & "#STOP#" & Str$(IPCP_loading) & "#"
             frm_exec.Socket2.SendLen = Len(MessageString)
             frm_exec.Socket2.SendData = MessageString
             Call Sleep(200)
             
        Next i

        MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
        MSComm1.Output = "OVPUs reset OK"
        Call Sleep(2000)
        
           
 'Return to main terminal
 ElseIf (Mid$(ENTRY, 1, 1) = "R") Then

               MSComm1.Output = Chr$(HHEsc) & Chr$(HHLightOff) & Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
               MSComm1.Output = "exit to VDU"
               Exit Sub

ElseIf (Mid$(ENTRY, 1, 1) = "N") Then

              MSComm1.Output = Chr$(HHEsc) & Chr$(HH_HomeU) & Chr$(HHEsc) & Chr$(HHClear)
              MSComm1.Output = "*TNFAULTQ*"
              Call Q_faults.FaultEnterCommand_Click
              Call Sleep(1000)
               
           
End If


GoTo line_70:

end_it_all:
End Sub

Sub retry_tcpip_initialise()


Debug.Print "Entered retry_tcpip_initialize "


' Make the new sockets, as socket(0) already exists
'For i = 1 To Max_Socket
   'Load Socket1(i)
'Next i

'Re- Initialise incoming sockets
'For i = 0 To Max_Socket
    'Socket1(i).LocalPort = PORT_1 + i 'IPPORT_ECHO
    'Socket1(i).AddressFamily = AF_INET
    'Socket1(i).Protocol = IPPROTO_IP
    'Socket1(i).SocketType = SOCK_STREAM
    'Socket1(i).Binary = False
    'Socket1(i).BufferSize = 32767
    'Socket1(i).Blocking = False
    'Socket1(i).Action = SOCKET_LISTEN
    'counter(i) = 0
'Next i




'if at first you don't succeed then try and try again ....

Socket2.HostAddress = HOST
Socket2.AddressFamily = AF_INET
Socket2.Protocol = IPPROTO_IP
Socket2.SocketType = SOCK_STREAM
Socket2.Blocking = False
Socket2.Binary = False
Socket2.BufferSize = 32767
If PCP_Executive Then
     Socket2.RemotePort = PORT_0
Else
     Socket2.RemotePort = 2005
End If

'IPPORT_ECHO ' - 1024 and below are reserved for the system
Socket2.Action = SOCKET_CONNECT

'Connection_Status = "Connected OK"
'frm_exec.StatusBar1.Panels.Item(3).Text = Connection_Status
End Sub

Sub SendMessage(TextString As String)

If Not test_flag Then
    frm_exec.Socket2.SendLen = Len(TextString)
    frm_exec.Socket2.SendData = TextString
    Debug.Print "Sent " & TextString
End If
End Sub

Sub tcpip_close()
Dim i As Integer

' Close all of the RX sockets
For i = 0 To Max_Socket
    If Socket1(i).Connected Then Socket1(i).Action = SOCKET_CLOSE
Next i

' Close the Tx  socket
If Socket2Listening Then Socket2.Action = SOCKET_CLOSE

Debug.Print " Unloading and Exiting : Entered tcpip_close event "


End Sub

Sub tcpip_initialise()
Dim i As Integer

Debug.Print "Entered tcpip_initialise"
' Make the new sockets, as socket(0) already exists
For i = 1 To Max_Socket
    Load Socket1(i)
Next i

' Initialise incoming sockets
For i = 0 To Max_Socket
    If PCP_Executive Then
        Socket1(i).LocalPort = PORT_1 + i ' For PCP operation
    Else
        Socket1(i).LocalPort = PORT_1 + i + 2 ' For SCP operation
    End If
    Socket1(i).AddressFamily = AF_INET
    Socket1(i).Protocol = IPPROTO_IP
    Socket1(i).SocketType = SOCK_STREAM
    Socket1(i).Binary = False
    Socket1(i).BufferSize = 32767
    Socket1(i).Blocking = False
    Socket1(i).Action = SOCKET_LISTEN
    counter(i) = 0
Next i

' Initialise outgoing socket

'On Error GoTo ErrorHandler  ' Enable error-handling routine.
Socket2.HostAddress = HOST
Socket2.AddressFamily = AF_INET
Socket2.Protocol = IPPROTO_IP
Socket2.SocketType = SOCK_STREAM
'Socket2.Protocol = IPPROTO_UDP
'Socket2.SocketType = SOCK_DGRAM
Socket2.Blocking = False
Socket2.Binary = False
Socket2.BufferSize = 32767
If PCP_Executive Then
     Socket2.RemotePort = PORT_0
Else
     Socket2.RemotePort = 2005
End If
Socket2.Action = SOCKET_CONNECT

'Connection_Status = "Connected OK"
'frm_exec.StatusBar1.Panels.Item(3).Text = Connection_Status

Exit Sub ' and avoid error handler


'ErrorHandler:   ' Error-handling routine.
'Msg = "Attempting to connect to host "
'Call Sleep(1000)
'Resume  ' Resume execution at same line that caused the error.


End Sub



Private Sub GenOp_Click()

End Sub

Private Sub AItest_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AI_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "AI Test ON"
End If

 frm_exec.StatusBar1.Panels.Item(5).Text = "AI Test ON"

        Load DI_AI_Test
        DI_AI_Test.Show

End Sub

Private Sub AnalogOutputTest_Click()
Load AOTest
AOTest.Show
End Sub

Private Sub BreakerTest_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   DO_Test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "Breaker Test ON"
End If
'===  21/4/97    *** DO TEST ***
'If Rx_strings(1) = "DO_TEST" Then

'     itest_type = 4  ' For DO test

'     itest_no = Val(Rx_strings(2))
'    itest_no = 0    ' Breaker test (HMS Sultan)
'    itest_no = 1    ' Lamp test    (HMS Sultan)
'    itest_no = 2    ' Restore Breakers and Lamps (after Breaker or Lamp Test)
'    itest_no = 3    ' Individual DO test (HMS Sultan)
'    itest_no = 4    ' Individual DO test (General e.g. TEST-BOX)

 '   ichannel = Val(Rx_strings(3))
 '     IVALUE = 0
            
'    Call TPTEST(itest_type, itest_no, ichannel, IVALUE)
'End If

    MessageString = "#DO_TEST#" & "0#0#0#"
Call SendMessage(MessageString)

    ans = MsgBox("Breaker Test ON  . Click OK  to clear", 0, " *** BreakerTest ***")

    MessageString = "#DO_TEST#" & "2#0#0#"     ' To RESTORE BREAKER STATUS
Call SendMessage(MessageString)

HWtest = False
DO_Test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub DItest_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   DI_Test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "DI Test ON"
End If


' 25/4/97
   MessageString = "#DI_TEST#" & "0#0#0#"
Call SendMessage(MessageString)

'Call fill_IOtext

' 21/4/97 Clear the DIchanges array
    MessageString = "#CLEAR_DI_CHANGES#"
Call SendMessage(MessageString)
Call Sleep(1000)

' Use the DI_AI_Test form to display the DIchanges
        Load DI_AI_Test
        DI_AI_Test.Show


' ( THIS IS THE HOST CODE comented out , shown for reference only)
'===  21/4/97   TO STOP ALL I/O TESTS
'If Rx_strings(1) = "STOP_IO_TESTS" Then
'
'    itest_type = 5  ' To STOP I/O tests
 '     itest_no = 0
'      ichannel = 0
'        IVALUE = 0
'
'       Call TPTEST(itest_type, itest_no, ichannel, IVALUE)
'End If

End Sub

Private Sub DOIndividual_Click()
'Dim ichannel As Integer, ivalue As Integer

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   DO_Test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "DOTest ON"
End If

' Function : Toggles the state of the DO
try_DO_channel:
    ichannel = Val(InputBox("Enter channel number (0 to 5119)", "DO Individual Test"))
If ichannel < 0 Or ichannel > 5119 Then GoTo try_DO_channel:


' 22/4/97
 '   MessageString = "#DO_TEST#" & "3#" & ichannel & "#"          ' HMS Sultan DO Test
     MessageString = "#DO_TEST#" & "4#" & ichannel & "#"          ' General DO Test
Call SendMessage(MessageString)

HWtest = False
DO_Test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub GenCharacteristics_Click()
' 4/4/97
frm_genop.Show

End Sub

Private Sub Initialise_Click()

Static i As Integer
Dim msg As String

'Dim MessageString As String
'Check if tcpip initialized ok


'6/5/97 warning message if already initialized once before..

            If SimInitialized Then
                    msg = MsgBox("Initializing will stop the Simulator and return to a dead-board state." _
                    & " Are you sure you want to continue ? ", 52, "Simulator Initialize")
                    If msg <> vbYes Then
                        Exit Sub
                    End If
            End If



                ' send info on simulator configuration, instructor name etc


               'Split or Normal Configuration
                MessageString = "#INITLZ#" & SimConfiguration & "#"
                Call SendMessage(MessageString)

                ' 1/4/97 changed 2 secs to 200mS
                Sleep (200)

'  7/5/97  Don't need now
                'Start the host's tasks running
  '              MessageString = "#INIT_TASKS#"
 '               Call SendMessage(MessageString)
'                Sleep (200)



'                'Send Instructor info
              If PCP_Executive Then
                    MessageString = "#TPEXEC#" & InstructorName & "#" & ClassName & "#" & CourseName & "#"
                Else
                    MessageString = "#TSEXEC#" & InstructorName & "#" & ClassName & "#" & CourseName & "#"
                End If
                Call SendMessage(MessageString)
                Sleep (200)

                If PCP_Executive Then
                    MessageString = "#TPRUN#"
                Else
                     MessageString = "#TSRUN#"
                End If
                Call SendMessage(MessageString)
                Sleep (200)
                HeaderPrinted = False               'print log header if log print requested - see logging.frm

                'SimInitialized = True
                Call Sleep(200)
    '7/5/97 dont send di changes on initialize
                MessageString = "#CLEAR_DI_CHANGES#" & Str$("3071") & "#"
                Call SendMessage(MessageString)
                Call Sleep(500)
    '7/5/97
                Load Logging
                Logging.Hide
                SimInitialized = True

    
End Sub

Private Sub LampTest_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   DO_Test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "Lamp Test ON"
End If

' 21/4/97
    MessageString = "#DO_TEST#" & "1#0#0#"
Call SendMessage(MessageString)

    ans = MsgBox("Lamp Test ON  . Click OK  to clear", 0, " *** Lamp Test ***")

    MessageString = "#DO_TEST#" & "2#0#0#"     ' To RESTORE LAMP STATUS
Call SendMessage(MessageString)

HWtest = False
DO_Test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub LoadShare_Click()
' 9/4/97
'Load frm_LoadShare
' 4/4/97
frm_LoadShare.Show

End Sub

Private Sub MDIForm_Load()
'Dim MessageString As String



'  12/5/97
DI363 = False
DI363_count = 0
HostAvailable = True            'assume host is OK by default, intro timer will set it false if no comms available

' 1/5/97
If PCP_Executive Then
    mnu_Panel_A.Caption = "PCP Panel"
    mnu_panel_B.Caption = ""
Else
    mnu_Panel_A.Caption = "SCP-H Panel"
    mnu_panel_B.Caption = "SCP-J Panel"
End If




' 1/4/97
'Me.StatusBar1.Panels.Item(1).Text = Op_Mode_Name

' move to main page  as a test ..   tcpip_initialise
'Sleep (5000)

'   17/3/97
'  Use COM1  to communicate to the Hand-Held VDU
    MSComm1.CommPort = 1
' 9600 baud, no parity, 8 data, and 1 stop bit.
    MSComm1.Settings = "9600,N,8,1"
     
 ' Tell the control to read entire buffer when Input is used.
    MSComm1.InputLen = 0
  
' Codes for the HH VDU
HHEsc = 27
HHLightOn = 81
HHLightOff = 82
HHClear = 74
HH_HomeU = 76
HH_HomeD = 72
ClearText = "                                                  "

Call fill_faultext
Call fill_IOtext
Call FillAIText

If Not PCP_Executive Then
    frm_exec.Caption = "T42 Simulator SCP Station"
End If


'load the logging form as this can start immediately (before any request to show or print)
'Load Logging

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    tcpip_close
End Sub


Private Sub MeterFSD_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "AO Test ON"
End If

' 21/4/97
   MessageString = "#AO_TEST#" & "2#"
Call SendMessage(MessageString)

    ans = MsgBox("Meter HALF test ON  . Click OK  to clear", 0, " Full Scale Meter Test ")

HWtest = False
AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub MeterHalf_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "AO Test ON"
End If

  ' MessageString = "#AO_TEST#" & "1"
   MessageString = "#AO_TEST#" & "1" & "#"
  Call SendMessage(MessageString)

    ans = MsgBox("Meter HALF test ON  . Click OK  to clear", 0, " Half Scale Meter Test ")

HWtest = False
AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub MeterIndividual_Click()
Dim ichannel As Integer, ivalue As Integer

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "Meter INDIVIDUAL Test"
End If

try_channel:
    ichannel = Val(InputBox("Enter channel number (0 to 767)", "Meter Individual Test"))
If ichannel < 0 Or ichannel > 767 Then GoTo try_channel:

try_value:
   ivalue = Val(InputBox("Enter value (0 to 255)", "Meter Individual Test"))
If ivalue < 0 Or ivalue > 255 Then GoTo try_value:
  
' 21/4/97
   MessageString = "#AO_TEST#" & "3#" & Str$(ichannel) & "#" & Str$(ivalue) & "#"
Call SendMessage(MessageString)


    HWtest = False
   AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub MeterZero_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "Meter ZERO Test"
End If
'===  21/4/97    *** AO TEST ***
'If Rx_strings(1) = "AO_TEST" Then

'     itest_type = 2  ' For AO test

'     itest_no = Val(Rx_strings(2))
'    itest_no = 0    ' Meters ZERO         (HMS Sultan)
'    itest_no = 1    ' Meters 1/2 FSD      (HMS Sultan)
'    itest_no = 2    ' Meters FSD          (HMS Sultan)
'    itest_no = 3    ' Meters (Individual)
'    itest_no = 4    ' PCP   Synchro-scope (HMS Sultan)
'    itest_no = 5    ' SCP-H Synchro-scope (HMS Sultan)
'    itest_no = 6    ' SCP-M Synchro-scope (HMS Sultan)

'     ichannel = 0
'     IVALUE = 0
            
'    Call TPTEST(itest_type, itest_no, ichannel, IVALUE)
'End If

    MessageString = "#AO_TEST#" & "0#"
Call SendMessage(MessageString)

    ans = MsgBox("Meter ZERO test ON  . Click OK  to clear", 0, "Meter ZERO Test ")

HWtest = False
AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub mnu_all_ao_Click()


End Sub

Private Sub mnu_data_record_Click()
Dim SnapCommand As String

If Not SPLITMODE Then

    If RecordingOn Then
        msg = MsgBox("Stop Recording Now ? ", vbOKCancel + vbQuestion, "Data Record and Replay")
        If msg = vbCancel Then
            Exit Sub
        End If
        MessageString = "#TPRECORD#" & "STOP#" & "0#" & "0#"         'tell model to stop recording
        RecordingOn = False
        'frm_exec.StatusBar1.Panels.Item(6).Text = "Recording OFF"
        mnu_data_record.Caption = "Start R&Ecord"
        mnu_data_replay.Enabled = True
        ReplayAvailable = True
        
    Else
        msg = MsgBox("Start Recording Now ? ", vbOKCancel + vbQuestion, "Data Record and Replay")
        If msg = vbCancel Then
            Exit Sub
        End If
        mnu_data_record.Caption = "End R&Ecord"
        'take snapshot #100 first
       SnapCommand = "#TPSNAP#"
        MessageString = SnapCommand & Str$(10) & "#"
        Call SendMessage(MessageString)
        
        'wait for data to be recorded
        Load SnapProgressForm
        SnapProgressForm.Show
        SnapProgressForm.text1.Text = "Saving current state "
        SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Min
        SnapProgressForm.ProgressBar1.Visible = True
        IncrementVal = (SnapProgressForm.ProgressBar1.Max - SnapProgressForm.ProgressBar1.Min) / 20
        For i = 1 To 20
            SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Value + IncrementVal
            Select Case i
                Case 1
                    SnapProgressForm.Text2.Text = "Opening File"
                Case 5
                    SnapProgressForm.Text2.Text = "Saving Data"
                Case 16
                    SnapProgressForm.Text2.Text = "Closing File"
            End Select
            DoEvents
            Call Sleep(500)
        Next
        Unload SnapProgressForm
        
        'tell model to start recording
        MessageString = "#TPRECORD#" & "START#" & "0#" & "0#"
       ' frm_exec.StatusBar1.Panels.Item(6).Text = "Recording ON"
        RecordingOn = True
        
    End If
    
    
        'Call SendMessage(MessageString) to start or end recording
       Call SendMessage(MessageString)
Else
     msg = MsgBox("Record and Replay functions are not available in split mode ", vbOK + vbInformation, "Data Record and Replay")
End If
End Sub

Private Sub mnu_data_replay_Click()
Dim SnapCommand As String


If Not SPLITMODE Then

    If ReplayOn Then
        msg = MsgBox("Stop Replay Now ? ", vbOKCancel + vbQuestion, "Data Record and Replay")
        
         If msg = vbCancel Then
            Exit Sub
        End If
        MessageString = "#TPREPLAY#" & "STOP#" & "0#" & "0#"
        ReplayOn = False
       ' frm_exec.StatusBar1.Panels.Item(6).Text = "Replay OFF"
        mnu_data_replay.Caption = "Start Re&Play"
    Else
        msg = MsgBox("Start Replay Now ? ", vbOKCancel + vbQuestion, "Data Record and Replay")
        If msg = vbCancel Then
            Exit Sub
        End If
        mnu_data_replay.Caption = "End Re&Play"
        'first load the replay snapshot
        If Not test_flag Then
                SnapCommand = "#TPRELSNAP#"
                MessageString = SnapCommand & Str$(10) & "#"                   'Number 100 is a special snaphot reserved for record/replay
                Call SendMessage(MessageString)
        End If
        Load SnapProgressForm
        SnapProgressForm.Show
        SnapProgressForm.text1.Text = "Reloading Recording "
        SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Min
        SnapProgressForm.ProgressBar1.Visible = True
        IncrementVal = (SnapProgressForm.ProgressBar1.Max - SnapProgressForm.ProgressBar1.Min) / 20
        For i = 1 To 20
            SnapProgressForm.ProgressBar1.Value = SnapProgressForm.ProgressBar1.Value + IncrementVal
            Select Case i
                Case 1
                    SnapProgressForm.Text2.Text = "Opening File"
                Case 5
                    SnapProgressForm.Text2.Text = "Loading Data"
                Case 16
                    SnapProgressForm.Text2.Text = "Closing File"
            End Select
            DoEvents
            Call Sleep(500)
        Next
        
        
        Unload SnapProgressForm
        MessageString = "#TPREPLAY#" & "START#" & "0#" & "0#"
        ReplayOn = True
       ' frm_exec.StatusBar1.Panels.Item(6).Text = "Replay ON"

    End If
       
   
     Call SendMessage(MessageString)
Else
     msg = MsgBox("Record and Replay functions are not available in split mode ", vbOK + vbInformation, "Data Record and Replay")
End If
End Sub

Private Sub mnu_exit_Click()
Unload Me
End Sub

Private Sub mnu_faluts_entry_Click()

' 12/2/97
'Call fill_faultext
frm_faults.Show

End Sub

Private Sub mnu_faults_queued_Click()
Q_faults.Show
End Sub

Private Sub mnu_g_cvp_Click()
    frm_CV.Show
End Sub


Private Sub mnu_g_offline_cvp_Click()
frm_offline_CV.Show
End Sub

Private Sub mnu_g_offline_rr_Click()
frm_offline_RR.Show
End Sub

Private Sub mnu_g_rr_Click()
frm_RR.Show
End Sub


Private Sub mnu_g_status_Click()

' RWJ 23/12/96 added:
' Show the status form
frm_t42_values.Show

End Sub

Private Sub mnu_g_synch_phasors_Click()
frm_synch.Show
End Sub


Private Sub mnu_instrp_gen_m1_Click()

End Sub

Private Sub mnu_instrp_ovpu_Click()

End Sub

Private Sub mnu_instruct_params_Click()

InstrParam.Show

End Sub

Private Sub mnu_log_view_Click()
Logging.WindowState = 0
Logging.Visible = True
'Logging.Show

End Sub

Private Sub mnu_mobile_fault_control_Click()


'   JBF Tidied up the VB equivalent of TPMMU.F90 code
'   He made a new routine called mfu_control
'   However we now need to modify it for SCP in SPLIT

     Call mfu_control


Exit Sub



End Sub

Private Sub mnu_PCP_panel_Click()

' 10/2/97
PCP_panel.Show

End Sub

Private Sub mnu_next_fault_control_Click()
Dim msg As String

'all written 6/5/97
 Q_faults.Hide
                    
If NoOfFaultsInQueue < 1 Then
    msg = MsgBox("The fault Queue is empty. Do you want to add some ?")
    If msg = vbYes Then
        'Hide Me
        Q_faults.Show
    Else
        Exit Sub
    End If
Else
    Call Q_faults.FaultEnterCommand_Click
End If

End Sub

Private Sub mnu_Panel_A_Click()

' 1/5/97
If PCP_Executive Then
         PCP_panel.Show
Else
       SCPH_panel.Show
End If

End Sub

Private Sub mnu_panel_B_Click()

' 1/5/97
If PCP_Executive Then
         PCP_panel.Show
Else
       SCPJ_panel.Show
End If

End Sub


Private Sub mnu_power_distribution_Click()

' 10/2/97
frm_power.Show

End Sub

Private Sub mnu_sim_halt_Click()

If PCP_Executive Then
      MessageString = "#TPHALT#"
Else
      MessageString = "#TSHALT#"
End If
Call SendMessage(MessageString)

End Sub

Private Sub mnu_sim_reset_Click()
Dim msg As String

'3/4/97 added warning

msg = MsgBox("Resetting will stop the Simulator and return to a dead-board state." _
& " Are you sure you want to continue ? ", 52, "Simulator Reset")
If msg = vbYes Then

' 1/4/97 added
    MessageString = "#END_TASKS#"
    Call SendMessage(MessageString)
 '   Unload Me
    Load frm_intro
    frm_intro.Visible = True
    
' 1/4/97
'    Unload Me
    frm_exec.Visible = False
    frm_intro.Visible = True

Else
    Exit Sub
End If

'6/5/97
SimInitialized = False

End Sub

Private Sub mnu_sim_resume_Click()
' 1/4/97
If PCP_Executive Then
    MessageString = "#TPRESUME#"
Else
    MessageString = "#TSRESUME#"
End If
Call SendMessage(MessageString)
Sleep (200)
End Sub

Private Sub mnu_sim_run_Click()

' == 1/4/97 Code moved to Initialise_click

'Static i As Integer

''Split or Normal Configuration
'MessageString = "#INITLZ#" & SimConfiguration & "#"
'Call SendMessage(MessageString)

'' 1/4/97 change 2 secs to 200mS
'Sleep (200)

'Start the host's tasks running
'MessageString = "#INIT_TASKS#"
'Call SendMessage(MessageString)
'Sleep (200)

'Send Instructor info
'MessageString = "#TPEXEC#" & InstructorName & "#" & ClassName & "#" & CourseName & "#"
'Call SendMessage(MessageString)
'Sleep (200)

'MessageString = "#TPRUN#"
'Call SendMessage(MessageString)
'Sleep (200)



End Sub



Private Sub mnu_snapshot_load_Click()
TakeSnapFlag = False
Load SnapForm

End Sub

Private Sub mnu_snapshot_save_Click()
TakeSnapFlag = True
Load SnapForm

End Sub

Private Sub mnu_tcpip_array_Click()
    frm_tcpip.Show
End Sub



Private Sub PCPsynchro_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "PCP Synchro Test"
End If


' 21/4/97
   MessageString = "#AO_TEST#" & "4#"
Call SendMessage(MessageString)

    ans = MsgBox("PCP Synchro test ON  . Click OK  to clear", 0, " *** PCP Synchro Test ***")

HWtest = False
AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""

End Sub

Private Sub Resume_Click()

' 1/4/97
If PCP_Executive Then
    MessageString = "#TPRESUME#"
Else
    MessageString = "#TSRESUME#"
End If
Call SendMessage(MessageString)
Sleep (200)

End Sub

Private Sub SCPHSynchro_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "SCP-H Synchro Test"
End If


' 21/4/97
   MessageString = "#AO_TEST#" & "5#"
Call SendMessage(MessageString)

    ans = MsgBox("SCP-H Synchro test ON  . Click OK  to clear", 0, " *** SCP-H Synchro Test ***")

HWtest = False
AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""


End Sub

Private Sub SCPMsynchro_Click()

If HWtest Then
    ans = MsgBox("Only ONE HW Test is allowed at any one time", 0, " *** HW Test ***")
    Exit Sub
Else
    HWtest = True
   AO_test = True
   frm_exec.StatusBar1.Panels.Item(5).Text = "SCP-M Synchro Test"
End If


' 21/4/97
   MessageString = "#AO_TEST#" & "6#"
Call SendMessage(MessageString)

    ans = MsgBox("SCP-M Synchro test ON  . Click OK  to clear", 0, " *** SCP-M Synchro Test ***")

HWtest = False
AO_test = False
frm_exec.StatusBar1.Panels.Item(5).Text = ""


End Sub

Private Sub Socket1_Accept(Index As Integer, SocketId As Integer)
Dim i(0 To 127) As Integer

Debug.Print "Entered Socket1(i) Accept,   Index = " & Index
Socket1(Index).Action = SOCKET_ACCEPT
' Socket1(Index).Accept = SocketId
End Sub

Private Sub Socket1_Connect(Index As Integer)
'RichTextBox1.Enabled = True
'Command2.Enabled = True

Debug.Print "Socket1 Connected, Index = " & Index
Connection_Status = "Rx Connected"

' 22/4/97
'frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

End Sub


Private Sub Socket1_Disconnect(Index As Integer)
Dim istatus(0 To 127)  As Integer

' Entered when any Socket1is disconnected. This would happen as a result of the executive client
' sending a RETRY command to the host causing the host's Tx sockets to be closed then re-opened OR
' the host shutting down.
' This routine effectively re-initialises the listening sockets so the host can re-connect.
' Re-initialization is by CLOSING, waiting 3 secs and then OPENING again.
    

' First put message on status bar
    Connection_Status = "Rx Disconnected"
    
' 22/4/97
'    frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

'Close the listening (Rx) socket which has just been disconnected
    Socket1(Index).Action = SOCKET_CLOSE
    istatus(Index) = Socket1(Index).State
    Debug.Print "Entered Socket1 - Disconnect  Status of Socket " & Index & "  = " & istatus(Index)
    
'Wait then re-open so host can re-connect at some time to fresh socket(s)
    Call Sleep(3000)
    Socket1(Index).Action = SOCKET_OPEN
    Socket1(Index).Action = SOCKET_LISTEN
    istatus(Index) = Socket1(Index).State
    Debug.Print "CLOSED and then OPENED Socket1(" & Index & ")   Status of Socket    = " & istatus(Index)
    Connection_Status = "Rx Waiting"
    
'  22/4/97
'     frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

End Sub


Private Sub Socket1_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
'Debug.Print "Socket1 - LastError : " & ErrorString


MsgBox "Socket 1(" & Index & "), LastError: " & ErrorString
End Sub


Private Sub Socket1_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
Dim recv As Variant
Dim n
Dim i As Integer
Dim StartPosition As Integer, EndPosition As Integer
Dim msg As String
Dim RejectIntegerString As Integer
Dim RejectRealString As Integer
Dim RxLength As Integer
Dim TxLength As Integer


'On Error GoTo ErrorHandler

Socket1(Index).RecvLen = DataLength
RxLength = DataLength
'Debug.Print ("socket1 read , data length = : " & DataLength & "      Index = " & Index)
recv = Socket1(Index).RecvData
TxLength = Val(Mid(recv, 2, 5))
'Debug.Print "RxLength = " & RxLength & "       TxLength = " & TxLength

If RxLength = TxLength + 1 Then

'If DataLength < 1200 Then
    'Debug.Print ("socket1 read, recv =  : " & recv)
    '   Debug.Print " "
    'Array variables are received as a continuous character string with "#" delimiters between each element
    'This is done to minimise tcpip overhead when Tx byte at a time

    ' 30/4/97 move up      On Error GoTo ErrorHandler

    Select Case Index
    Case 0

    'If Not RejectRealString Then                'reject first string < 1200 long after bad batch because this too could be bad
        '  NOTE:  The following method is very inefficient  ( CPU IDLE time = 72% to 79% )
        '  So the  line [T42DIESELR(counter(Index)) = recv ]  is commented out and only include if HOST is transmitting a small string at a time
        '  assume number of varaibles (for now) is 200
    
        'For Array T42DIESELR
        'Get the Start and End position of each command within the string, extract and convert to a real variable
        
        '30/4/97 StartPosition = 1
        StartPosition = 7

        For i = 0 To 200
            EndPosition = InStr(StartPosition + 1, recv, "#", 1)
        
            If EndPosition - StartPosition - 1 > 0 Then
                PC_REALS(i) = Mid(recv, StartPosition + 1, EndPosition - StartPosition - 1)
            Else
                Debug.Print " ** Rejecting data in Socket 1(0) - bad variable length **" 'data is invalid so discard
            
            End If
            StartPosition = EndPosition
        
        Next i
    'Else
        'RejectRealString = False
   ' End If
    
    
    ' ====22/1/97 Tried to receive BOTH arrays in SOCKET1(0)  =====     TESTED WORKS OK ==========
    '
    '  However the penalty for using only ONE Rx Socket is a reduction in IDLE time from 89% to 82%
    '  Note : The HOST program remained idling for 97% of the time (i.e. NO penalty)
    '  So the following code was removed (removed from HOST program as well)
    '
    '    For i = 0 To 200
    '        EndPosition = InStr(StartPosition + 1, recv, "#", 1)
    '        T42BUSR(i) = Mid(recv, StartPosition + 1, EndPosition - StartPosition - 1)
    '        StartPosition = EndPosition
    '    Next i
    ' ===============================================================================

    Case 1
    
    
    
    
    'If Not RejectIntegerString Then
    
        'For Array T42BUSR
        'T42BUSR(counter(Index)) = recv                                  ' only include if HOST is transmitting a small string at a time
        'Get the Start and End position of each command within the string, extract and convert to a real variable
        
        
        
        '30/4/97 StartPosition = 1
        StartPosition = 7           'preceeded by #xxxxx# where xxxxx is string length

        For i = 0 To 200
            EndPosition = InStr(StartPosition + 1, recv, "#", 1)
            If EndPosition - StartPosition - 1 > 0 Then                     'data is invalid so discard
            PC_INTS(i) = Mid(recv, StartPosition + 1, EndPosition - StartPosition - 1)
            Else
                Debug.Print " ** Rejecting data in Socket 1(1) - bad integer length **"
            End If
            StartPosition = EndPosition
        Next i
        'Debug.Print " received PC_INTS                 PC_INTS(69-72) = " & PC_INTS(69) & "  " & PC_INTS(70) & "  " & PC_INTS(71)
  'Else
        'RejectRealString = False
    'End If
    End Select
    
Else                        'set flags if a bad length string is received and  dont do anything as this indicates Tx/ Rx error

Socket1(Index).Flush
Debug.Print "flushing buffers on Socket 1 (" & Index & ")"
    'Select Case Index
        'Case 0
             'RejectRealString = True
              'Debug.Print " Setting Reject Real String **"
        'Case 1
            'RejectIntegerString = True
            'Debug.Print " Setting Reject Integer String **"

        'End Select
        
End If

'only include these lines if HOST is transmitting a Byte at a time
'counter(Index) = counter(Index) + 1
'If counter(Index) > array_size Then counter(Index) = 0
Exit Sub

ErrorHandler:
msg = MsgBox("Error Reading TCP/IP Socket 1:  Index =  " & Index & "Error = " & ERROR, vbOK + vbCritical, "Error Report")
Debug.Print " "
Debug.Print "Socket1 Read, last data length = " & DataLength
Debug.Print "Socket1 Read, last recv = " & Mid(recv, 1, 50)
Exit Sub


End Sub


Sub tcpip_send_array_single(Socket_No As Integer, ByRef l_array() As Single)
    ' Send an array of singles via Socket1
    Dim i As Integer
    Dim ch As String
    
    For i = 0 To array_size
        ch = Str$(l_array(i))
        Socket1(Socket_No).SendLen = Len(ch + Chr$(10))
        Socket1(Socket_No).SendData = ch + Chr$(10)
        Socket1(Socket_No).Action = SOCKET_FLUSH
    Next i

End Sub

Private Sub Socket2_Accept(SocketId As Integer)
Debug.Print "Socket2 Accept "
End Sub

Private Sub Socket2_Blocking(status As Integer, Cancel As Integer)
Debug.Print "Socket2 Blocking "

End Sub


Private Sub Socket2_Cancel(status As Integer, Response As Integer)
Debug.Print "Socket2 Cancel "

End Sub


Private Sub Socket2_Connect()

'When Tx connection established with host, put message on Status Bar

Connection_Status = "Tx Connected"
Debug.Print " Socket2.Connect - Connected OK"

'   22/4/97
'    frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

End Sub

Private Sub Socket2_Disconnect()

'If host is stopped and the Tx socket (2) is disconnected, then try to reconnect every 5 seconds

'first put message on Status Bar

Connection_Status = "Tx Disconnected"

' 22/4/97
'    frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

Debug.Print "Socket2 CLOSED in Socket2.Disconnect calling Retry_tcpip_initialize"

'Close the Tx socket and wait before re-initializing
Socket2.Action = SOCKET_CLOSE
Call Sleep(5000)
Call retry_tcpip_initialise

End Sub


Private Sub Socket2_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

'If host is stopped and the Tx socket (2) is disconnected, then try to reconnect every 5 seconds

'first put message on Status Bar
Connection_Status = "Tx Disconnected"

' 22/4/97
'    frm_exec.StatusBar1.Panels.Item(5).Text = Connection_Status

'Close the Tx socket and wait before re-initializing
Socket2.Action = SOCKET_CLOSE
Debug.Print "Socket2 CLOSED in Socket2.Last Error calling Retry_tcpip_initialize"
Call Sleep(5000)

Call retry_tcpip_initialise
End Sub


Private Sub Socket2_Write()
Debug.Print " sent message written from Socket2.Read"

End Sub




Private Sub Timer1_Timer()
' 7/4/97
Dim offset As Integer

' 12/5/97
Dim exec_string As String * 10
Dim msg As String
Dim ans As Integer


If HOST = HOST_A Then
    Me.StatusBar1.Panels.Item(7).Text = "Host A"
Else
    Me.StatusBar1.Panels.Item(7).Text = "Host B"
End If

'On Error GoTo ErrorHandler
'  12/5/97
'If PCP_Executive And AcceptButtonPressed Then
If PCP_Executive And AcceptButtonPressed Then


       DI363 = PC_INTS(167)
       'If DI363 Then
           'Debug.Print "DI363 = true"
        'Else
           'Debug.Print "DI363 = false"
       'End If
        
       '13/5/97 If Not DI363 Then
       If (DI363 Or Not HostAvailable) Then

            Debug.Print "inside the not DI363 loop"
            DI363_count = DI363_count + 1
      End If

    If DI363_count > 5 Then
         DI363_count = 0
     
           If HOST = HOST_A Then
                      If Not HostAvailable Then
                                msg = "The computer Host A is not responding. Ensure it is switched on. " & Chr(10) & Chr(13) & _
                                  "Otherwise do you wish to use Host B instead? "
                      Else
                                msg = "The host computer A does not appear to be connected to the simulator. " & Chr(10) & Chr(13) & _
                                  "Do you wish to reconfigure to use Host B instead ? "
                      End If
                      
                     Mid(exec_string, 4, 7) = "#HOST_B"
           Else
                     If Not HostAvailable Then
                                msg = "The computer Host B is not responding. Ensure it is switched on. " & Chr(10) & Chr(13) & _
                                  "Otherwise do you wish to use Host A instead? "
                      Else
                                msg = "The host computer B does not appear to be connected to the simulator. " & Chr(10) & Chr(13) & _
                                  "Do you wish to reconfigure to use Host A instead ? "
                      End If
                     Mid(exec_string, 4, 7) = "#HOST_A"
            End If
            ans = MsgBox(msg, 36, "Host Computer Selection")
            If ans = vbYes Then
loop1:             msg = "Please confirm that the SCP computer is on and running Windows before continuing. "
                        ans = MsgBox(msg, 33, "Host Computer Selection")
                        If ans = vbOK Then
                                Mid(exec_string, 1, 3) = "SCP"
                                Open "\\SCP\sultan\sultan.cfg" For Output As #1
                                Print #1, exec_string
                                Debug.Print "exec_string = " & exec_string
                                Close #1
                                Mid(exec_string, 1, 3) = "PCP"
                                Open "C:\sultan\sultan.cfg" For Output As #1
                                Print #1, exec_string
                                Debug.Print "exec_string = " & exec_string
                                Close #1
                                msg = "The PCP and SCP computers have been reconfigured to work with " & Mid(exec_string, 5, 6) & Chr(10) & Chr(13) & _
                                        "If the SCP Simulator Program is running exit back  to windows and run it again" & Chr(10) & Chr(13) & _
                                        "This  PCP Simulator Program will now exit. Please Run it again in the normal way"
                                        ans = MsgBox(msg, 32, "Host computer selection")
                                        End
                       Else
                             msg = "It is essential that the SCP computer is running at this stage in order to complete " & Chr(10) & Chr(13) & _
                                        "the configuration. If this is not possible, then the SCP computer cannot be used. Please confirm " & Chr(10) & Chr(13) & _
                                        "the  SCP computer is on and running Windows"
                                        ans = MsgBox(msg, 32, "Host computer selection")
                              GoTo loop1
                       End If
                        
            End If
    End If
            
End If


If Not test_flag Then
    SPLITMODE = PC_INTS(89)
Else
'Splitmode flag set in intro form
End If

T42RUN = PC_INTS(90)
PCPRUN = PC_INTS(91)
SCPRUN = PC_INTS(92)


   If SPLITMODE Then
   
      Me.StatusBar1.Panels.Item(1).Text = "Split mode"
        If PCP_Executive Then
             offset = 0
        Else
             offset = 6
        End If
 ' 9/4/97  Find out which alternators are connected in parallel
           No1Alt = PC_INTS(95)
           No2Alt = PC_INTS(96)
       
  Else
  
      Me.StatusBar1.Panels.Item(1).Text = "Normal mode"
            offset = 0
  ' 9/4/97  Find out which alternators are connected in parallel
           No1Alt = PC_INTS(93)
           No2Alt = PC_INTS(94)
          
  End If
  
  If T42RUN Then
           Me.StatusBar1.Panels.Item(2).Text = "T42 Running"
  Else
            Me.StatusBar1.Panels.Item(2).Text = "T42 Halted"
  End If
  If PCPRUN Then
           Me.StatusBar1.Panels.Item(3).Text = "PCP Running"
  Else
            Me.StatusBar1.Panels.Item(3).Text = "PCP Halted"
  End If
  If SCPRUN Then
           Me.StatusBar1.Panels.Item(4).Text = "SCP Running"
  Else
            Me.StatusBar1.Panels.Item(4).Text = "SCP Halted"
  End If
        
                  
 '   18/3/97    Alternator Status
For i = 1 To 6
' 7/4/97   If PC_INTS(16 + i) Then
                If PC_INTS(16 + i + offset) Then
' 24/3/97                    gen_ON(i) = True
                     On_line_gen_on(i) = True
               Else
' 24/3/97                     gen_ON(i) = False
                     On_line_gen_on(i) = False
               End If
Next i

'Enable/Disable run & halt options in menu
If (PCP_Executive And PCPRUN) Or (Not PCP_Executive And SCPRUN) Then
    mnu_sim_run.Enabled = False
Else
    mnu_sim_run.Enabled = True
End If
mnu_sim_halt.Enabled = Not mnu_sim_run.Enabled


'7/5/97 disable all SCP functions except graphics, if we are SCP in combined mode

'9/5/97 If (Not PCP_Executive And Not SPLITMODE) Then
If (PCP_Executive) Then
    If SPLITMODE Then                       ' PCP split, all functions available except replay
        mnu_data_record.Enabled = False
        mnu_data_replay.Enabled = False
     Else                                   ' PCP combined, all functions available
        mnu_data_record.Enabled = True
        mnu_data_replay.Enabled = True
     End If
Else
    If SPLITMODE Then                       ' SCP split, all functions available except record replay
        
        mnu_snapshot_load.Enabled = True
        mnu_snapshot_save.Enabled = True
        mnu_data_record.Enabled = False
        mnu_data_replay.Enabled = False
        mnu_log_actions.Enabled = True
        mnu_simulator.Enabled = True
    Else                                    'SCP combined(not normally used but only allow graphics"
        mnu_snapshot_load.Enabled = False
        mnu_snapshot_save.Enabled = False
        mnu_data_record.Enabled = False
        mnu_data_replay.Enabled = False
        mnu_log_actions.Enabled = False
        mnu_simulator.Enabled = False
    End If
End If


'Debug.Print "faultqlist count = " & Q_faults.FaultQList.ListCount

''For i = 11 To 26
'     PC_REALS(i) = T42SCOPER(i - 10)
'Next i
' 2/5/97
If Synchro_Scope_string = "PCP_synchro" Then

         
' 2/5/97  TESTING:
    If test_flag Then
           VINC = 410#
           FINC = 60.01
         VRUN = 440#
         FRUN = 60#
         LOCK_SCOPE = False
         PARK_SCOPE = False
'          AIPHASE_model = AIPHASE_model + 255# * (FINC - FRUN) * 0.5   '  (  i.e. * timer1.interval)
           AIPHASE_model = AIPHASE_model + 255# * test_speed * 0.5        '  (  i.e. * timer1.interval)
            If AIPHASE_model > 255# Then AIPHASE_model = AIPHASE_model - 255#
            If AIPHASE_model < 0# Then AIPHASE_model = AIPHASE_model + 255#
        
    Else
           VINC = PC_REALS(11)
           FINC = PC_REALS(12)
         VRUN = PC_REALS(13)
         FRUN = PC_REALS(14)
         LOCK_SCOPE = PC_INTS(161)
         PARK_SCOPE = PC_INTS(164)
         AIPHASE_model = PC_INTS(112)      '           AI(13)   (PCP)
    End If
         
ElseIf Synchro_Scope_string = "SCPH_synchro" Then
           VINC = PC_REALS(15)
           FINC = PC_REALS(16)
         VRUN = PC_REALS(17)
         FRUN = PC_REALS(18)
           LOCK_SCOPE = PC_INTS(162)
           PARK_SCOPE = PC_INTS(165)
        AIPHASE_model = PC_INTS(118)      '           AI(19)   (SCP H)
Else
           VINC = PC_REALS(19)
           FINC = PC_REALS(20)
         VRUN = PC_REALS(21)
         FRUN = PC_REALS(22)
        LOCK_SCOPE = PC_INTS(163)
        PARK_SCOPE = PC_INTS(166)
      AIPHASE_model = PC_INTS(115)      '           AI(19)   (SCP M)
End If
     

If Not SimInitialized Then
    mnu_sim_resume.Enabled = False
Else
        mnu_sim_resume.Enabled = True
        Call LogEvents
End If


'add status bar messages for record + replay
If PC_INTS(168) Then        'recording is on
    frm_exec.StatusBar1.Panels.Item(6).Text = "Record ON"
    mnu_data_replay.Enabled = False
    mnu_data_record.Enabled = True
    mnu_data_record.Caption = "End Re&Cord"
    mnu_data_replay.Caption = "Start Re&Play"
ElseIf PC_INTS(169) Then    'replay is on
    'ReplayAvailable = True
     frm_exec.StatusBar1.Panels.Item(6).Text = "Record ON"
     mnu_data_replay.Enabled = True
     mnu_data_record.Enabled = False
     mnu_data_record.Caption = "Start Re&Cord"
     mnu_data_replay.Caption = "End Re&Play"
Else
     frm_exec.StatusBar1.Panels.Item(6).Text = "               "
     mnu_data_replay.Enabled = True
     mnu_data_record.Enabled = True
     mnu_data_record.Caption = "Start Re&Cord"
     mnu_data_replay.Caption = "Start Re&Play"
End If


   

Exit Sub

ErrorHandler:

Debug.Print "Error : " & ERROR

End Sub


