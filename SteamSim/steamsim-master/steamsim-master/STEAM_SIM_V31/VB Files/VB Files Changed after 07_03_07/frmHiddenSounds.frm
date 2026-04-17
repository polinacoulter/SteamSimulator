VERSION 5.00
Object = "{3AD1CF99-F37B-11D0-901E-0020AF7543C2}#5.0#0"; "fxsnd50g.ocx"
Begin VB.Form frmHiddenSounds 
   Caption         =   "Sound Code (Hidden)"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   1  'Minimized
   Begin FXSND50GLib.FXSnd snd 
      Left            =   1140
      Top             =   600
      _Version        =   327680
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Persistence     =   -1  'True
      _StdProps       =   558723
      ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
      ErrCode         =   601885277
      ErrInfo         =   1262060308
      Notify          =   -1  'True
   End
End
Attribute VB_Name = "frmHiddenSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InitialiseSounds()
    Dim i As Integer
    Dim SndPath As String
    Dim SoundPathname As String
    Dim resp As Integer
    
    
    On Error GoTo Errh1
    
    
    
'VOLUME
'The DSVolume property value can be set from 0 (0 dB, no volume adjustment)
'to -10,000 (-100 dB, essentially silent).  Amplification is not currently supported.
'Volume units are hundredths of a decibel, where zero is the original volume of the sound.

' FREQUENCY
'The DSFrequency property can be set to a value between 100 and 100,000, inclusive.
'Setting the DSFrequency to 0 resets the playback frequency to its original value.
'Increasing or decreasing the frequency changes the perceived pitch of the audio sound.
'Setting this property does not change the format of the actual sound data.
    
' PAN (Left & Right)
'Zero is the neutral value and indicates that both speakers are at full volume.
'At any other setting, one of the speakers is at full volume and the other is attenuated.
'For example, a pan of -2173 means that the left speaker is at full volume and the right
'speaker is attenuated by 21.73 db.


    frmHiddenSounds.snd.SoundType = 2 ' DirectSound
    frmHiddenSounds.snd.DSCommand = 1 ' Initialize DirectSound
    
    If (frmHiddenSounds.snd.FXError > 0) Then
        MsgBox ("Failed to initialize DirectSound Driver" + Chr$(13) + "and requires the MS-Game SDK for Windows")
    End If
    
    ' Initialise the sound levels to maximum for now
    '11/12/00 use saved volumes
    'For i = 1 To 31
    '    sndVolume(0, i) = 0 'for PC 0
    '    sndVolume(1, i) = 0 'for PC 1
    'Next i
    
    Call LoadDefaultSoundMixData
    
    'SetVolumes ' Setup the volume levels for playback
    SndPath = "c:\steam_sim\Sounds\"
    
    'Force for now
    'DummyPCNumber = 1
    
'    If This_PC_number = 0 Then 'use host PC for Main ER sounds on Left Ch & Elec MCCB on Right
    '04/04/07 allow sounds on all PCs
    'If This_PC_number = 0 Then 'use host PC for Main ER sounds on Left Ch & Elec MCCB on Right
        
        frmHiddenSounds.snd.DSChannel = 1 ' Channel 1 Diesel Alternator
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Diesel Alternator.wav"
        frmHiddenSounds.snd.Filename = SoundPathname 'Diesel Alternator
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 1)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 1)

        frmHiddenSounds.snd.DSChannel = 2 ' Channel 2 SSTG #1
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Turbo Alternator.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 2)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 2)

        frmHiddenSounds.snd.DSChannel = 3 ' Channel 3 SSTG #2
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Turbo Alternator.wav"
        '10/12/00 add SSTG2 frmHiddenSounds.snd.Filename = SndPath + "pump2.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 3)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 3)

        frmHiddenSounds.snd.DSChannel = 4 ' Channel 4 Pump
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Pump.wav"
        '10/12/00 frmHiddenSounds.snd.Filename = SndPath + "Full Steam Vent.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 4)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 4)
    
    
        frmHiddenSounds.snd.DSChannel = 5 ' Channel 5 Steam Vent
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Drum Safety Valve Blowing.wav"
        '10/12/00  frmHiddenSounds.snd.Filename = SndPath + "FD.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 5)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 5)
    
        frmHiddenSounds.snd.DSChannel = 6 ' Channel 6 FD Fan 1
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "FD Fan.wav"
        '10/12/00  frmHiddenSounds.snd.Filename = SndPath + "prop2.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 6)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 6)
    
        frmHiddenSounds.snd.DSChannel = 7 ' Channel 7 FD Fan 2
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "FD Fan.wav"
        '10/12/00  frmHiddenSounds.snd.Filename = SndPath + "prop2.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 7)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 7)
       
        frmHiddenSounds.snd.DSChannel = 8 ' Channel 8 Prop
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Propshaft.wav"
        '10/12/00  frmHiddenSounds.snd.Filename = SndPath + "compoff2.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000  'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 8)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 8)

        frmHiddenSounds.snd.DSChannel = 9 ' Channel 8 Compressor
        frmHiddenSounds.snd.DSRepeat = False
        SoundPathname = SndPath + "Compressor.wav"
        '10/12/00   frmHiddenSounds.snd.Filename = SndPath + "brkoff4.wav" ' Breaker Sound to Right Channel
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 9)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 9)
    
    '10/12/00
        frmHiddenSounds.snd.DSChannel = 10 ' Channel 10 Breaker (from Right Channel)
        frmHiddenSounds.snd.DSRepeat = False
        SoundPathname = SndPath + "Big Breaker.wav"
        frmHiddenSounds.snd.Filename = SoundPathname ' Main Breaker Sound to Right Channel
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = 10000 'Right Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 10)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 10)

        frmHiddenSounds.snd.DSChannel = 11 ' Channel 11 Small Breaker (from Right Channel)
        frmHiddenSounds.snd.DSRepeat = False
        SoundPathname = SndPath + "Small Breaker.wav"
        frmHiddenSounds.snd.Filename = SoundPathname ' Breaker Sound to Right Channel
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = 10000 'Right Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 11)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 11)

   
        frmHiddenSounds.snd.DSChannel = 12 ' Channel 12 Superheater Vent
        frmHiddenSounds.snd.DSRepeat = True 'think the vent sound stays on till vent closes
        SoundPathname = SndPath + "Superheater Vent.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 12)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 12)

        frmHiddenSounds.snd.DSChannel = 13 ' Channel 13 Superheater Safety Valve
        frmHiddenSounds.snd.DSRepeat = False
        SoundPathname = SndPath + "Superheater Safety Valve Blowing.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000 'Left Channel Only
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 13)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 13)
        
        
        

        frmHiddenSounds.snd.DSChannel = 14 ' Channel 14 Scott's Boom
        frmHiddenSounds.snd.DSRepeat = False
        SoundPathname = SndPath + "mine1.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
               
       ' === 26/11/01   14.051 Scott wants the Boom to come from the ER ====
      ' frmHiddenSounds.snd.DSPan = 0       'Both Channels
        frmHiddenSounds.snd.DSPan = -10000  'Left Channel Only (Engine Room)
              
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 13)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 13)
        
        
        
    
        frmHiddenSounds.snd.DSChannel = 15 ' Channel 15 main Turbine
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Main Turbine.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000  'Left Channel Only
        ' for test frmHiddenSounds.snd.DSPan = 0  '
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 15)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 15)
    
    
       ' === 12/11/02 8.00 ======================
        frmHiddenSounds.snd.DSChannel = 16 ' Channel 16    EOT Bell
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "BELL_EOT.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000  'Left Channel Only
        ' for test frmHiddenSounds.snd.DSPan = 0  '
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 15)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 15)
    
       ' === 12/11/02 8.00 ======================
        frmHiddenSounds.snd.DSChannel = 17 ' Channel 17    Transfer Request Bell
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "BELL_Transfer_Request.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000  'Left Channel Only
        ' for test frmHiddenSounds.snd.DSPan = 0  '
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 15)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 15)
        
       ' === 12/11/02 8.00 ======================
        frmHiddenSounds.snd.DSChannel = 18 ' Channel 18    Lub Oil Bell
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "BELL_Lub_Oil.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer'=====================
        frmHiddenSounds.snd.DSPan = -10000  'Left Channel Only
        ' for test frmHiddenSounds.snd.DSPan = 0  '
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 15)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 15)
        
        '=== 04/04/07 move Turbo Feed Pump sounds to single PC for PC version
        
        frmHiddenSounds.snd.DSChannel = 19 ' Channel 19
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Turbo Feed Pump.wav"
        frmHiddenSounds.snd.Filename = SoundPathname
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = -10000
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 19)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 19)
    
        frmHiddenSounds.snd.DSChannel = 20 ' Channel 20
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "Turbo Feed Pump.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname '
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = 10000
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 20)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 20)
    
        '15/06/07 add Bell for Gen Alarm
        frmHiddenSounds.snd.DSChannel = 21 ' Channel 21
        frmHiddenSounds.snd.DSRepeat = True
        SoundPathname = SndPath + "BELL_GenAlarm.wav" '
        frmHiddenSounds.snd.Filename = SoundPathname '
        frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
        frmHiddenSounds.snd.DSPan = 10000
        frmHiddenSounds.snd.DSVolume = -sndVolume(0, 21)
        frmHiddenSounds.snd.DSFrequency = sndFreq(0, 21)
    
    'End If
    
    'ElseIf This_PC_number = 1 Then 'use second Instr PC for Turbo Feed Pumps &1 and #2
    
    '04/04/07 move turbo sounds to one PC
    'ElseIf This_PC_number = 1 Then 'use second Instr PC for Turbo Feed Pumps &1 and #2
    
   '     frmHiddenSounds.snd.DSChannel = 1 ' Channel 1
   '     frmHiddenSounds.snd.DSRepeat = True
   '     SoundPathname = SndPath + "Turbo Feed Pump.wav"
   '     frmHiddenSounds.snd.Filename = SoundPathname
   '     frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
   '     frmHiddenSounds.snd.DSPan = -10000
   '     frmHiddenSounds.snd.DSVolume = -sndVolume(1, 1)
   '     frmHiddenSounds.snd.DSFrequency = sndFreq(1, 1)
        
   '     frmHiddenSounds.snd.DSChannel = 2 ' Channel 2
   '     frmHiddenSounds.snd.DSRepeat = True
   '     SoundPathname = SndPath + "Turbo Feed Pump.wav" '
   '     frmHiddenSounds.snd.Filename = SoundPathname '
   '     frmHiddenSounds.snd.DSCommand = 2 ' Load wave data into buffer
   '     frmHiddenSounds.snd.DSPan = 10000
   '     frmHiddenSounds.snd.DSVolume = -sndVolume(1, 2)
   '     frmHiddenSounds.snd.DSFrequency = sndFreq(1, 2)
   ' End If
    
     Debug.Print "Initialsed Sounds"
Exit Sub


Errh1:

resp = MsgBox("There was an error reading " & SoundPathname & Chr(10) & _
       "this sound will not be available", vbCritical + vbOKOnly, "Initialization Error")

        SoundPathname = SndPath + "silence.wav"
        Resume Next
        
End Sub






Public Sub SoundWhenModelRunning()
Dim i As Integer
Dim eng_set_point As Single, eng_accn As Single, eng_speed As Single
Dim turb_set_point As Single, turb_accn As Single, turb_speed As Single
Dim prop_set_point As Single, prop_accn As Single, prop_speed As Single
Dim Sound_Speed As Double
Dim pump1ON, pump2ON As Boolean
Dim freq1, freq2 As Integer

Dim single_char As String * 1

'snd.AutoPlay
      
      
'These are the Sounds which should be present
'Safety valves 1/2        Loud Hiss continues until safety shuts
'Forced Draft Fan #1/#2   'FD Fan
'Superheater vents 1/2
'Steam to Turbo-Generator
'Steam admission to #1/#2 Feed Pump
'Turbo-Generator volume and noise
'#1 Feed Pump turbine
'#2 Feed Pump turbine
'Turbo-Generator vibration
'#1 Feed Pump load
'#2 Feed Pump load
'Air compressor noise (cyclical)
'Main turbine volume and speed
'Diesel Generator noise and load
'Main engine rumble and white noise
'Superheater safety valves
'Turbine and main gearing noise
'Water impingement on Feed Pumps

'Other sounds to be supported include:
'cavitation of pumps,
'poor operation of feed pumps due to water impingement on turbine blades
'generic grinding, mechanical noise.
      
If Sounds_Enabled Then

'Diesel Alternator
    '04/04/07 now o/p sound on all PCs
    'If This_PC_number = 0 Then 'o/p main mix of sounds on left and CCB on right
    
        If LSS_reals(45) > 10 Then 'assume SDG is running
            SndEnabled(0, 1) = True
            'Sound_Speed = LSS_reals(45) 'max speed 760 rpm sounds ok with freq set to 16000, so
            '10/12/00 dont understand why this was commented out before
            Sound_Speed = LSS_reals(45) 'max speed 760 rpm sounds ok with freq set to 16000, so
            sndFreq(0, 1) = Sound_Speed * 21 '16000 / 760 = 21
        Else
            SndEnabled(0, 1) = False
        End If


         If LSS_reals(47) > 100 Then 'assume #1 SSTG is running
            SndEnabled(0, 2) = True
            Sound_Speed = LSS_reals(47)
            sndFreq(0, 2) = Sound_Speed * 9 / 8 'max speed 8000 rpm sounds ok with freq set to 9000, so
        Else
            SndEnabled(0, 2) = False
        End If
        
'10/12/00 added SSTG #2, so move all sounds down 1 channel from before

         If LSS_reals(46) > 100 Then 'assume #2 SSTG is running
            SndEnabled(0, 3) = True
            Sound_Speed = LSS_reals(46)
            sndFreq(0, 3) = Sound_Speed * 9 / 8 'max speed 8000 rpm sounds ok with freq set to 9000, so
        Else
            SndEnabled(0, 3) = False
        End If


        If VB_steam_logicals(100) Then 'Pump sound
            '10/12/00 SndEnabled(0, 3) = True
            SndEnabled(0, 4) = True
            sndFreq(0, 4) = 24000 'This sounds a lot faster than our orig pump but is preferred by SG 11/12/00
        Else
            '10/12/00 SndEnabled(0, 3) = False
            SndEnabled(0, 4) = False
        End If

        
        'Safeties - 3 on each boiler
        If VB_steam_logicals(80) Or VB_steam_logicals(89) Or _
           VB_steam_logicals(81) Or VB_steam_logicals(90) Then 'a drum safety has blown
            '10/12/00  SndEnabled(0, 4) = True
            SndEnabled(0, 5) = True
        Else
            '10/12/00  SndEnabled(0, 4) = True
            SndEnabled(0, 5) = False
        End If
        
        
        'FD Fan STBD Sound
        
        single_char = Mid$(PC_INT_STRING, 255, 1)
        If single_char = "0" Then
            pp_status_LSS(5) = 0 'Pump Off
        ElseIf single_char = "1" Then
            pp_status_LSS(5) = 1 'Pump Off and  in AUTO
        ElseIf single_char = "2" Then
            pp_status_LSS(5) = 2 'Pump On and in AUTO
        Else
            pp_status_LSS(5) = 3 'Pump On and not in AUTO
        End If

        
        '10/12/00 If pp_status_LSS(5) = 3 Then      '  STBD FD Fan  is ON
        If pp_status_LSS(5) = 3 Then      '  STBD FD Fan is ON and running SLOW
            pump1ON = True
            freq1 = 3000
            SndEnabled(0, 6) = True
            'NOTE According to RWJ, VB_steam_logicals(12) is an alternative flag for fast(0) and slow(1)
            'If (VB_steam_logicals(12)) Then '  STBD FD Fan  is ON ( SLOW )
            'Else                            '  STBD FD Fan  is ON ( FAST )
            'End If
        ElseIf pp_status_LSS(5) = 2 Then      '  STBD FD Fan is ON and running FAST
            pump1ON = True
            freq1 = 4500
            SndEnabled(0, 6) = True
        Else        '  STBD FD Fan  is OFF
            pump1ON = False
            SndEnabled(0, 6) = False
        End If

        
        
        
        'FD Fan PORT Sound
        single_char = Mid$(PC_INT_STRING, 267, 1)
        If single_char = "0" Then
            pp_status_LSS(17) = 0
        ElseIf single_char = "1" Then
            pp_status_LSS(17) = 1
        ElseIf single_char = "2" Then
            pp_status_LSS(17) = 2
        Else
            pp_status_LSS(17) = 3
        End If
    
      If pp_status_LSS(17) = 3 Then      '  PORT FD Fan  is ON AND slow
        pump2ON = True
        freq1 = 3000
        SndEnabled(0, 7) = True

        'If (VB_steam_logicals(13)) Then  '  PORT FD Fan  is ON ( SLOW )
        'Else                             '  PORT FD Fan  is ON ( FAST )
        'End If
        ElseIf pp_status_LSS(17) = 2 Then      '  port FD Fan is ON and running FAST
            pump1ON = True
            freq1 = 4500
            SndEnabled(0, 7) = True
        Else        '  port FD Fan  is OFF
            pump1ON = False
            SndEnabled(0, 7) = False
        End If
      
      
       '10/12/00 Add Prop sound
        Sound_Speed = Abs(LSS_reals(80)) 'assume range is 0 to about 100
        If Sound_Speed > 1 Then 'assume Prop turning
            SndEnabled(0, 8) = True
            sndFreq(0, 8) = 2000 + (Sound_Speed * 50) 'good range is 2000 (slow) to 7000 (fast)
        Else
            SndEnabled(0, 8) = False
        End If

        'NOTE could make this more sophisticated by introducing an on but offloading sound
        'I_SOUNDS = 0 (Off), 1 ON but not loaded, 2 = ON and loaded
        '10/12/00 Add Compressor sound
        If I_sounds(5) > 0 Or I_sounds(6) > 0 Then 'a compressor is on
            SndEnabled(0, 9) = True
        Else
            SndEnabled(0, 9) = False
        End If
     
      
        If VB_steam_logicals(188) Then 'Main Feeder Breaker sound
            '10/12/00 SndEnabled(0, 8) = True
            SndEnabled(0, 10) = True
            status_flag = False
            Call set_STEAM_logicals(188, status_flag)
        Else
            '10/12/00 SndEnabled(0, 8) = True
            SndEnabled(0, 10) = False
        End If
   
        'Add Small Breaker sound here
        If VB_steam_logicals(99) Then 'Small Breaker sound
            SndEnabled(0, 11) = True
            status_flag = False
            Call set_STEAM_logicals(99, status_flag)
        Else
            SndEnabled(0, 11) = False
        End If
      
        'Add Superheater Vent sound here, waiting for Flag from Richard
        
        ' === 14/12/00 ====
        ' STBD Superheater Vent
         vv_status_LSS(444) = VB_VV_extras(444 - 250)
         ' PORT Superheater Vent
         vv_status_LSS(474) = VB_VV_extras(474 - 250)
               
        ' === 06/03/01 Point 10.103 ===
        '      PC_REALS(127) = Psteam(5)
        '      PC_REALS(128) = Psteam(105)
     
       ' === 06/03/01 Point 10.103 ===
       'If vv_status_LSS(474) = 3 Or vv_status_LSS(474) = 3 Then 'Superheater Vent sound
        If ((vv_status_LSS(474) = 3 And LSS_reals(128) > 2#) _
         Or (vv_status_LSS(444) = 3 And LSS_reals(127) > 2#)) Then       'Superheater Vent sound
              
            SndEnabled(0, 12) = True
        Else
            SndEnabled(0, 12) = False
        End If

        'Add Superheater Safety Valve sound here,
        If VB_steam_logicals(91) Or VB_steam_logicals(92) Then 'Superheater Safety Valve
            SndEnabled(0, 13) = True
            status_flag = False
            Call set_STEAM_logicals(91, status_flag)
            Call set_STEAM_logicals(92, status_flag)
        Else
            SndEnabled(0, 13) = False
        End If

        '14/12/00 If Fault(1255) Then 'Explosion
        If Explosion_count < 25 Then 'Explosion
            SndEnabled(0, 14) = True
            '14/12/00 Fault(1255) = False
            
            Explosion_count = Explosion_count + 1
        Else
            SndEnabled(0, 14) = False
        End If

      
        Sound_Speed = Abs(LSS_reals(80))
        If Sound_Speed > 5# Then  'assume Main Turbine turning
            SndEnabled(0, 15) = True
            sndFreq(0, 15) = 2000 + (Sound_Speed * 280) 'freq = 2000 for slowest speed, 30,000 for full
        Else
            SndEnabled(0, 15) = False
        End If
        
        '=== 04/04/07 added turbo fp sounds here (not on second PC)
        If LSS_reals(398) > 0.02 Then
               SndEnabled(0, 19) = True
               sndFreq(0, 19) = LSS_reals(398) * 9000 'Turbo FeedPump #1
        Else
               SndEnabled(0, 19) = False
        End If
            
            
       
        If LSS_reals(399) > 0.02 Then
                SndEnabled(0, 20) = True
                sndFreq(0, 20) = LSS_reals(399) * 9000 'Turbo FeedPump #2
        Else
                SndEnabled(0, 20) = False
        End If
        
        
        
    ' === 13/11/02 TEXAS point 8.00 BELL SOUNDS PTT ===
    If (PTT) Then
        If VB_More_integers(5) = 1 Then ' EOT_MECD_BELL_steady equivalence  (DO(767),EOT_MECD_BELL_steady)
            SndEnabled(0, 16) = True
        Else
            SndEnabled(0, 16) = False
        End If
        If VB_More_integers(6) = 1 Then ' TRANSFER REQUEST SIREN ( Whooper )    /1063/
            SndEnabled(0, 17) = True
        Else
            SndEnabled(0, 17) = False
        End If
        If VB_More_integers(7) = 1 Then ' LUBE_OIL_Bell_DO      /1067/
            SndEnabled(0, 18) = True
        Else
            SndEnabled(0, 18) = False
        End If
        '15/06/07 add Gen Alarm Bell
        If Gen_Alarm_Bell_flag = True Then
            SndEnabled(0, 21) = True
        Else
            SndEnabled(0, 21) = False
        End If
        
        
        
    Else
            SndEnabled(0, 16) = False
            SndEnabled(0, 17) = False
            SndEnabled(0, 18) = False
            '15/07/08
            SndEnabled(0, 21) = False
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ' =====================================================
    
    

      
    '10/12/00 now use 2 separate channels for FD fans, no need to mix
    'If pump1ON Or pump2ON Then 'FD fans
    '    If (pump1ON) Then
    '        If (VB_steam_logicals(12)) Then
    '            freq1 = 3000
    '        Else
    '            freq1 = 4500
    '        End If
    '    End If
    '    If (pump2ON) Then
    '        If (VB_steam_logicals(13)) Then
    '            freq2 = 3000
    '        Else
    '            freq2 = 4500
    '        End If
    '    End If
    '    If freq2 > freq1 Then
    '      sndFreq(0, 5) = freq2
    '    Else
    '      sndFreq(0, 5) = freq1
    '    End If
    '    SndEnabled(0, 5) = True
    'Else
    '    SndEnabled(0, 5) = False
    'End If
        
  '  ElseIf This_PC_number = 1 Then  'o/p to turbo feedpumps.
                                    'At max speed c8000rpm LSS_REALS is 0.95),and this
                                     ' sounds ok whith freq = 9000
                                   
    
        '   If LSS_reals(398) > 0.02 Then
        '       SndEnabled(1, 1) = True
        '        sndFreq(1, 1) = LSS_reals(398) * 9000 'Turbo FeedPump #1
        '    Else
        '       SndEnabled(1, 1) = False
        '    End If
            
            
       
        '     If LSS_reals(399) > 0.02 Then
        '        SndEnabled(1, 2) = True
        '        sndFreq(1, 2) = LSS_reals(399) * 9000 'Turbo FeedPump #2
        '    Else
        '        SndEnabled(1, 2) = False
        '    End If
    

'    End If
Else

' ====04/04/07
        For i = 1 To 18
          SndEnabled(0, i) = False
        Next
        '15/06/07
        SndEnabled(0, 21) = False
        
        
        ' ====04/04/07
        'For i = 1 To 2
        '  SndEnabled(1, i) = False
        'Next
End If
    
    
    
    
'==== 04/04/07 ====
'If This_PC_number = 0 Then 'do for PC 0
    
'            === 13/11/02 ===
'            For i = 1 To 16

'             ==== 04/04/07 ====
'             For i = 1 To 18
              '15/06/07
              'For i = 1 To 20
              For i = 1 To 21
                frmHiddenSounds.snd.DSChannel = i        '
                frmHiddenSounds.snd.DSFrequency = sndFreq(0, i)
                '13/12/00 frmHiddenSounds.snd.DSVolume = 0
                frmHiddenSounds.snd.DSVolume = sndVolume(0, i)
                If SndEnabled(0, i) Then
                    '15/06/07
                    'Debug.Print "frmHiddenSounds.snd.DSStatus  = " & frmHiddenSounds.snd.DSStatus & "  " & Time
                    If frmHiddenSounds.snd.DSStatus <> 1 Then
                        frmHiddenSounds.snd.DSCommand = 3       ' Play the sound
                    End If
                Else
                    frmHiddenSounds.snd.DSCommand = 4           ' Stop the sound
                End If
            Next
'             ==== 04/04/07 ==== No sound required from PC01
            
'ElseIf This_PC_number = 1 Then
    
'           For i = 1 To 2
'                frmHiddenSounds.snd.DSChannel = i        '
 '               frmHiddenSounds.snd.DSFrequency = sndFreq(1, i)

 '               frmHiddenSounds.snd.DSVolume = sndVolume(1, i)
 '               If SndEnabled(1, i) Then
 '                   If frmHiddenSounds.snd.DSStatus <> 1 Then
 '                       frmHiddenSounds.snd.DSCommand = 3       ' Play the sound
 '                   End If
'                Else
 '                   frmHiddenSounds.snd.DSCommand = 4           ' Stop the sound
  '              End If
 '           Next
        
        
'End If






'End If

      
GoTo TheEnd
  
      
' === 13/11/02 Because of GoTo TheEnd (above) do NOT NEED....  ====
'If Sounds_Enabled Then
'     ' AIR COMPRESSORS
'            ' I_sounds(5) = 1 = compressor 1 is RUNNING  ( = 2 when loaded )
'            ' I_sounds(6) = 1 = compressor 2 is RUNNING  ( = 2 when loaded )
'        If (I_sounds(5) > 0 Or I_sounds(6) > 0) Then
'             frmHiddenSounds.snd.DSChannel = 6         ' Compressor RUNNING sound ON
'            If (frmHiddenSounds.snd.DSStatus = 0) Then frmHiddenSounds.snd.DSCommand = 3   ' If not playing, then play it
'        Else
'                 frmHiddenSounds.snd.DSChannel = 6     ' Compressor RUNNING sound OFF
'           If (frmHiddenSounds.snd.DSStatus = 1) Then frmHiddenSounds.snd.DSCommand = 4   ' Stop the sound
'        End If
'        If (I_sounds(5) = 1 And old_I_sounds(5) = 2) Or (I_sounds(6) = 1 And old_I_sounds(6) = 2) Then
'            Debug.Print "Activating OFF-Loading Sound"
'              frmHiddenSounds.snd.DSChannel = 7       ' Compressor OFF-LOADING sound ON
'             If (frmHiddenSounds.snd.DSStatus = 0) Then frmHiddenSounds.snd.DSCommand = 3   ' If not playing, then play it
'        End If
'                 old_I_sounds(5) = I_sounds(5)
'                 old_I_sounds(6) = I_sounds(6)
'
'
'
'      ' PUMP SOUND
'            frmHiddenSounds.snd.DSChannel = 5        ' Pump sound
'        If I_sounds(4) > 0 Then        '  I_sounds(4) = Number of pumps RUNNING
'           If (frmHiddenSounds.snd.DSStatus = 0) Then frmHiddenSounds.snd.DSCommand = 3   ' If not playing, then play it
'        Else
'           If (frmHiddenSounds.snd.DSStatus = 1) Then frmHiddenSounds.snd.DSCommand = 4   ' Stop the sound
'        End If
'
'       ' MAIN ENGINE SOUND
'       'eng_set_point = (Engspeed / 255) * 100 ' Scale it to be the same as the slider
'        eng_set_point = (I_sounds(1) / 255) * 100 ' Scale it to be the same as the slider
'        eng_accn = (eng_set_point - eng_speed) * 0.5
'        eng_speed = eng_speed + eng_accn * 0.1 '(Timer1.Interval * 0.001)
'
'        frmHiddenSounds.snd.DSChannel = 1        ' Main Engine sound
'        frmHiddenSounds.snd.DSFrequency = 4000 + Int(eng_speed * 100)
'        If (frmHiddenSounds.snd.DSStatus = 0) And (I_sounds(1) > 0) Then frmHiddenSounds.snd.DSCommand = 3 ' If not playing, then play it
'        If (frmHiddenSounds.snd.DSStatus = 1) And (I_sounds(1) = 0) Then frmHiddenSounds.snd.DSCommand = 4 ' Else if too low, then stop the sound
'
'
'        '  18/8/98
'        '   Debug.Print " I_sounds(1) = " & I_sounds(1)
'
'       ' TURBO-CHARGER SOUND
'       'turb_set_point = (Turbspeed / 255) * 100
'        turb_set_point = (I_sounds(2) / 255) * 100
'        turb_accn = (turb_set_point - turb_speed) * 0.5
'        turb_speed = turb_speed + turb_accn * 0.1 '(Timer1.Interval * 0.001)
'        frmHiddenSounds.snd.DSChannel = 2        ' Turbo sound
'        frmHiddenSounds.snd.DSFrequency = 4000 + Int(turb_speed * 100)
'
'    ' 8/4/98
'    '  If (frmHiddenSounds.snd.DSStatus = 0) And (I_sounds(2) > 0) Then frmHiddenSounds.snd.DSCommand = 3 ' If not playing and the engine is on, then play
'    '  If (frmHiddenSounds.snd.DSStatus = 1) And (I_sounds(2) = 0) Then frmHiddenSounds.snd.DSCommand = 4 'Else, don't play
'        If I_sounds(2) > 2 Then
'          If (frmHiddenSounds.snd.DSStatus = 0) Then frmHiddenSounds.snd.DSCommand = 3  ' If not playing and the engine is on, then play
'        Else
'          If (frmHiddenSounds.snd.DSStatus = 1) Then frmHiddenSounds.snd.DSCommand = 4 'Else, don't play
'        End If
'
'       ' PROP-SHAFT SOUND
'        frmHiddenSounds.snd.DSChannel = 3     ' Prop shaft sound
'       'prop_set_point = (Propspeed / 255) * 100
'        prop_set_point = (I_sounds(7) / 255) * 100
'        prop_accn = (prop_set_point - prop_speed) * 0.5
'        prop_speed = prop_speed + prop_accn * 0.1   '(Timer1.Interval * 0.001)
'        frmHiddenSounds.snd.DSFrequency = 4000 + Int(prop_speed * 100)
'
''       If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3
''        If (frmHiddenSounds.snd.DSStatus = 0) And (Propspeed > 1) Then frmHiddenSounds.snd.DSCommand = 3 ' If not playing and the engine is on, then play
''        If (frmHiddenSounds.snd.DSStatus = 1) And (Propspeed < 2) Then frmHiddenSounds.snd.DSCommand = 4 'Else, don't play
' 29/10/97 Got clicks because Propspeed is a REAL VARIABLE so...
'       If Propspeed > 2# Then
'        If I_sounds(7) > 2# Then
'          If (frmHiddenSounds.snd.DSStatus = 0) Then frmHiddenSounds.snd.DSCommand = 3  ' If not playing and the engine is on, then play
'        Else
'          If (frmHiddenSounds.snd.DSStatus = 1) Then frmHiddenSounds.snd.DSCommand = 4 'Else, don't play
'        End If
'
'
'        ' ENGINE ROOM FAN SOUND
'        If I_sounds(3) > 0 Then
'          frmHiddenSounds.snd.DSChannel = 4     ' Fan sound
'          If frmHiddenSounds.snd.DSStatus <> 1 Then frmHiddenSounds.snd.DSCommand = 3
'        Else
'          frmHiddenSounds.snd.DSChannel = 4     ' Fan sound
'          frmHiddenSounds.snd.DSCommand = 4     ' STOP the FANSOUND
'        End If
'
'Else
'        For i = 1 To 6
'             frmHiddenSounds.snd.DSChannel = i
'             frmHiddenSounds.snd.DSCommand = 4   ' SWITCH SOUND OFF
'         Next i
'End If


TheEnd:

End Sub

Private Sub Form_Load()
'29/11/99 initialize all sounds on
Dim i As Integer

Explosion_count = 25 'prevent explosion on start up


For i = 1 To 31
    'SndEnabled(0, i) = True 'for sounds from pc00
    'SndEnabled(1, i) = True 'for sounds from pc01
    '30/11/99 let the model set the sounds
    SndEnabled(0, i) = False 'for sounds from pc00
    SndEnabled(1, i) = False 'for sounds from pc01
Next

'put MCCB sound off for now
'SndEnabled(0, 8) = False
End Sub

