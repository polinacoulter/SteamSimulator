Attribute VB_Name = "IO_General"


' === 12/12/01 ===
Global AI_update_required As Boolean



'18/2/98
Global GPIOTestOn As Byte


'10/2/98
Global DIBaseDefault As Integer
Global DITopDefault As Integer
Global DIBaseSelected As Integer
Global DITopSelected As Integer
Global LastDIBase As Integer
Global LastDITop As Integer
Global ResetUpDowns As Byte



' 6/2/98
Global DO_number As Integer
Global AO_Number As Integer

'20/10/97
Global HalfScaleReading(256) As Integer

'28/9/97
Global AOSliderValue As Integer



' 20/2/98
Global VB_bus_test As Byte

'5/8/97
Global ResetAOChannels As Byte
'6/8/97

'4/11/97 logging ai
Global AILogDummy(0 To 255) As Integer


' 20/4/98 Global AI(0 To 255) As Integer
' 20/4/98 Global AO(0 To 767) As Integer
' 20/4/98 Global AOold(0 To 767) As Byte ' need an AOold for sound o/p (0 to 255)
' 20/4/98 Global DI(0 To 3071) As Byte
' 20/4/98 Global DIold(0 To 3071) As Byte
' 20/4/98 Global DIchanges(0 To 127) As Byte ' used for logging

Global DI_Changes(0 To 127) As Integer ' used for logging
Global DI_States(0 To 127) As Byte ' used for logging
Global DiCount As Integer
Global AiCount As Integer
Global AiNew(0 To 255) As Integer
Global AiOld(0 To 255) As Integer
Global AIChanging(0 To 255) As Boolean
Global AIChanged(0 To 255) As Boolean
Global AI_Changes(0 To 127) As Integer


'Global DITEXT(0 To 1700) As String * 48
Global DITEXT(0 To 3071) As String * 48




' 20/4/98
'Global VB_DO(0 To 5119) As Byte 'can't use "DO" as an array name

Global DO_Temp(0 To 63) As Byte
Global LampTestOn As Byte
Global DOTestON As Byte
Global BuzzTestOn As Byte
Global AOTestOn As Byte
Global LampTestCounter As Byte
Global DelayLoopCount As Double 'used to determine fixed delay during i/o (dependant on processor speed)

Global IOCounter As Integer

' 20/4/98 Global HWTestOn As Byte

' 20/4/98  Global test_flag As Byte

Global ICODE As Integer
'Global Instructor_name As String * 20
'Global Class_name As String * 20
'Global Course_name As String * 20





Global AI_chan_no As Integer, AI_Value As Integer






Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)






Sub ReportError(ByVal Code As Integer)


    Dim message As String
    Dim DgDef%, response%

    Select Case (Code)
    Case OK
        message = "Error code 0: OK!"
    Case ERRSUPPORT
        message = "Error code -1: Board/feature not supported"
    Case ERRBASE
        message = "Error Code -2: Address already registered"
    Case ERRIRQ
        message = "Error Code -3: IRQ level already registered"
    Case ERRHANDLE
        message = "Error Code -4: Board/resource not registered"
    Case ERRCHAN
        message = "Error Code -5: Invalid channel"
    Case ERRDATA
        message = "Error Code -6: Invalid data"
    Case ERRRANGE
        message = "Error Code -7: Out of range"
    Case ERRMEMORY
        message = "Error Code -8: Insufficient memory"
    Case ERRBUFFER
        message = "Error Code -9: Buffer not registered"
    Case ERRPC226
        message = "Error Code = -10: PC226E not found"
    Case Else
        message = "Error Code " + str$(Code) + ": Undefined code."
    End Select
    

    DgDef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2                     'describe dialogue
    response = MsgBox(message, DgDef, "PC214 Board DLL Error")           'get user response

End Sub



