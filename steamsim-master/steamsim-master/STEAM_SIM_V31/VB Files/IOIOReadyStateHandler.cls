VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOIOReadyStateHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
    ' Snapshot the global reference. Request_IOIO_Data unconditionally nils ioioXmlHttp
    ' on each poll, and a callback for a previous request can fire after that nil — so
    ' we work off our own ref and let the handle drop when this Sub returns.
    On Error Resume Next

    Dim http As MSXML2.XMLHTTP
    Set http = ioioXmlHttp

    If http Is Nothing Then Exit Sub
    If http.readyState <> 4 Then Exit Sub

    Dim responseText As String
    responseText = http.responseText

    Call TryLoadAnalogInputsFromCsv(responseText)
    Call TryLoadDigitalInputsFromCsv(responseText)
End Sub

Private Function TryLoadAnalogInputsFromCsv(responseText As String) As Boolean
    ' Server emits one or more lines; we want the line that starts with "ain=".
    ' Example payload:
    '   poll_ms=500
    '   ain=0,12,55,...
    '   aout=...
    Dim ainLine As String
    ainLine = ExtractValueAfterPrefix(responseText, "ain=")
    If Len(ainLine) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(ainLine, ",")

    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        If ApplyArrayAnalogInput(i, parts(i)) Then
            TryLoadAnalogInputsFromCsv = True
        End If
    Next i
End Function

Private Function ExtractValueAfterPrefix(responseText As String, prefix As String) As String
    Dim startPos As Integer
    startPos = InStr(1, responseText, prefix, vbTextCompare)
    If startPos = 0 Then Exit Function

    Dim valStart As Integer
    valStart = startPos + Len(prefix)

    Dim crPos As Integer, lfPos As Integer
    crPos = InStr(valStart, responseText, vbCr)
    lfPos = InStr(valStart, responseText, vbLf)

    If crPos = 0 Then crPos = Len(responseText) + 1
    If lfPos = 0 Then lfPos = Len(responseText) + 1

    Dim endPos As Integer
    If crPos < lfPos Then
        endPos = crPos
    Else
        endPos = lfPos
    End If

    ExtractValueAfterPrefix = Mid$(responseText, valStart, endPos - valStart)
End Function

Private Function ApplyArrayAnalogInput(sourceIndex As Integer, rawValue As String) As Boolean
    Dim normalizedValue As String
    normalizedValue = Trim$(rawValue)
    If Len(normalizedValue) = 0 Then Exit Function

    Dim temp As Integer
    temp = CInt(Val(normalizedValue))
    If temp < 0 Or temp > 255 Then Exit Function

    ApplyArrayAnalogInput = StoreAnalogInputValue(sourceIndex, IOIO_INPUT_SOURCE_BASE, CByte(temp))
End Function

Private Function StoreAnalogInputValue(sourceIndex As Integer, sourceBase As Integer, Value As Byte) As Boolean
    Dim targetIndex As Integer
    targetIndex = sourceIndex - sourceBase + IOIO_INPUT_TARGET_BASE

    If targetIndex < 0 Then Exit Function
    If targetIndex > UBound(A_INPUT) Then Exit Function

    A_INPUT(targetIndex) = Value
    StoreAnalogInputValue = True
End Function

Private Function TryLoadDigitalInputsFromCsv(responseText As String) As Boolean
    ' Same shape as analog: line "din=v0,v1,v2,..." with 0/1 per channel.
    Dim dinLine As String
    dinLine = ExtractValueAfterPrefix(responseText, "din=")
    If Len(dinLine) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(dinLine, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If ApplyArrayDigitalInput(i, parts(i)) Then
            TryLoadDigitalInputsFromCsv = True
        End If
    Next i
End Function

Private Function ApplyArrayDigitalInput(sourceIndex As Long, rawValue As String) As Boolean
    Dim normalizedValue As String
    normalizedValue = Trim$(rawValue)
    If Len(normalizedValue) = 0 Then Exit Function

    ' Anything non-zero is "on". D_INPUT is a Byte array (see IOTypes.bas), so we
    ' store 0 or 1 — `If D_INPUT(idx) Then` evaluates either correctly.
    Dim asserted As Byte
    If CInt(Val(normalizedValue)) <> 0 Then
        asserted = 1
    Else
        asserted = 0
    End If

    ApplyArrayDigitalInput = StoreDigitalInputValue(sourceIndex, IOIO_DIGITAL_SOURCE_BASE, asserted)
End Function

Private Function StoreDigitalInputValue(sourceIndex As Long, sourceBase As Long, Value As Byte) As Boolean
    Dim targetIndex As Long
    targetIndex = sourceIndex - sourceBase + IOIO_DIGITAL_TARGET_BASE

    If targetIndex < 0 Then Exit Function
    If targetIndex > UBound(D_INPUT) Then Exit Function

    D_INPUT(targetIndex) = Value
    StoreDigitalInputValue = True
End Function
