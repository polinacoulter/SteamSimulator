VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOAppHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Async response handler for Get_Outputs_From_Python. Parses aout= and dout=
' lines from the Python server's /ioio/status response and writes them into
' A_OUTPUT() and D_OUTPUT(). Mirrors IOIOReadyStateHandler.cls in the simulator
' but reads outputs instead of inputs.

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
    On Error Resume Next

    Dim http As MSXML2.XMLHTTP
    Set http = ioStatusHttp

    If http Is Nothing Then Exit Sub
    If http.readyState <> 4 Then Exit Sub

    Dim responseText As String
    responseText = http.responseText

    Call ApplyOutputsFromCsv(responseText)
End Sub

Private Sub ApplyOutputsFromCsv(responseText As String)
    Dim aoutLine As String
    aoutLine = ExtractValueAfterPrefix(responseText, "aout=")
    If Len(aoutLine) > 0 Then
        Dim aoutParts() As String
        aoutParts = Split(aoutLine, ",")
        Dim i As Long
        For i = LBound(aoutParts) To UBound(aoutParts)
            If i <= UBound(A_OUTPUT) Then
                Dim aValue As Integer
                aValue = CInt(Val(Trim$(aoutParts(i))))
                If aValue >= 0 And aValue <= 255 Then
                    A_OUTPUT(i) = CByte(aValue)
                End If
            End If
        Next i
    End If

    Dim doutLine As String
    doutLine = ExtractValueAfterPrefix(responseText, "dout=")
    If Len(doutLine) > 0 Then
        Dim doutParts() As String
        doutParts = Split(doutLine, ",")
        For i = LBound(doutParts) To UBound(doutParts)
            If i <= UBound(D_OUTPUT) Then
                Dim dValue As Integer
                dValue = CInt(Val(Trim$(doutParts(i))))
                If dValue >= 0 And dValue <= 255 Then
                    D_OUTPUT(i) = CByte(dValue)
                End If
            End If
        Next i
    End If
End Sub

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
