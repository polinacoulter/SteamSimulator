Attribute VB_Name = "IOAppNet"
Option Explicit

' HTTP plumbing for the I/O App. Inverse of what the simulator does:
'   GET  /ioio/status -> parse aout= and dout= -> A_OUTPUT() / D_OUTPUT()
'   POST /ioio/inputs -> body "ain=...&din=..."  from A_INPUT() / D_INPUT()
'
' Both calls are async fire-and-forget. WinINet caches GETs, so we use the
' same cache-buster pattern as the simulator's Request_IOIO_Data.

Public Sub Get_Outputs_From_Python()
    ' Drop any previous request unconditionally so a stuck callback can't
    ' permanently freeze polling. Same pattern as simulator's Request_IOIO_Data.
    On Error Resume Next
    Set ioStatusHttp = Nothing
    On Error GoTo 0

    Set ioStatusHttp = New MSXML2.XMLHTTP

    If StatusHandler Is Nothing Then
        Set StatusHandler = New IOAppHandler
    End If

    Static cacheBuster As Long
    cacheBuster = cacheBuster + 1

    Dim url As String
    url = PYTHON_STATUS_URL & "?t=" & CStr(Timer) & "&n=" & CStr(cacheBuster)

    On Error Resume Next
    ioStatusHttp.Open "GET", url, True
    ioStatusHttp.setRequestHeader "Cache-Control", "no-cache"
    ioStatusHttp.setRequestHeader "Pragma", "no-cache"
    ioStatusHttp.OnReadyStateChange = StatusHandler
    ioStatusHttp.send
    On Error GoTo 0
End Sub

Public Sub Send_Inputs_To_Python()
    ' Build "ain=v0,v1,...&din=v0,v1,..." from the in-memory arrays and POST.
    ' During Phase 2 (no Profibus yet), A_INPUT and D_INPUT will be all zeros
    ' which is fine for verifying the wire - Phase 3 fills them from real cards.
    On Error Resume Next
    Set ioInputsHttp = Nothing
    On Error GoTo 0

    Set ioInputsHttp = New MSXML2.XMLHTTP

    Dim ainParts() As String
    ReDim ainParts(0 To UBound(A_INPUT))
    Dim i As Long
    For i = 0 To UBound(A_INPUT)
        ainParts(i) = CStr(A_INPUT(i))
    Next i

    Dim dinParts() As String
    ReDim dinParts(0 To UBound(D_INPUT))
    For i = 0 To UBound(D_INPUT)
        dinParts(i) = CStr(D_INPUT(i))
    Next i

    Dim body As String
    body = "ain=" & Join(ainParts, ",") & "&din=" & Join(dinParts, ",")

    On Error Resume Next
    ioInputsHttp.Open "POST", PYTHON_INPUTS_URL, True
    ioInputsHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    ioInputsHttp.send body
    On Error GoTo 0
End Sub
