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
    If ioioXmlHttp.readyState = 4 Then
        ' Belt-and-suspenders cleanup: if parsing throws, we still nil the handle so
        ' Request_IOIO_Data isn't blocked. Pairs with the unconditional reset over there.
        On Error Resume Next

        Dim responseText As String
        responseText = ioioXmlHttp.responseText

        If Not TryLoadAnalogInputsFromXml(responseText) Then
            Call TryLoadAnalogInputsFromJson(responseText)
        End If

        Set ioioXmlHttp = Nothing
        On Error GoTo 0
    End If
End Sub

Private Function TryLoadAnalogInputsFromXml(responseText As String) As Boolean
    Dim xmlDoc As MSXML2.DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    
    If Not xmlDoc.LoadXML(responseText) Then
        Exit Function
    End If

    Dim nodeList As MSXML2.IXMLDOMNodeList
    Set nodeList = xmlDoc.selectNodes("/pins/pin")
    
    If nodeList Is Nothing Then
        Exit Function
    End If
    
    Dim node As MSXML2.IXMLDOMNode
    Dim loadedAny As Boolean
    
    For Each node In nodeList
        Dim nameNode As MSXML2.IXMLDOMNode
        Dim calibratedNode As MSXML2.IXMLDOMNode
        
        Set nameNode = node.selectSingleNode("@name")
        Set calibratedNode = node.selectSingleNode("@calibrated")
        
        If Not nameNode Is Nothing Then
            If Not calibratedNode Is Nothing Then
                If ApplyNamedAnalogInput(nameNode.text, calibratedNode.text) Then
                    loadedAny = True
                End If
            End If
        End If
    Next node

    TryLoadAnalogInputsFromXml = loadedAny
End Function

Private Function TryLoadAnalogInputsFromJson(responseText As String) As Boolean
    Dim arrayText As String
    arrayText = ExtractJsonArray(responseText, "analog_input")
    
    If Len(arrayText) = 0 Then arrayText = ExtractJsonArray(responseText, "analog_inputs")
    If Len(arrayText) = 0 Then arrayText = ExtractJsonArray(responseText, "analogInput")
    If Len(arrayText) = 0 Then arrayText = ExtractJsonArray(responseText, "analogInputs")
    If Len(arrayText) = 0 Then Exit Function
    
    Dim parts() As String
    parts = Split(arrayText, ",")
    
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        If ApplyArrayAnalogInput(i, parts(i)) Then
            TryLoadAnalogInputsFromJson = True
        End If
    Next i
End Function

Private Function ExtractJsonArray(responseText As String, key As String) As String
    Dim keyToken As String
    keyToken = """" & key & """"
    
    Dim keyPos As Integer
    keyPos = InStr(1, responseText, keyToken, vbTextCompare)
    If keyPos = 0 Then Exit Function
    
    Dim startPos As Integer
    startPos = InStr(keyPos, responseText, "[")
    If startPos = 0 Then Exit Function
    
    Dim endPos As Integer
    endPos = InStr(startPos + 1, responseText, "]")
    If endPos = 0 Then Exit Function
    
    ExtractJsonArray = Mid$(responseText, startPos + 1, endPos - startPos - 1)
End Function

Private Function ApplyNamedAnalogInput(name As String, calibrated As String) As Boolean
    If Not startsWith(name, "ain") Then Exit Function
    If Len(Trim$(calibrated)) = 0 Then Exit Function
    
    Dim temp As Integer
    temp = CInt(Val(calibrated))
    If temp < 0 Or temp > 255 Then Exit Function
    
    Dim sourceIndex As Integer
    Dim commaPosition As Integer
    commaPosition = InStr(name, ",")
    
    If commaPosition = 0 Then
        sourceIndex = CInt(Val(Mid$(name, 4)))
    Else
        sourceIndex = CInt(Val(Mid$(name, 4, commaPosition - 4)))
    End If
    
    ApplyNamedAnalogInput = StoreAnalogInputValue(sourceIndex, IOIO_PIN_NAME_BASE, CByte(temp))
End Function

Private Function ApplyArrayAnalogInput(sourceIndex As Integer, rawValue As String) As Boolean
    Dim normalizedValue As String
    normalizedValue = Trim$(Replace$(rawValue, """", ""))
    If Len(normalizedValue) = 0 Then Exit Function
    If LCase$(normalizedValue) = "null" Then Exit Function
    
    Dim temp As Integer
    temp = CInt(Val(normalizedValue))
    If temp < 0 Or temp > 255 Then Exit Function
    
    ApplyArrayAnalogInput = StoreAnalogInputValue(sourceIndex, IOIO_JSON_ARRAY_BASE, CByte(temp))
End Function

Private Function StoreAnalogInputValue(sourceIndex As Integer, sourceBase As Integer, Value As Byte) As Boolean
    Dim targetIndex As Integer
    targetIndex = sourceIndex - sourceBase + IOIO_INPUT_TARGET_BASE
    
    If targetIndex < 0 Then Exit Function
    If targetIndex > UBound(A_INPUT) Then Exit Function
    
    A_INPUT(targetIndex) = Value
    StoreAnalogInputValue = True
End Function


Public Function startsWith(str As String, prefix As String) As Boolean
    Dim location As Integer
    location = InStr(1, str, prefix)
    
    If location = 1 Then
        startsWith = True
    Else
        startsWith = False
    End If
End Function

