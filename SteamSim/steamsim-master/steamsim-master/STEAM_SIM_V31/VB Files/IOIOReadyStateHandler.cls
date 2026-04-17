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
        'Debug.Print "Read Values"
        
        Dim xmlDoc As MSXML2.DOMDocument
        Set xmlDoc = ioioXmlHttp.responseXML

        Dim nodeList As MSXML2.IXMLDOMNodeList
        Set nodeList = xmlDoc.selectNodes("/pins/pin")
        
        If Not nodeList Is Nothing Then
            Dim node As MSXML2.IXMLDOMNode
            Dim name As String
            Dim calibrated As String
            
            For Each node In nodeList
                name = node.selectSingleNode("@name").text
                calibrated = node.selectSingleNode("@calibrated").text
                
                If startsWith(name, "ain") Then
                    Dim Value As Byte
                    Dim Index As Integer
                    Dim temp As Integer
                    
                    temp = CInt(Val(calibrated))
                    If temp >= 0 And temp <= 255 Then
                        Value = CByte(CInt(Val(calibrated)))
                        Dim commaPosition As Integer
                        commaPosition = InStr(name, ",")
                        If (commaPosition = 0) Then
                            'No Comma means we are of format ain18 or ain123
                            Index = CInt(Val(Mid$(name, 4)))
                        Else
                            Index = CInt(Val(Mid$(name, 4, commaPosition - 4)))
                        End If
                        
                        If Index < UBound(A_INPUT) Then
                            A_INPUT(Index) = Value
                        End If
                        
                        'Dim output As String
                        'output = "Detected value " & value & " on index " & index
                        ''MsgBox output
                    End If
                    
                End If
                
            Next node
        End If
        
        Set ioioXmlHttp = Nothing

        'MsgBox "Done"
    End If
End Sub


Public Function startsWith(str As String, prefix As String) As Boolean
    Dim location As Integer
    location = InStr(1, str, prefix)
    
    If location = 1 Then
        startsWith = True
    Else
        startsWith = False
    End If
End Function

