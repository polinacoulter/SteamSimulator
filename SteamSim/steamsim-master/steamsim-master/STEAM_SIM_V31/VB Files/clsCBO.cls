VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoxSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const CB_ERR As Long = (-1)
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SHOWDROPDOWN As Long = &H14F

Private Declare Function SendMessageStr Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long

Public Sub FindIndexStr(ctlSource As Control, _
   ByVal str As String, intKey As Integer, _
   Optional ctlTarget As Variant)
Dim lngIdx As Long
Dim FindString As String

If (intKey < 32 Or intKey > 127) And _
   (Not (intKey = 13 Or intKey = 8)) Then Exit Sub

If Not intKey = 13 Or intKey = 8 Then
    If Len(ctlSource.Text) = 0 Then
        FindString = str & Chr$(intKey)
    Else
        FindString = Mid$(str, 1, ctlSource.SelStart) & Chr$(intKey)
    End If
End If

If intKey = 8 Then
   If Len(ctlSource.Text) = 0 Then Exit Sub
   Dim numChars As Integer
   numChars = ctlSource.SelStart - 1
    'FindString = Left$( str, numChars)
   If numChars > 0 Then FindString = Mid$(str, 1, numChars)
   'Else
    'If numChars <= 0 Then
     ' finstring = ""
      'ctlSource.Text = ""
    'End If
  'End If
End If

If IsMissing(ctlTarget) And TypeName(ctlSource) = "ComboBox" Then
    Set ctlTarget = ctlSource
    If intKey = 13 Then
        Call SendMessageStr(ctlTarget.hWnd, _
            CB_SHOWDROPDOWN, True, 0&)
        Exit Sub
    End If
    lngIdx = SendMessageStr(ctlTarget.hWnd, CB_FINDSTRING, -1, FindString)
    If lngIdx <> -1 Then
            ctlTarget.ListIndex = lngIdx
            ctlSource.SelStart = Len(FindString)
            ctlSource.SelLength = Len(ctlSource.Text) - ctlSource.SelStart
    End If
    intKey = 0
Else
    Exit Sub
End If
End Sub


