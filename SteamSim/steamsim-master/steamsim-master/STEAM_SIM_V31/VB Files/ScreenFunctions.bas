Attribute VB_Name = "ScreenFunctions"
'Constant Declaration
Global Const MAXSCRS = 300 ' Maximum number of forms in project

' 2/8/99 RWJ ( feed water has 249 , need MORE )
'Global Const MAXOBJECTS = 250 ' Maximum number of objects in form
'
' 1/11/99
'Global Const MAXOBJECTS = 350 ' Maximum number of objects in form
Global Const MAXOBJECTS = 500 ' Maximum number of objects in form


'STRUCTURE DECLARATION
' Defines the contained object's sizes
Type JYAObjSizes
    iWidth As Long
    iHeight As Long
    iTop As Long
    iLeft As Long
    sName As String
    sFontSize As Integer
End Type

' Defines the control's width and height
Type JYAScrContent
    iWidth As Long
    iHeight As Long
    iMinWidth As Long
    iMinHeight As Long
    objcount As Integer
    iFlagLoaded As Boolean 'True if form is finally loaded
    iFlagHasResized As Boolean 'True if form has resized
    JYAobjsize(MAXOBJECTS) As JYAObjSizes
End Type

Global JYAscrsizeOrig(MAXSCRS) As JYAScrContent

Public Sub JYAScreenInit(ifrm As Form, iId As Integer)
    ' The following is required
    On Error Resume Next
    Dim foObj As Object
    Dim iCount As Integer
    Dim iDiff As Integer
    Dim sObjName As String ' The name of the object
    'Set flag to false it will be set to true after first resize event has occured
    JYAscrsizeOrig(iId).iFlagLoaded = False
    'Record form size Note: This is not required if HavenStretch OCX is present
    JYAscrsizeOrig(iId).iHeight = ifrm.ScaleHeight
    JYAscrsizeOrig(iId).iWidth = ifrm.ScaleWidth
    'Process content of form
    iCount = 0
    For Each foObj In ifrm
    'Increment counter
        iCount = iCount + 1
    'Change font type
        If foObj.Font <> "Symbol" Then
            foObj.Font = "Arial"
        End If
    'Handle Error
        Err.clear
    ' Set all fonts as normal
        foObj.Font.Bold = False
    'Handle Error
        Err.clear
    'Record object name
        sObjName = TypeName(foObj)
        JYAscrsizeOrig(iId).JYAobjsize(iCount).sName = sObjName
    'Record font size
        JYAscrsizeOrig(iId).JYAobjsize(iCount).sFontSize = foObj.Font.Size
    'Handle Error
        Err.clear
    ' Record object position and size
        Select Case sObjName
        Case Is = "Line" 'TypeOf objCur Is Line
            With foObj
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight = .Y2
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth = .X2
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft = .X1
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop = .Y1
                JYAscrsizeOrig(iId).JYAobjsize(iCount).sName = .Name
            End With
        Case Is = "SysInfo"
            With foObj
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft = .X1
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop = .Y1
                JYAscrsizeOrig(iId).JYAobjsize(iCount).sName = .Name
            End With
        
        Case Is = "HavenStretch"
            With foObj
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight = .Height
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth = .Width
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft = .Left
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop = .Top
                JYAscrsizeOrig(iId).JYAobjsize(iCount).sName = .Name
            ' Use HavenStretch width and height properties to record size and width of form
                If .Name = "HavenStretch1" Then
                    JYAscrsizeOrig(iId).iHeight = .Height + ifrm.Height - ifrm.ScaleHeight
                    JYAscrsizeOrig(iId).iWidth = .Width + ifrm.Width - ifrm.ScaleWidth
                    JYAscrsizeOrig(iId).iMinWidth = JYAscrsizeOrig(iId).iWidth
                    JYAscrsizeOrig(iId).iMinHeight = JYAscrsizeOrig(iId).iHeight
                End If
            End With
        
        ' === 28/01/02 ADDED THIS ======
        Case Is = "Frame"
            With foObj
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight = .Height
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth = .Width
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft = .Left
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop = .Top
                JYAscrsizeOrig(iId).JYAobjsize(iCount).sName = .Name
            ' Use HavenStretch width and height properties to record size and width of form
                If .Name = "HavenStretch1" Then
                    JYAscrsizeOrig(iId).iHeight = .Height + ifrm.Height - ifrm.ScaleHeight
                    JYAscrsizeOrig(iId).iWidth = .Width + ifrm.Width - ifrm.ScaleWidth
                    JYAscrsizeOrig(iId).iMinWidth = JYAscrsizeOrig(iId).iWidth
                    JYAscrsizeOrig(iId).iMinHeight = JYAscrsizeOrig(iId).iHeight
                End If
            End With
        ' === 28/01/02 END OF EDIT ===
        
        Case Else
            With foObj
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight = .Height
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth = .Width
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft = .Left
                JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop = .Top
                JYAscrsizeOrig(iId).JYAobjsize(iCount).sName = .Name
            End With
        End Select
        
     Next foObj
    'Record number of objects in current form
    JYAscrsizeOrig(iId).objcount = iCount
    'Record minimum form size as about a quarter of screen size or smaller
    If JYAscrsizeOrig(iId).iWidth < Screen.Width / 2 Then
        JYAscrsizeOrig(iId).iMinWidth = JYAscrsizeOrig(iId).iWidth
    Else
        JYAscrsizeOrig(iId).iMinWidth = Screen.Width / 2
    End If
    JYAscrsizeOrig(iId).iMinHeight = JYAscrsizeOrig(iId).iHeight * JYAscrsizeOrig(iId).iMinWidth / JYAscrsizeOrig(iId).iWidth
    ' Modify form size from previously recorded values
    ifrm.Width = JYAscrsizeOrig(iId).iWidth
    ifrm.Height = JYAscrsizeOrig(iId).iHeight
    JYAscrsizeOrig(iId).iFlagLoaded = True
End Sub

Public Sub JYAScreenResize(ifrm As Form, iId As Integer)
    ' The following is required
    On Error Resume Next
    Dim foObj As Object
    Dim iCount As Integer
    
    Dim sObjName As String ' The name of the object
    Dim dRatioWidth As Single
    Dim dRatioHeight As Single
    
    'Only resize if form is loaded
    If JYAscrsizeOrig(iId).iFlagLoaded = True Then
        JYAscrsizeOrig(iId).iFlagHasResized = False
        'Detect if new window is to be set to minimum size
        If (ifrm.Width < JYAscrsizeOrig(iId).iMinWidth) Or (ifrm.Height < JYAscrsizeOrig(iId).iMinHeight) Then
            ifrm.Height = JYAscrsizeOrig(iId).iMinHeight
        End If
        'Maintain ratio constant
        ifrm.Width = JYAscrsizeOrig(iId).iWidth * ifrm.Height / JYAscrsizeOrig(iId).iHeight
        If JYAscrsizeOrig(iId).iFlagHasResized = False Then
            'Process content of form
            iCount = 0
            'Calculates new ratio
            dRatioWidth = ifrm.Width / JYAscrsizeOrig(iId).iWidth
            dRatioHeight = ifrm.Height / JYAscrsizeOrig(iId).iHeight
            
            For Each foObj In ifrm
                iCount = iCount + 1
                    
            'Process and resize object
                sObjName = TypeName(foObj)
                Select Case sObjName
                Case Is = "Line" 'TypeOf objCur Is Line
                    With foObj
                        .X1 = JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft * dRatioWidth
                        .X2 = JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth * dRatioWidth
                        .Y1 = JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop * dRatioHeight
                        .Y2 = JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight * dRatioHeight
                    End With
                ' Catch a Timer, as it cannot be resized/moved as run time
                Case Is = "Timer"
                ' Catch a SysInfo object, as it cannot be resized/moved as run time
                Case Is = "SysInfo"
                Case Is = "Pump" 'Do not recalculate height as it is resized automatically in the OCX
                    With foObj
                        .Left = JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft * dRatioWidth
                        .Top = JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop * dRatioHeight
                        .Width = JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth * dRatioWidth
                    End With
                    
                '1/9/99 DO NOT NEED:
                ' Case Is = "Gauge"   ' 6/8/99
                 '   foObj.Value = foObj.Value
                 '   foObj.Resize JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight * dRatioHeight
                 '   foObj.Left = JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft * dRatioWidth
                 '   foObj.Top = JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop * dRatioHeight
                 
                
                Case Else
                    With foObj
                        .Left = JYAscrsizeOrig(iId).JYAobjsize(iCount).iLeft * dRatioWidth
                        .Top = JYAscrsizeOrig(iId).JYAobjsize(iCount).iTop * dRatioHeight
                        .Width = JYAscrsizeOrig(iId).JYAobjsize(iCount).iWidth * dRatioWidth
                        .Height = JYAscrsizeOrig(iId).JYAobjsize(iCount).iHeight * dRatioHeight
                    End With
                End Select
            'Handle Error
                Err.clear
                
                foObj.Font.Size = JYAscrsizeOrig(iId).JYAobjsize(iCount).sFontSize * dRatioHeight
            'Handle Error
                Err.clear
            Next foObj
        End If
        JYAscrsizeOrig(iId).iFlagHasResized = True
    End If
'Note: after this the software will generate another resize control
End Sub

