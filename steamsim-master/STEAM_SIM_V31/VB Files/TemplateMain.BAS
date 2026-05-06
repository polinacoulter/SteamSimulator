Attribute VB_Name = "TemplateModule"
'Public Const idTemplate As Integer = 0
'Global ScreenSettings(10) As Integer
Type ScreenSize
    ' The current size of a form
    Width As Integer
    Height As Integer
    ' The minimum allowed size of a form
    MinWidth As Integer
    MinHeight As Integer
End Type

' The Const's representing each screen
Global Const MaxScreens = 1 ' NOTE, this must always be the largest!


Global Const idTemplate As Integer = 1


' Holds each of the screens that are defined below
Global ScreenSettings(MaxScreens) As ScreenSize

Public Sub InitialiseScreens()
    Dim i As Integer
    
    For i = 1 To MaxScreens
        ' Make each screens default start size be a quarter of the screen size
        ' also make the minimum size the same
        With ScreenSettings(i)
            .Width = (Screen.Width / 2)
            Debug.Print "Screen.Width = " & Screen.Width
            .Height = (Screen.Height / 2)
            .MinWidth = .Width
            .MinHeight = .Height
        End With
    Next i
    
    ' Override (minimum) sizes here for particular forms
    ' ScreenSettings(idalarms).Width = 8200
    ' ScreenSettings(idalarms).Height = 14000
     
   
End Sub

Public Sub HandleScreenSizing(frmMe As Form, frmId As Integer)
    ' When a screen is resized, then control comes here to check that the minimum width
    ' of the screen is not exceeded.  The new height is calculated from the old
    ' to keep the aspect ratio the same, and then these values are recorded
    
    Debug.Print " frmMe.width = " & frmMe.Width & " frmId = " & frmId
    Debug.Print "ScreenSettings(frmId).Width) = " & ScreenSettings(frmId).Width
 
    If frmMe.WindowState = 0 Then
        If frmMe.Width < ScreenSettings(frmId).MinWidth Then frmMe.Width = ScreenSettings(frmId).MinWidth
        Debug.Print " frmMe.width = " & frmMe.Width & " ScreenSettings(frmId).MinWidth = " & ScreenSettings(frmId).MinWidth
        ' new height=orig height * (current width / last width)
        frmMe.Height = ScreenSettings(frmId).Height * (frmMe.Width / ScreenSettings(frmId).Width)
        ScreenSettings(frmId).Width = frmMe.Width
        ScreenSettings(frmId).Height = frmMe.Height
    End If
End Sub
