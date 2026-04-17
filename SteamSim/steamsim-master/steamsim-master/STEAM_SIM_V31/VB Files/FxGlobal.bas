Attribute VB_Name = "FxModule"
Option Explicit
' This file is included in all of the SAMPLE programs that access files.
Global FXControl As Integer

Sub CenterForm(frm As Form)
'Centers a form on the screen
frm.Top = (Screen.Height - frm.Height) / 2
frm.Left = (Screen.Width - frm.Width) / 2
End Sub

Function FileExists(FileName$) As Integer
  On Error GoTo FileError
  X = FileLen(FileName$)
  FileExists = True
  Exit Function
FileError:
  FileExists = False
  Exit Function
End Function

' Returns the FXTools Directory
Function FXToolsDir() As String
  FXToolsDir = App.Path & "\..\..\"
End Function

Sub InitForm(frm As Form)
' Make the form 640x480 pixels centered.
frm.WindowState = 0
frm.Width = 9600
frm.Height = 7320
frm.Top = ((Screen.Height - frm.Height) / 2)
frm.Left = ((Screen.Width - frm.Width) / 2)
End Sub

' Returns the Media Directory
Function MediaDir() As String
  MediaDir = App.Path & "\..\..\Media\"
End Function

