VERSION 5.00
Begin VB.MDIForm MDIBoiler 
   BackColor       =   &H8000000C&
   Caption         =   "Steam Sim"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFOserv 
      Caption         =   "FO Service"
   End
   Begin VB.Menu mnuBoiler 
      Caption         =   "Boiler"
   End
End
Attribute VB_Name = "MDIBoiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuBoiler_Click()

Boilers.Show

End Sub

Private Sub mnuFOserv_Click()

FO_Service.Show

End Sub
