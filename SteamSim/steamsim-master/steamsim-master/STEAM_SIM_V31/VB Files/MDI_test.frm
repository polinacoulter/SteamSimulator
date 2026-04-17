VERSION 5.00
Begin VB.MDIForm MDI_test 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   900
      Top             =   480
   End
   Begin VB.Menu mnu_Ballast 
      Caption         =   "Ballast"
   End
End
Attribute VB_Name = "MDI_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_Ballast_Click()


Ballast_form.Show


End Sub
