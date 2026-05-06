VERSION 5.00
Begin VB.MDIForm frm_Template 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7650
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7800
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnu_plot2 
      Caption         =   "Dynamics"
   End
End
Attribute VB_Name = "frm_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_plot2_Click()
    Template.Show
End Sub
