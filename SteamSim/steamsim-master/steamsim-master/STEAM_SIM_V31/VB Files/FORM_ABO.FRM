VERSION 5.00
Begin VB.Form FORM_ABOUT 
   Caption         =   "ExplWord - About"
   ClientHeight    =   1380
   ClientLeft      =   945
   ClientTop       =   2955
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1380
   ScaleWidth      =   5040
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Picture         =   "FORM_ABO.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   1575
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright © 1996"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "ExplWord - VB 4.0 (32bits)   Version 1.0 "
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "FORM_ABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload FORM_ABOUT
End Sub


