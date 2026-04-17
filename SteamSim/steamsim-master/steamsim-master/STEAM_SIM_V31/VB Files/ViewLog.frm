VERSION 5.00
Begin VB.Form ViewLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7410
   Begin VB.OLE OLE1 
      AutoActivate    =   0  'Manual
      BackColor       =   &H80000018&
      Class           =   "WordPad.Document.1"
      Height          =   2805
      Left            =   120
      OLETypeAllowed  =   0  'Linked
      SizeMode        =   2  'AutoSize
      TabIndex        =   0
      Top             =   240
      Width           =   8370
   End
End
Attribute VB_Name = "ViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
