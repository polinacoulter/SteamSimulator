VERSION 5.00
Object = "{62E9D822-2EAC-11D1-8915-000021240899}#2.1#0"; "HavenStretch.ocx"
Begin VB.Form Template 
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   13140
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6615
      Left            =   2280
      Top             =   360
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   11668
      EditContents    =   0   'False
   End
End
Attribute VB_Name = "Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub form_Load()

    Me.Width = ScreenSettings(idsy064).Width
    
   
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        Call HandleScreenSizing(Me, idsy064)
        
        If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
            HavenStretch1.Visible = False
            HavenStretch1.Width = Me.ScaleWidth
            HavenStretch1.Height = Me.ScaleHeight
            HavenStretch1.Visible = True
        End If
    End If
End Sub











