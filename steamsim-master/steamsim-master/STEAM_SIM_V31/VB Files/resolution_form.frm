VERSION 5.00
Begin VB.Form resolution_form 
   Caption         =   "Screen Resolution"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   1320
      Width           =   1875
   End
   Begin VB.OptionButton Option1 
      Caption         =   "High Resolution   (1280x1024)"
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   2955
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Low Resolution    ( 800x600)"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Value           =   -1  'True
      Width           =   2955
   End
End
Attribute VB_Name = "resolution_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Option1(0).Value = True Then

   Debug.Print "Low Res"
   
   ' NB Should be frm_intro_low_res.Show
   frm_intro.Show
 ' frm_intro_low_res.Show
   
Else

   Debug.Print "High Res"
   frm_intro.Show
   
End If

End Sub
