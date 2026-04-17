VERSION 5.00
Begin VB.Form FORM_GETWORD 
   Caption         =   "GETWORD"
   ClientHeight    =   2055
   ClientLeft      =   2985
   ClientTop       =   1890
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2055
   ScaleWidth      =   3660
   Begin VB.TextBox V_NCARD 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox V_NB 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton B_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton B_CANCEL 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox V_ADR 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label T_NCARD 
      Alignment       =   1  'Right Justify
      Caption         =   "Numéro carte (1-8)"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label T_ADR 
      Alignment       =   1  'Right Justify
      Caption         =   "adresse premier mot"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label T_VALEUR 
      Alignment       =   1  'Right Justify
      Caption         =   "nb mot"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "FORM_GETWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub B_CANCEL_Click()
    Rem on ferme la fenêtre GETWORD
    Unload FORM_GETWORD
End Sub


Private Sub B_OK_Click()
    Rem récuperation des valeurs
    I_ncard = Val(V_NCARD)
    I_adr = Val(V_ADR)
    I_nb = Val(V_NB)
    Rem Getword peut être effectué
    GetWord_OK = True
    Rem on ferme la fenetre GetWord
    Unload FORM_GETWORD
End Sub


