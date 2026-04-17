VERSION 5.00
Begin VB.Form FORM_SETWORD 
   Caption         =   "SETWORD"
   ClientHeight    =   2115
   ClientLeft      =   2925
   ClientTop       =   3375
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2115
   ScaleWidth      =   3675
   Begin VB.TextBox V_ADR 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton B_CANCEL 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton B_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox V_VALEUR 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox V_NCARD 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label T_VALEUR 
      Alignment       =   1  'Right Justify
      Caption         =   "valeur mot"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label T_ADR 
      Alignment       =   1  'Right Justify
      Caption         =   "adresse mot (0-32767)"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label T_NCARD 
      Alignment       =   1  'Right Justify
      Caption         =   "Numéro carte (1-8)"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FORM_SETWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub B_CANCEL_Click()
    Unload FORM_SETWORD
End Sub


Private Sub B_OK_Click()
    Rem Recupération des valeurs
    O_ncard = Val(V_NCARD)
    O_adr = Val(V_ADR)
    O_tabl(1) = Val(V_VALEUR)
    Rem on ecrit une seule valeur
    O_nb = 1
    Rem fonction applicom SETWORD
    Call setword(O_ncard, O_nb, O_adr, O_tabl(1), O_status)
    Rem Vérification du status
    If O_status <> 0 Then
            GetWord_OK = False
            text = " status = "
            buffer = Str(O_status)
            text = text + buffer
            reponse = MsgBox(text, vbExclamation)
    Else
        reponse = MsgBox("Ecriture effectuée", vbInformation)
    End If
    Unload FORM_SETWORD
End Sub


