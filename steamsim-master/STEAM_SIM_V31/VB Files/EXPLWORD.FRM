VERSION 5.00
Begin VB.Form ExplWord 
   Caption         =   "ExplWord - VB 4.0 (32 bits)"
   ClientHeight    =   3930
   ClientLeft      =   510
   ClientTop       =   4305
   ClientWidth     =   6675
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   6675
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   420
      Top             =   840
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   120
      Width           =   2955
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2820
      TabIndex        =   2
      Top             =   2460
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4860
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   1395
   End
   Begin VB.CommandButton Test 
      Caption         =   "Test"
      Height          =   435
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   420
      Top             =   120
   End
   Begin VB.Menu FICHIER 
      Caption         =   "Fichier"
      Begin VB.Menu EXIT 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu DATABASE 
      Caption         =   "Database"
      Begin VB.Menu LECTURE 
         Caption         =   "Read"
      End
      Begin VB.Menu ECRITURE 
         Caption         =   "Write"
      End
   End
   Begin VB.Menu AIDE 
      Caption         =   "?"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "ExplWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem ************************************************************************
Rem *                    [Généralitées du module]                          *
Rem ************************************************************************
Rem * nom du module .....: EXPLWORD.C                                      *
Rem *----------------------------------------------------------------------*
Rem * version ...........: 2.9                                             *
Rem *----------------------------------------------------------------------*
Rem * Outil de développement : VB 4.0 32 bits                              *
Rem *----------------------------------------------------------------------*
Rem * projet ............:                                                 *
Rem *----------------------------------------------------------------------*
Rem * Date de creation .: 21/08/96                                         *
Rem *----------------------------------------------------------------------*
Rem * Description ......: Exemple d'utilisation d'une carte Applicom       *
Rem *                     en SDK WINDOWS 3.11                              *
Rem *                     Lecture et ecriture dans la DataBase de la carte *
Rem *----------------------------------------------------------------------*
Rem *auteur ............: B LAURET                                         *
Rem ************************************************************************
Rem *               [ Rubrique modifications du module ]                   *
Rem ************************************************************************
Rem * Date .............:                                                  *
Rem *----------------------------------------------------------------------*
Rem * intervenant ......:                                                  *
Rem *----------------------------------------------------------------------*
Rem * description ......:                                                  *
Rem *----------------------------------------------------------------------*
Rem * numéro des lignes :                                                  *
Rem ************************************************************************
Rem *                               [ NOTES ]                              *
Rem ************************************************************************
Rem * note 1 : le getWord est effectué dans le timer tant que le status est*
Rem *          OK. Les valeurs de retours sont affichées à l'écran         *
Rem *          le setWord est effectué dans le boutton OK de la fenêtre    *
Rem *          SetWord. Si status OK alors affichage d'un message sinon    *
Rem *          affichage du status                                         *
Rem *----------------------------------------------------------------------*

Rem ************************************************************************
Sub LoopTest()

End Sub


Private Sub barre_Click()
    
End Sub

Private Sub About_Click()
    FORM_ABOUT.Show 1
End Sub

Private Sub Command1_Click()
Text1.text = "Start"

End Sub

Private Sub ECRITURE_Click()
     FORM_SETWORD.Show 1
     
End Sub


Private Sub EXIT_Click()
    End
End Sub

Private Sub Form_Load()
    Dim text As String
    Dim buffer As String
    Rem Initialisation du dialogue avec APPLICOM.DLL
    Call initbus(I_status)
    Rem Test du retour de status
    If I_status <> 0 Then
          text = " Initbus erreur - status = "
          buffer = Str(I_status)
          text = text + buffer
          reponse = MsgBox(text, vbExclamation)
          End
    End If
End Sub


Private Sub Form_Terminate()
    Rem Fin du dialogue avec APPLICOM.DLL
    Call exitbus(I_status)
End Sub


Private Sub LECTURE_Click()
     FORM_GETWORD.Show 1
End Sub


Private Sub Test_Click()
Dim i As Long
Dim flash As Boolean
Text1.text = "start"
Dim chan As Integer
Dim nes As Integer
Dim nb As Integer
Dim tadr As Long
Dim tabl As Integer
Dim stat As Integer

    'Call getword(I_ncard, I_nb, I_adr, I_tabl(1), I_status)
    chan = 0 'we use PC1500 which has 1 channel (0)
    
' FOR DO...( Takes 11 seconds )
    nes = 2 'node number for this particular DO/DI block
    nb = 8 'number of data items (in this case bits as we are using writepackqbit

' FOR AO...( Takes 5 seconds )
'    nes = 1 'node number for this particular DO/DI block
'    nb = 16 'number of data items (in this case bits as we are using writepackqbit
    
    tadr = 0 'start address of bit to be written. always 0 for profibus dp


Text1.text = "Start"

For i = 1 To 1500
    
  
    
    If flash Then
            tabl = &H55

        ' Output 55Hex
         Call writepackqbit(chan, nes, nb, tadr, tabl, stat)
        If stat <> 0 Then
            i = 1500
            Text1.text = stat
        End If
    Else
        
            tabl = &HAA

        ' Output AAHex
        Call writepackqbit(chan, nes, nb, tadr, tabl, stat)
        ' output AAHex
        If stat <> 0 Then
            i = 1500
            Text1.text = stat
        End If
    
    End If
    
    flash = Not flash
    
Next i

If stat = 0 Then
    Text1.text = "Done"         ' Takes 15 seconds for 1500 flashes
End If

End Sub

Private Sub Timer1_Timer()

Dim text As String
Dim buffer As String
Dim i As Integer
        
    CurrentY = 30
    Rem si demande de lecture
    If GetWord_OK = True Then
    
        Text2.text = "GetWord_OK = True"
    
        Rem lecture effectuée
        Call getword(I_ncard, I_nb, I_adr, I_tabl(1), I_status)
        Rem test du retour de status
        If I_status <> 0 Then
            Rem erreur dans la communication
            Rem fin de demande de lecture
            Rem avec les paramètres actuels
            GetWord_OK = False
            Rem affichage d'un message
            text = " status = "
            buffer = Str(I_status)
            text = text + buffer
            reponse = MsgBox(text, vbExclamation)
        Else
            Rem Lecture OK on rafraichie l'écran
            Refresh
            For i = 1 To I_nb
                Rem affichage des valeurs lues
                buffer = Str(I_adr + i - 1)
                text = " Valeur mot " + buffer + " = "
                buffer = Str(I_tabl(i))
                text = text + buffer
                CurrentX = 25
                CurrentY = CurrentY + 20
                Print text
            Next i
        End If
    End If
    
End Sub


Private Sub Timer2_Timer()

Dim i As Long
Dim flash As Boolean
Text1.text = "start"
Dim chan As Integer
Dim nes As Integer
Dim nb As Integer
Dim tadr As Long
Dim tabl As Integer
Dim stat As Integer

    chan = 0 'we use PC1500 which has 1 channel (0)
    

' FOR AO...
    nes = 1 'node number for this particular DO/DI block
    nb = 16 'number of data items (in this case bits as we are using writepackqbit
    
    tadr = 0 'start address of bit to be written. always 0 for profibus dp

    tabl = &H55

    Call writepackqbit(chan, nes, nb, tadr, tabl, stat)
    If stat <> 0 Then
        Text1.text = stat
    End If

End Sub
