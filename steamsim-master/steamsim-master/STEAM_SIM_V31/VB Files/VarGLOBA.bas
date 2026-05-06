Attribute VB_Name = "VarGLOBAL"
Rem ************************************************************************
Rem *                               [ Variables Globales ]                 *
Rem ************************************************************************

Rem * Standard Variables
Public GetWord_OK As Boolean
Public reponse As Integer
Rem * Input Variables : GetWord
Public I_ncard As Integer 'Applicom Board Number / Numéro de carte applicom  (1-8)
Public I_nb As Integer    'Number of words te be read / Nombre de mot à lire ( 1 -128)
Public I_adr As Integer   'Address of the first word to be read / adresse du premier mot à lire (0-32767)
Public I_tabl(128) As Integer 'table receiving the data / tableau recevant les données
Public I_status As Integer 'Interchange status / Status de l'échange
Rem * Output Variables : SetWord
Public O_ncard As Integer 'Applicom Board Number / Numéro de carte applicom  (1-8)
Public O_nb As Integer    'Number of words te be read / Nombre de mot à lire ( 1 -128)
Public O_adr As Integer   'Address of the first word to be read / adresse du premier mot à lire (0-32767)
Public O_tabl(128) As Integer 'table receiving the data / tableau recevant les données
Public O_status As Integer 'Interchange status / Status de l'échange

Global RUN_AO As Boolean
Global AO_value As Integer
Global Timer_count As Integer
