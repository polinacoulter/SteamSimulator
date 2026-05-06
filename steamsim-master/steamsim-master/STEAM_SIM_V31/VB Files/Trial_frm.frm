VERSION 5.00
Begin VB.Form Trial_frm 
   Caption         =   "Trial Version"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "Trial_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   6795
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label TRIAL_HAS_ENDED_label 
      Alignment       =   2  'Center
      Caption         =   "TRIAL HAS ENDED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   1980
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Warning: Setting the system date back will end the trial period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "The"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   300
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "day trial began on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1800
      TabIndex        =   2
      Top             =   300
      Width           =   1995
   End
   Begin VB.Label Start_Date_label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3780
      TabIndex        =   1
      Top             =   300
      Width           =   2535
   End
   Begin VB.Label TRial_label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1260
      TabIndex        =   0
      Top             =   300
      Width           =   495
   End
End
Attribute VB_Name = "Trial_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me


End Sub

Private Sub Form_Load()
Dim SSPRO_start_date_LONG As Long
Dim SSPRO_start_date_date As Date


Me.Left = frm_Host.Width / 4
Me.Top = frm_Host.Height / 4
Me.Height = 2925
Me.Width = 6915

'=== 03/02/03 ===
If TRIAL_HAS_ENDED = True Then
    TRIAL_HAS_ENDED_label.Visible = True
End If



' === 03/02/03 ===
TRial_label.Caption = Format(Timeout_period, "#0")


SSPRO_start_date_LONG = SSPRO_start_date

' Add 30000 so that the date can be correctly converted back to a date...
SSPRO_start_date_LONG = SSPRO_start_date_LONG + 30000

SSPRO_start_date_date = CDate(SSPRO_start_date_LONG)

Start_Date_label.Caption = Format(SSPRO_start_date_date, "Long Date")



End Sub
