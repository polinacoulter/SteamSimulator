VERSION 5.00
Begin VB.Form Model_Constants 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Constants"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7275
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   795
      Left            =   5460
      TabIndex        =   20
      Top             =   2100
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   60
      Top             =   180
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   15
      X1              =   480
      X2              =   5100
      Y1              =   6390
      Y2              =   6390
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   14
      X1              =   450
      X2              =   5070
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Port Boiler Burner 1 Diesel Oil Burn Rate Multiplier  (0.1 to 1.0)"
      Height          =   405
      Index           =   12
      Left            =   540
      TabIndex        =   31
      Top             =   5910
      Width           =   3255
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3900
      TabIndex        =   30
      Top             =   5910
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Stbd Boiler Burner 1 Diesel Oil Burn Rate Multiplier  (0.1 to 1.0)"
      Height          =   405
      Index           =   11
      Left            =   540
      TabIndex        =   29
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3900
      TabIndex        =   28
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3900
      TabIndex        =   27
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   13
      X1              =   480
      X2              =   5100
      Y1              =   5340
      Y2              =   5340
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "FO Low Pressure S/by start  ( 1.0 to 500 psi )"
      Height          =   315
      Index           =   10
      Left            =   540
      TabIndex        =   26
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   12
      X1              =   480
      X2              =   5100
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "FO Low Pressure Trip  ( 1.0 to 500.0  psi )"
      Height          =   315
      Index           =   4
      Left            =   540
      TabIndex        =   25
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   11
      X1              =   480
      X2              =   5100
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3900
      TabIndex        =   24
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "3. Clicking the Update box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   570
      TabIndex        =   23
      Top             =   7590
      Width           =   4215
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3900
      TabIndex        =   22
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   10
      X1              =   480
      X2              =   5100
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ADT Tank Size ( 0.1 to 2.0 )"
      Height          =   315
      Index           =   9
      Left            =   540
      TabIndex        =   21
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   8
      X1              =   480
      X2              =   5100
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3900
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   4
      X1              =   480
      X2              =   480
      Y1              =   6360
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   9
      X1              =   480
      X2              =   5100
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Boiling Rate ( 1.0 to 100.0)"
      Height          =   315
      Index           =   8
      Left            =   540
      TabIndex        =   18
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   5
      X1              =   480
      X2              =   5100
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3900
      TabIndex        =   17
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship Load Multiplier   ( 0.1 to 2.0 )"
      Height          =   315
      Index           =   7
      Left            =   540
      TabIndex        =   16
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3900
      TabIndex        =   15
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3900
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   7
      X1              =   480
      X2              =   5100
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   6
      X1              =   480
      X2              =   5100
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Main Breaker Instant Trip (100 to 2000) kW"
      Height          =   315
      Index           =   6
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Main Breaker Timed Trip (100 to 2000) kW"
      Height          =   315
      Index           =   5
      Left            =   480
      TabIndex        =   12
      Top             =   2700
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5865
      Left            =   3840
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3900
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "White Smoke Multiplier    ( 0.1 to 2.0 )"
      Height          =   315
      Index           =   3
      Left            =   420
      TabIndex        =   10
      Top             =   2280
      Width           =   3315
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3900
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3900
      TabIndex        =   8
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Constant 
      Alignment       =   2  'Center
      Caption         =   "0.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3900
      TabIndex        =   7
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   480
      X2              =   5100
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   480
      X2              =   5100
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   480
      X2              =   5100
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   5100
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label5 
      Caption         =   "2. Any snapshot re-load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   570
      TabIndex        =   6
      Top             =   7230
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "1. A cold start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   570
      TabIndex        =   5
      Top             =   6870
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Note: The Current Value is changed (up-dated) on either of the following events:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   6510
      Width           =   6915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Current Value"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Propellor Pitch (metres)        ( 3.0 to 9.0 )"
      Height          =   315
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   1860
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship  Inertia Multiplier   ( 0.1 to 2.0 )"
      Height          =   315
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1440
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Shaft Inertia Multiplier   ( 0.1 to 2.0 )"
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1020
      Width           =   3255
   End
End
Attribute VB_Name = "Model_Constants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


' === 17/12/02  (Read from the file c:\Steam_Sim\Text\Constants.txt ) ===
Call SET_Constants


End Sub

Private Sub Form_Load()

Me.Width = 7305

' === 17/12/02 ===
'Me.Height = 6780

' === 31/03/03 ===
' === 15/07/09 ===
'Me.Height = 7680
Me.Height = 8685


End Sub

Private Sub Timer1_Timer()

'c             === 12/12/02 19.605 ===
'c        CONSTANT(5) = Main Breaker Trip Timed         (kW)
'c        CONSTANT(6) = Main Breaker Trip Instantaneous (kW)


'c     === 23/08/02 ===
'           PC_REALS(855) = Constant(1)
'           PC_REALS(856) = Constant(2)
'           PC_REALS(857) = Constant(3)

'           PC_REALS(859) = Constant(4)
'c          === 12/12/02 19.506 ===
'           PC_REALS(860) = Constant(5)
'           PC_REALS(861) = Constant(6)

    Constant(1).Caption = Format(LSS_reals(855), "#####0.00")
    Constant(2).Caption = Format(LSS_reals(856), "#####0.00")
    Constant(3).Caption = Format(LSS_reals(857), "#####0.00")
    
    ' === 24/09/02 ===
    Constant(4).Caption = Format(LSS_reals(859), "#####0.00")

' === 12/12/02 19.506 ===
    Constant(5).Caption = Format(LSS_reals(860), "#####0.0")
    Constant(6).Caption = Format(LSS_reals(861), "#####0.0")
    
    Constant(7).Caption = Format(LSS_reals(862), "#####0.0")
'     === 17/12/02 ===
'        CONSTANT(8) = Boiling Rate
    Constant(8).Caption = Format(LSS_reals(863), "#####0.0")


'     === 17/12/02 ===
'        CONSTANT(9) = ADT Tank Size Multiplier
    Constant(9).Caption = Format(LSS_reals(864), "#####0.0")
    
'     === 31/03/03 ===
'        CONSTANT(10) = FO Low Pressure Boiler trip point
    Constant(10).Caption = Format(LSS_reals(869), "#####0.0")
'     === 31/03/03 ===
'        CONSTANT(11) = FO Low Pressure Stand-by cut in
    Constant(11).Caption = Format(LSS_reals(870), "#####0.0")
    
'   === 15/07/09 ===
    Constant(12).Caption = Format(LSS_reals(871), "#####0.0")
    Constant(13).Caption = Format(LSS_reals(872), "#####0.0")
    
    

End Sub
