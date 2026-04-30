VERSION 5.00
Begin VB.Form frmIOApp
   Caption         =   "I/O App"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart
      Caption         =   "Start"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdStop
      Caption         =   "Stop"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox lstLog
      Height          =   3585
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6720
   End
   Begin VB.Timer Timer1
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   240
   End
   Begin VB.Label lblStatusValue
      Caption         =   "Stopped"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2000
   End
   Begin VB.Label lblStatusLabel
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   800
   End
End
Attribute VB_Name = "frmIOApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Phase 3: Profibus + HTTP.
' Start: try to initialize both Profibus cards (IO_Init); on failure show a
'        MsgBox like the original simulator did, but keep the timer running
'        in HTTP-only mode so the I/O App is still useful on dev machines
'        without Profibus hardware.
' Timer: per-card gated reads from Profibus, push to Python, pull from
'        Python, per-card gated writes to Profibus.
'
' We use a VB.ListBox for the log instead of a multi-line VB.TextBox because
' multi-line TextBoxes require a companion .frx file in VB6.

Private Const MAX_LOG_LINES As Long = 1000
Private Const STATUS_LOG_EVERY_N_TICKS As Long = 4

Private g_TickCount As Long

Private Sub Form_Load()
    AppendLog "I/O App started"
    lblStatusValue.Caption = "Stopped"
End Sub

Private Sub cmdStart_Click()
    AppendLog "Initializing Profibus cards..."

    Dim Status As Integer
    g_CardA_OK = IO_Init(WINCARD_A, Status)
    If g_CardA_OK Then
        AppendLog "Card A initialized OK"
    Else
        AppendLog "Card A failed to initialize, status=" & Status
        ' Same MsgBox the original simulator showed when a card was missing.
        MsgBox "IO Card A failed to initialise", vbCritical, "Error"
    End If

    g_CardB_OK = IO_Init(WINCARD_B, Status)
    If g_CardB_OK Then
        AppendLog "Card B initialized OK"
    Else
        AppendLog "Card B failed to initialize, status=" & Status
        MsgBox "IO Card B failed to initialise", vbCritical, "Error"
    End If

    If g_CardA_OK Or g_CardB_OK Then
        AppendLog "Starting I/O loop"
        lblStatusValue.Caption = "Running"
    Else
        AppendLog "No Profibus cards available; HTTP-only mode (zeros for ain/din, no Profibus writes)"
        lblStatusValue.Caption = "Running (no cards)"
    End If

    Timer1.Enabled = True
End Sub

Private Sub cmdStop_Click()
    AppendLog "Stop clicked"
    Timer1.Enabled = False
    lblStatusValue.Caption = "Stopped"
End Sub

Private Sub Timer1_Timer()
    g_TickCount = g_TickCount + 1

    ' --- Read inputs from Profibus into the arrays ---
    If g_CardA_OK Then
        If Not Get_A_Input(WINCARD_A, A_INPUT_START_ADDR, A_INPUT_END_ADDR) Then
            AppendLog "Card A AI read failed"
        End If
        If Not Get_D_input(WINCARD_A, D_INPUT_START_ADDR, D_INPUT_END_ADDR) Then
            AppendLog "Card A DI read failed"
        End If
    End If
    If g_CardB_OK Then
        If Not Get_A_Input(WINCARD_B, A_INPUT_START_ADDR_B, A_INPUT_END_ADDR_B) Then
            AppendLog "Card B AI read failed"
        End If
        If Not Get_D_input(WINCARD_B, D_INPUT_START_ADDR_B, D_INPUT_END_ADDR_B) Then
            AppendLog "Card B DI read failed"
        End If
    End If

    ' --- Push inputs to the Python server ---
    Send_Inputs_To_Python

    ' --- Pull outputs from the Python server (async; previous tick's response
    '     is already in A_OUTPUT/D_OUTPUT, so we use that this tick) ---
    Get_Outputs_From_Python

    ' --- Write outputs to Profibus ---
    If g_CardA_OK Then
        If Not Set_A_Output(WINCARD_A, A_OUTPUT_START_ADDR, A_OUTPUT_END_ADDR) Then
            AppendLog "Card A AO write failed"
        End If
        If Not Set_D_Output(WINCARD_A, D_OUTPUT_START_ADDR, D_OUTPUT_END_ADDR) Then
            AppendLog "Card A DO write failed"
        End If
    End If
    If g_CardB_OK Then
        If Not Set_A_Output(WINCARD_B, A_OUTPUT_START_ADDR_B, A_OUTPUT_END_ADDR_B) Then
            AppendLog "Card B AO write failed"
        End If
        If Not Set_D_Output(WINCARD_B, D_OUTPUT_START_ADDR_B, D_OUTPUT_END_ADDR_B) Then
            AppendLog "Card B DO write failed"
        End If
    End If

    ' Periodic status line (don't log every tick - too noisy at 500 ms).
    If g_TickCount Mod STATUS_LOG_EVERY_N_TICKS = 0 Then
        AppendLog "tick  ain[0]=" & A_INPUT(0) & " aout[0]=" & A_OUTPUT(0) & _
                  "  din[0]=" & D_INPUT(0) & " dout[0]=" & D_OUTPUT(0)
    End If
End Sub

Public Sub AppendLog(msg As String)
    ' Avoid the name "Log" because VB6's built-in Log() is the natural logarithm.
    Dim line As String
    line = Format$(Now, "hh:nn:ss") & "  " & msg
    lstLog.AddItem line
    ' Scroll to bottom so the newest line is visible.
    If lstLog.ListCount > 0 Then
        lstLog.TopIndex = lstLog.ListCount - 1
    End If
    ' Cap memory: drop the oldest line once we exceed MAX_LOG_LINES.
    If lstLog.ListCount > MAX_LOG_LINES Then
        lstLog.RemoveItem 0
    End If
End Sub
