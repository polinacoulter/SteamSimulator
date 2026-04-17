VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Begin VB.Form frm_LogView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exercise Log"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10830
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8235
      Left            =   0
      Top             =   60
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   14526
      Begin VB.CommandButton PrintLog 
         Caption         =   "Print"
         Height          =   495
         Left            =   9480
         TabIndex        =   1
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   9480
         TabIndex        =   0
         Top             =   600
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9720
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   7935
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   13996
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frm_LogView.frx":0000
      End
   End
End
Attribute VB_Name = "frm_LogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
    LogViewOn = False
    Unload Me
    
End Sub

Private Sub Form_Load()

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub PrintLog_Click()

CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If RichTextBox1.SelLength = 0 Then
        CommonDialog1.Flags = CommonDialog1.Flags + cdlPDAllPages
    Else
        CommonDialog1.Flags = CommonDialog1.Flags + cdlPDSelection
    End If
    CommonDialog1.ShowPrinter
    Printer.Print ""
    RichTextBox1.SelPrint CommonDialog1.hdc


End Sub
