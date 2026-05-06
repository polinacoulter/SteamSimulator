VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SnapProgressForm 
   Caption         =   "Snapshot"
   ClientHeight    =   2865
   ClientLeft      =   4680
   ClientTop       =   2745
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   4440
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2220
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   585
         Left            =   540
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "SnapProgressForm.frx":0000
         Top             =   900
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   3015
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "SnapProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()
    SnapProgressForm.CmdOK.Visible = False
    Unload Me
End Sub
