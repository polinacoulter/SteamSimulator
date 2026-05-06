VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "havenstretch.ocx"
Object = "{918E6E43-F23A-11D0-901E-0020AF7543C2}#5.0#0"; "fximg50g.ocx"
Begin VB.Form Boiler_form 
   Caption         =   "Boiler"
   ClientHeight    =   11685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   11685
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   11175
      Left            =   60
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   19711
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   360
         Top             =   9060
      End
      Begin FXIMG50GLib.FXImage FXImage4 
         Height          =   6060
         Left            =   7380
         TabIndex        =   1
         Top             =   1440
         Width           =   3690
         _Version        =   327680
         _ExtentX        =   6509
         _ExtentY        =   10689
         _StockProps     =   65
         BackColor       =   12632256
         Persistence     =   -1  'True
         _StdProps       =   -3043161
         ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
         ErrCode         =   1056336839
         ErrInfo         =   -2041638552
         Picture         =   "Boiler_form.frx":0000
      End
      Begin FXIMG50GLib.FXImage FXImage2 
         Height          =   3600
         Left            =   6360
         TabIndex        =   0
         Top             =   6600
         Width           =   4800
         _Version        =   327680
         _ExtentX        =   8467
         _ExtentY        =   6350
         _StockProps     =   65
         BackColor       =   12632256
         Persistence     =   -1  'True
         _StdProps       =   8473162
         ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
         ErrCode         =   1056336839
         ErrInfo         =   -2041638552
         Picture         =   "Boiler_form.frx":18BB2
         DrawStyle       =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Flame Inspection"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "Move to the HANDLE to OPEN/CLOSE"
         Top             =   5760
         Width           =   975
      End
      Begin VB.Label Cover_handle 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C0C0C0&
         Height          =   75
         Index           =   1
         Left            =   960
         TabIndex        =   4
         ToolTipText     =   "CLOSE Flame Inspection Cover"
         Top             =   5220
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Cover_handle 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C0C0C0&
         Height          =   75
         Index           =   0
         Left            =   1380
         TabIndex        =   3
         ToolTipText     =   "OPEN Flame Inspection Cover"
         Top             =   5220
         Width           =   195
      End
      Begin VB.Shape grey_cover 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   1
         Left            =   840
         Shape           =   2  'Oval
         Top             =   4860
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape grey_cover 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   0
         Left            =   840
         Shape           =   2  'Oval
         Top             =   4860
         Width           =   795
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   540
         X2              =   4560
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   540
         X2              =   4560
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   4560
         X2              =   4560
         Y1              =   660
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         X1              =   540
         X2              =   540
         Y1              =   660
         Y2              =   6840
      End
      Begin VB.Image Flames 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Index           =   0
         Left            =   240
         Picture         =   "Boiler_form.frx":2BC04
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   675
      End
      Begin VB.Image Flames 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Index           =   3
         Left            =   2760
         Picture         =   "Boiler_form.frx":3EC46
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   675
      End
      Begin VB.Image Flames 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Index           =   2
         Left            =   1920
         Picture         =   "Boiler_form.frx":51C88
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   675
      End
      Begin VB.Image Flames 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Index           =   1
         Left            =   1080
         Picture         =   "Boiler_form.frx":64CCA
         Stretch         =   -1  'True
         Top             =   8160
         Width           =   675
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   1455
         Index           =   0
         Left            =   7380
         Shape           =   2  'Oval
         Top             =   5820
         Width           =   1395
      End
      Begin VB.Label Label1 
         Height          =   795
         Left            =   5340
         TabIndex        =   2
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   2400
         Left            =   5160
         Picture         =   "Boiler_form.frx":77D0C
         Stretch         =   -1  'True
         Top             =   1500
         Width           =   2340
      End
      Begin VB.Image Boiler_flame 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   960
         Picture         =   "Boiler_form.frx":8AD4E
         Stretch         =   -1  'True
         ToolTipText     =   "WARNING FIRE VERY HOT"
         Top             =   4980
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   720
         Shape           =   1  'Square
         Top             =   4740
         Width           =   1035
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6180
         Left            =   540
         Picture         =   "Boiler_form.frx":9DD90
         Stretch         =   -1  'True
         Top             =   660
         Width           =   4050
      End
   End
End
Attribute VB_Name = "Boiler_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cover_handle_Click(Index As Integer)

If Cover_handle(0).Visible Then

    grey_cover(0).Visible = False
    grey_cover(1).Visible = True
    
    Cover_handle(0).Visible = False
    Cover_handle(1).Visible = True
    
    Boiler_flame.Visible = True
    
Else
    grey_cover(0).Visible = True
    grey_cover(1).Visible = False
    Cover_handle(0).Visible = True
    Cover_handle(1).Visible = False
    
    Boiler_flame.Visible = False
    
End If

End Sub

Private Sub Timer1_Timer()

Flame_count = Flame_count + 1

If Flame_count = 4 Then
    Flame_count = 0
End If

If grey_cover(1).Visible Then
    Boiler_flame.Picture = Flames(Flame_count).Picture
End If

End Sub
