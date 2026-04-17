VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Gland_Steam_Full_away 
   Caption         =   "Gland_Steam"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11100
   ScaleWidth      =   13020
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7155
      Left            =   60
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12621
      Begin VB.TextBox F_seal_IN_REG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2700
         TabIndex        =   30
         Top             =   3420
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Seal_REG_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   29
         Top             =   4080
         Width           =   675
      End
      Begin VB.TextBox F_seal_out 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6540
         TabIndex        =   28
         Top             =   3420
         Width           =   915
      End
      Begin VB.TextBox F_seal_IN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   915
      End
      Begin VB.TextBox P_steam_gland_seal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   20
         Text            =   "xx psi"
         Top             =   3840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Main Turbines"
         Height          =   435
         Left            =   3630
         TabIndex        =   1
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Condensate"
         Height          =   435
         Left            =   10620
         TabIndex        =   0
         Top             =   5280
         Width           =   1125
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   720
         Top             =   5340
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   106
         Left            =   1140
         TabIndex        =   2
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   2
         Left            =   8880
         TabIndex        =   21
         Top             =   2640
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   3
         Left            =   9420
         TabIndex        =   22
         Top             =   3600
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   8
         Left            =   9780
         TabIndex        =   23
         Top             =   3600
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   9
         Left            =   9390
         TabIndex        =   24
         Top             =   2640
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   4920
         Picture         =   "Gland_Steam_Full_away.frx":0000
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Image Arrow_up 
         Height          =   255
         Left            =   3360
         Picture         =   "Gland_Steam_Full_away.frx":030A
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   9060
         X2              =   9060
         Y1              =   3360
         Y2              =   3240
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Gland Exhaust Condenser"
         Height          =   675
         Left            =   7200
         TabIndex        =   26
         Top             =   5040
         Width           =   795
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   9060
         Picture         =   "Gland_Steam_Full_away.frx":0614
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   9060
         X2              =   9060
         Y1              =   3600
         Y2              =   4200
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   9480
         X2              =   9480
         Y1              =   3600
         Y2              =   3480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   9870
         X2              =   9870
         Y1              =   3960
         Y2              =   3240
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   9480
         X2              =   9480
         Y1              =   3240
         Y2              =   2460
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   15
         X1              =   8940
         X2              =   8940
         Y1              =   3480
         Y2              =   2460
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   14
         X1              =   8760
         X2              =   8760
         Y1              =   4140
         Y2              =   3480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   8790
         X2              =   9480
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   9870
         X2              =   9060
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   9180
         X2              =   9180
         Y1              =   2460
         Y2              =   2040
      End
      Begin VB.Label Label6 
         Caption         =   "150 psi."
         Height          =   255
         Left            =   8940
         TabIndex        =   25
         Top             =   1740
         Width           =   615
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   19
         X1              =   9450
         X2              =   8940
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   9060
         X2              =   9060
         Y1              =   4920
         Y2              =   5400
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   8790
         Top             =   4890
         Width           =   255
      End
      Begin VB.Shape Shape16 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   9615
         Top             =   4920
         Width           =   280
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   7
         Left            =   1920
         Picture         =   "Gland_Steam_Full_away.frx":091E
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   3240
         Picture         =   "Gland_Steam_Full_away.frx":0C28
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   3360
         Picture         =   "Gland_Steam_Full_away.frx":0F32
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   10320
         Picture         =   "Gland_Steam_Full_away.frx":123C
         Stretch         =   -1  'True
         Top             =   5040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   10320
         Picture         =   "Gland_Steam_Full_away.frx":1546
         Stretch         =   -1  'True
         Top             =   5700
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   5
         Left            =   8220
         Picture         =   "Gland_Steam_Full_away.frx":1850
         Stretch         =   -1  'True
         Top             =   6060
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   6
         Left            =   8700
         Picture         =   "Gland_Steam_Full_away.frx":1B5A
         Stretch         =   -1  'True
         Top             =   6060
         Width           =   255
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   7680
         Picture         =   "Gland_Steam_Full_away.frx":1E64
         Stretch         =   -1  'True
         Top             =   4740
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   8220
         Picture         =   "Gland_Steam_Full_away.frx":216E
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   8220
         Picture         =   "Gland_Steam_Full_away.frx":2478
         Stretch         =   -1  'True
         Top             =   1140
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   6540
         Picture         =   "Gland_Steam_Full_away.frx":2782
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   6180
         Picture         =   "Gland_Steam_Full_away.frx":2A8C
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   6180
         Picture         =   "Gland_Steam_Full_away.frx":2D96
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   5640
         Picture         =   "Gland_Steam_Full_away.frx":30A0
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   5640
         Picture         =   "Gland_Steam_Full_away.frx":33AA
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   4920
         Picture         =   "Gland_Steam_Full_away.frx":36B4
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   4800
         Picture         =   "Gland_Steam_Full_away.frx":39BE
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   3300
         Picture         =   "Gland_Steam_Full_away.frx":3CC8
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   3480
         Picture         =   "Gland_Steam_Full_away.frx":3FD2
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   1680
         X2              =   240
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line HP_steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   0
         X1              =   3480
         X2              =   3480
         Y1              =   2940
         Y2              =   4740
      End
      Begin VB.Line HP_steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   2
         X1              =   5040
         X2              =   2550
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line HP_steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   3
         X1              =   5040
         X2              =   5040
         Y1              =   2940
         Y2              =   4740
      End
      Begin VB.Label Label5 
         Caption         =   "Air Ejector"
         Height          =   285
         Left            =   10140
         TabIndex        =   17
         Top             =   4530
         Width           =   735
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   315
         Left            =   8640
         OleObjectBlob   =   "Gland_Steam_Full_away.frx":42DC
         SizeMode        =   1  'Stretch
         TabIndex        =   16
         Top             =   4260
         Width           =   195
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   315
         Left            =   8940
         OleObjectBlob   =   "Gland_Steam_Full_away.frx":5EF4
         SizeMode        =   1  'Stretch
         TabIndex        =   15
         Top             =   4260
         Width           =   195
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   495
         Left            =   9300
         OleObjectBlob   =   "Gland_Steam_Full_away.frx":7B0C
         SizeMode        =   1  'Stretch
         TabIndex        =   14
         Top             =   4080
         Width           =   315
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   495
         Left            =   9720
         OleObjectBlob   =   "Gland_Steam_Full_away.frx":9724
         SizeMode        =   1  'Stretch
         TabIndex        =   13
         Top             =   4080
         Width           =   255
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   2040
         X2              =   10680
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "To steam side of main condenser"
         Height          =   195
         Left            =   5700
         TabIndex        =   12
         Top             =   930
         Width           =   2415
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   2040
         X2              =   2040
         Y1              =   3120
         Y2              =   1260
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   10080
         Picture         =   "Gland_Steam_Full_away.frx":B33C
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   13
         X1              =   8220
         X2              =   8220
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   9360
         X2              =   9360
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Shape Shape22 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   9375
         Top             =   4860
         Width           =   225
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9095
         Top             =   4920
         Width           =   345
      End
      Begin VB.Shape Shape20 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8235
         Top             =   4860
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   7740
         X2              =   8220
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   8520
         X2              =   8760
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Shape Shape25 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   9360
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   255
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   9060
         Shape           =   3  'Circle
         Top             =   4680
         Width           =   315
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   9600
         Shape           =   3  'Circle
         Top             =   4680
         Width           =   315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   47
         X1              =   9900
         X2              =   9900
         Y1              =   5160
         Y2              =   4920
      End
      Begin VB.Shape Loop_Seal_shape 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   10035
         Top             =   5940
         Width           =   210
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   8640
         Top             =   4140
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   10260
         X2              =   10740
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   22
         X1              =   8460
         X2              =   8760
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   14
         X1              =   8340
         X2              =   8340
         Y1              =   5820
         Y2              =   6660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   42
         X1              =   8820
         X2              =   8820
         Y1              =   5820
         Y2              =   6660
      End
      Begin VB.Label Label18 
         Caption         =   "Gland Seal drains"
         Height          =   675
         Left            =   6810
         TabIndex        =   11
         Top             =   3930
         Width           =   555
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   8940
         Top             =   4140
         Width           =   255
      End
      Begin VB.Shape Shape23 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   9720
         Top             =   3960
         Width           =   255
      End
      Begin VB.Shape Shape24 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   9360
         Top             =   3960
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   48
         X1              =   10260
         X2              =   10260
         Y1              =   5820
         Y2              =   6180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   49
         X1              =   10020
         X2              =   10020
         Y1              =   5820
         Y2              =   6180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   55
         X1              =   8760
         X2              =   8760
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   54
         X1              =   9600
         X2              =   9600
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   9375
         Top             =   4860
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   59
         X1              =   8460
         X2              =   8460
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   60
         X1              =   8460
         X2              =   8460
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Shape Shape36 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   8220
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   255
      End
      Begin VB.Shape Shape37 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   8235
         Top             =   4800
         Width           =   225
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   5220
         Width           =   315
      End
      Begin VB.Shape Shape18 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9045
         Top             =   4920
         Width           =   285
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   10200
         X2              =   10200
         Y1              =   4020
         Y2              =   1980
      End
      Begin VB.Line Line20 
         BorderStyle     =   3  'Dot
         X1              =   9420
         X2              =   10200
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   8700
         X2              =   9300
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   9300
         X2              =   9300
         Y1              =   4560
         Y2              =   4200
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   8340
         X2              =   5760
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   5160
         X2              =   5760
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   8
         X1              =   5160
         X2              =   5160
         Y1              =   4920
         Y2              =   5280
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   3360
         X2              =   6300
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   3360
         X2              =   3360
         Y1              =   4920
         Y2              =   6060
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   5160
         X2              =   5760
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   5160
         X2              =   5160
         Y1              =   2460
         Y2              =   2760
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   5760
         X2              =   5760
         Y1              =   2460
         Y2              =   5280
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HP Turbine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3900
         TabIndex        =   10
         Top             =   2580
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LP Turbine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3960
         TabIndex        =   9
         Top             =   4500
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   4740
         Top             =   4740
         Width           =   555
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   4920
         X2              =   4920
         Y1              =   2760
         Y2              =   1920
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   3600
         X2              =   3600
         Y1              =   2760
         Y2              =   1920
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   3420
         X2              =   3420
         Y1              =   1560
         Y2              =   2760
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   8340
         X2              =   8340
         Y1              =   3840
         Y2              =   4560
      End
      Begin VB.Label Label3 
         Caption         =   "Equalising connection"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   1650
         Width           =   1695
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   3600
         X2              =   4920
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   3390
         Top             =   2760
         Width           =   315
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   4740
         Top             =   2760
         Width           =   555
      End
      Begin VB.Shape Shape15 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   3300
         Top             =   4740
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   23
         X1              =   9900
         X2              =   11160
         Y1              =   5160
         Y2              =   5175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   9060
         X2              =   9060
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Steam Flow"
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   2850
         Width           =   555
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 kg / hr"
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gland Steam Regulator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1740
         TabIndex        =   5
         Top             =   3240
         Width           =   795
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Steam Flow"
         Height          =   435
         Left            =   6510
         TabIndex        =   4
         Top             =   2790
         Width           =   1005
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "151 + 50   kg / hr"
         Height          =   285
         Left            =   6600
         TabIndex        =   3
         Top             =   1920
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   6300
         X2              =   6300
         Y1              =   6060
         Y2              =   1590
      End
      Begin VB.Line LP_steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   6300
         X2              =   3420
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   11
         X1              =   7680
         X2              =   8220
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   10020
         Shape           =   3  'Circle
         Top             =   6000
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   6480
         Top             =   2220
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   300
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1680
         Top             =   3150
         Width           =   915
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   3660
         OleObjectBlob   =   "Gland_Steam_Full_away.frx":B646
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\My Documents\visiotbne.vsd"
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   2145
         Left            =   3600
         OleObjectBlob   =   "Gland_Steam_Full_away.frx":12E5E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   19
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   8610
         Top             =   4560
         Width           =   1515
      End
      Begin VB.Shape Shape34 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   9630
         Top             =   4860
         Width           =   255
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1290
         Left            =   8100
         Top             =   4560
         Width           =   495
      End
   End
End
Attribute VB_Name = "Gland_Steam_Full_away"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

' 14/2/00
Condensate.Show

End Sub

Private Sub Command2_Click()

' 14/2/00
Main_Turbines.Show

End Sub

Private Sub Form_Load()

' === 01/02/02 ===
On Error Resume Next
Call JYAScreenInit(Me, idGland_Steam)

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

' === 01/02/02 ===
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idGland_Steam)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Timer1_Timer()


' === 01/02/02 ===
'           PC_REALS(837) = F_seal_out
'           PC_REALS(838) = F_seal_IN_REG
'           PC_REALS(839) = Seal_REG_pos * 100#
'           PC_REALS(840) = F_seal_IN



' 10/8/00
'       PC_REALS(726) = P_steam_gland_seal

If SI_units Then

    P_steam_gland_seal.text = Format(LSS_reals(726) * Press_mult, "##0.00") & Press_chars     ' Bar
    
    F_seal_out.text = Format(LSS_reals(837) * Flow_mult, "###0.000") & Flow_chars
    F_seal_IN_REG.text = Format(LSS_reals(838) * Flow_mult, "###0.000") & Flow_chars
    F_seal_IN.text = Format(LSS_reals(840) * Flow_mult, "###0.000") & Flow_chars
    
Else
    P_steam_gland_seal.text = Format(LSS_reals(726) * Press_mult, "###0.0") & Press_chars     ' psi
    
    F_seal_out.text = Format(LSS_reals(837) * Flow_mult / 3600#, "###0.000") & " lb/s"
    F_seal_IN_REG.text = Format(LSS_reals(838) * Flow_mult / 3600#, "###0.000") & " lb/s"
    F_seal_IN.text = Format(LSS_reals(840) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
End If

Seal_REG_pos.text = Format(LSS_reals(839), "##0") & "%"



' 14/2/00
' update_vvalve_LSS(syxxx,nValve,index1,index2)
   Call update_vvalve_LSS(Me, 106, 106, -1)
   
   
   Call update_vvalve_LSS(Me, 2, 2, -1)
   Call update_vvalve_LSS(Me, 3, 3, -1)
   Call update_vvalve_LSS(Me, 8, 8, -1)
   Call update_vvalve_LSS(Me, 9, 9, -1)



End Sub

Private Sub valve_Click(Index As Integer)

' 14/2/00
Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
Call valve_pcdi_LSS(Index)


End Sub
