VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Gland_Steam 
   Caption         =   "Gland Steam System"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   13215
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7785
      Left            =   60
      Top             =   60
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   13732
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
         TabIndex        =   26
         Text            =   "xx psi"
         Top             =   3900
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   900
         Top             =   1680
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Main Turbines"
         Height          =   435
         Left            =   3630
         TabIndex        =   21
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Condensate"
         Height          =   435
         Left            =   10860
         TabIndex        =   20
         Top             =   5340
         Width           =   1125
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   2
         Left            =   8820
         TabIndex        =   5
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
         Left            =   9360
         TabIndex        =   6
         Top             =   3540
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
         TabIndex        =   7
         Top             =   3540
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
         Left            =   9360
         TabIndex        =   8
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
         Height          =   435
         Index           =   106
         Left            =   1140
         TabIndex        =   11
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Gland Exhaust Condenser"
         Height          =   675
         Left            =   7200
         TabIndex        =   27
         Top             =   5040
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   9060
         X2              =   9060
         Y1              =   3180
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderStyle     =   2  'Dash
         X1              =   9240
         X2              =   9240
         Y1              =   4560
         Y2              =   5820
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "146 + 50   kg / hr"
         Height          =   285
         Left            =   6510
         TabIndex        =   25
         Top             =   3390
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Steam + Air Flow"
         Height          =   435
         Left            =   6510
         TabIndex        =   24
         Top             =   2820
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "436 kg / hr"
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   3390
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Steam Flow"
         Height          =   435
         Left            =   330
         TabIndex        =   22
         Top             =   2850
         Width           =   555
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
         Left            =   1710
         TabIndex        =   2
         Top             =   3240
         Width           =   795
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   19
         X1              =   9420
         X2              =   8880
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label Label6 
         Caption         =   "150 psi."
         Height          =   255
         Left            =   8940
         TabIndex        =   19
         Top             =   1740
         Width           =   615
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   7680
         Picture         =   "Gland_Steam.frx":0000
         Stretch         =   -1  'True
         Top             =   4740
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Air Ejector"
         Height          =   285
         Left            =   10140
         TabIndex        =   18
         Top             =   4530
         Width           =   735
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   450
         Picture         =   "Gland_Steam.frx":030A
         Stretch         =   -1  'True
         Top             =   3660
         Width           =   255
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   315
         Left            =   8640
         OleObjectBlob   =   "Gland_Steam.frx":0614
         SizeMode        =   1  'Stretch
         TabIndex        =   17
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
         OleObjectBlob   =   "Gland_Steam.frx":222C
         SizeMode        =   1  'Stretch
         TabIndex        =   16
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
         OleObjectBlob   =   "Gland_Steam.frx":3E44
         SizeMode        =   1  'Stretch
         TabIndex        =   15
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
         OleObjectBlob   =   "Gland_Steam.frx":5A5C
         SizeMode        =   1  'Stretch
         TabIndex        =   14
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "From 150 # System"
         Height          =   495
         Left            =   210
         TabIndex        =   13
         Top             =   3900
         Width           =   855
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   8220
         Picture         =   "Gland_Steam.frx":7674
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   255
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   2040
         X2              =   10620
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "To steam side of main condenser"
         Height          =   195
         Left            =   5700
         TabIndex        =   10
         Top             =   930
         Width           =   2415
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   7
         Left            =   1920
         Picture         =   "Gland_Steam.frx":797E
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   255
      End
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   2040
         X2              =   2040
         Y1              =   3120
         Y2              =   1200
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   5
         Left            =   8220
         Picture         =   "Gland_Steam.frx":7C88
         Stretch         =   -1  'True
         Top             =   6060
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   6
         Left            =   8700
         Picture         =   "Gland_Steam.frx":7F92
         Stretch         =   -1  'True
         Top             =   6060
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   9060
         Picture         =   "Gland_Steam.frx":829C
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   10080
         Picture         =   "Gland_Steam.frx":85A6
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   10320
         Picture         =   "Gland_Steam.frx":88B0
         Stretch         =   -1  'True
         Top             =   5040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   6540
         Picture         =   "Gland_Steam.frx":8BBA
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9615
         Top             =   4920
         Width           =   285
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
         Index           =   10
         X1              =   9060
         X2              =   9060
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
         Left            =   9075
         Top             =   4920
         Width           =   285
      End
      Begin VB.Shape Shape26 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8835
         Top             =   4860
         Width           =   225
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
         X2              =   8820
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   8820
         X2              =   8820
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   8820
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   9600
         X2              =   9600
         Y1              =   4860
         Y2              =   5520
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
         Index           =   23
         X1              =   9900
         X2              =   11160
         Y1              =   5175
         Y2              =   5175
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
      Begin VB.Image Image12 
         Height          =   255
         Left            =   10320
         Picture         =   "Gland_Steam.frx":8EC4
         Stretch         =   -1  'True
         Top             =   5700
         Width           =   255
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
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   8580
         Top             =   4560
         Width           =   1515
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
         Index           =   11
         X1              =   7680
         X2              =   8220
         Y1              =   4860
         Y2              =   4860
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
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   9
         X1              =   9180
         X2              =   9180
         Y1              =   2460
         Y2              =   2040
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   9840
         X2              =   9060
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   8760
         X2              =   9420
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label18 
         Caption         =   "Gland Seal drains"
         Height          =   675
         Left            =   6810
         TabIndex        =   9
         Top             =   3870
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
         Index           =   15
         X1              =   8880
         X2              =   8880
         Y1              =   3480
         Y2              =   2460
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   9450
         X2              =   9450
         Y1              =   3960
         Y2              =   2460
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   9840
         X2              =   9840
         Y1              =   3960
         Y2              =   3180
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   8220
         Picture         =   "Gland_Steam.frx":91CE
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   9060
         X2              =   9060
         Y1              =   4140
         Y2              =   3600
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
         Index           =   56
         X1              =   9000
         X2              =   9000
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Shape Shape33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   8760
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   255
      End
      Begin VB.Shape Shape34 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8775
         Top             =   4860
         Width           =   225
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   53
         X1              =   9300
         X2              =   9300
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   54
         X1              =   9540
         X2              =   9540
         Y1              =   4860
         Y2              =   5520
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   9300
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   9315
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
         Height          =   615
         Left            =   8235
         Top             =   4860
         Width           =   225
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   4680
         Width           =   315
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   435
         Left            =   9540
         Shape           =   3  'Circle
         Top             =   4680
         Width           =   315
      End
      Begin VB.Shape Shape18 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9015
         Top             =   4920
         Width           =   285
      End
      Begin VB.Shape Shape38 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   9555
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
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   3480
         Picture         =   "Gland_Steam.frx":94D8
         Stretch         =   -1  'True
         Top             =   2070
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   3300
         Picture         =   "Gland_Steam.frx":97E2
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   3240
         Picture         =   "Gland_Steam.frx":9AEC
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   9300
         Picture         =   "Gland_Steam.frx":9DF6
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   4920
         Picture         =   "Gland_Steam.frx":A100
         Stretch         =   -1  'True
         Top             =   4110
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   4920
         Picture         =   "Gland_Steam.frx":A40A
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   6180
         Picture         =   "Gland_Steam.frx":A714
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   5640
         Picture         =   "Gland_Steam.frx":AA1E
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
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
      Begin VB.Image Image2 
         Height          =   255
         Left            =   6180
         Picture         =   "Gland_Steam.frx":AD28
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   255
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   5640
         Picture         =   "Gland_Steam.frx":B032
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   255
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
      Begin VB.Line HHP_Steam_line 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   3600
         X2              =   4920
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Image Arrow_up 
         Height          =   255
         Left            =   3360
         Picture         =   "Gland_Steam.frx":B33C
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   3030
         Picture         =   "Gland_Steam.frx":B646
         Stretch         =   -1  'True
         Top             =   3660
         Width           =   255
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
         TabIndex        =   1
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
         TabIndex        =   0
         Top             =   4500
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   4800
         Picture         =   "Gland_Steam.frx":B950
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   255
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
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   4740
         Top             =   4740
         Width           =   555
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
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   3390
         Top             =   2760
         Width           =   315
      End
      Begin VB.Image Image15 
         Height          =   255
         Left            =   3360
         Picture         =   "Gland_Steam.frx":BC5A
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   255
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
         Index           =   1
         X1              =   9660
         X2              =   8580
         Y1              =   5100
         Y2              =   5100
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
      Begin VB.Line HP_steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   3
         X1              =   5040
         X2              =   5040
         Y1              =   2940
         Y2              =   4740
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
         OleObjectBlob   =   "Gland_Steam.frx":BF64
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\My Documents\visiotbne.vsd"
         TabIndex        =   3
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
         OleObjectBlob   =   "Gland_Steam.frx":1377C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   4
         Top             =   3900
         Width           =   1215
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
         Caption         =   "Equalising pipe"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1650
         Top             =   3180
         Width           =   915
      End
      Begin VB.Line HP_steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   2
         X1              =   5040
         X2              =   480
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   90
         Top             =   3300
         Width           =   1005
      End
      Begin VB.Shape Shape21 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   6450
         Top             =   3330
         Width           =   1425
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   5
      X1              =   8280
      X2              =   8280
      Y1              =   5280
      Y2              =   5940
   End
End
Attribute VB_Name = "Gland_Steam"
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

On Error Resume Next
Call JYAScreenInit(Me, idGland_Steam)


 
' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idGland_Steam)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub Timer1_Timer()


' 10/8/00
'       PC_REALS(726) = P_steam_gland_seal

If SI_units Then

    P_steam_gland_seal.text = Format(LSS_reals(726) * Press_mult, "##0.0") & Press_chars     ' Bar
    
Else
    P_steam_gland_seal.text = Format(LSS_reals(726) * Press_mult, "###0") & Press_chars     ' psi
    
End If


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
