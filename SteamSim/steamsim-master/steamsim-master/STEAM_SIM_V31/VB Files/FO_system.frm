VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form FO_system 
   Caption         =   "Fuel Oil System"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11475
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6435
      Left            =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11351
      Begin VB.HScrollBar Air_rgister_scroll 
         Height          =   135
         Index           =   4
         Left            =   9000
         Max             =   100
         TabIndex        =   60
         Top             =   4200
         Width           =   975
      End
      Begin VB.HScrollBar Air_rgister_scroll 
         Height          =   135
         Index           =   3
         Left            =   9000
         Max             =   100
         TabIndex        =   59
         Top             =   3480
         Width           =   975
      End
      Begin VB.HScrollBar Air_rgister_scroll 
         Height          =   135
         Index           =   2
         Left            =   9000
         Max             =   100
         TabIndex        =   58
         Top             =   1440
         Width           =   975
      End
      Begin VB.HScrollBar Air_rgister_scroll 
         Height          =   135
         Index           =   1
         Left            =   9000
         Max             =   100
         TabIndex        =   57
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Left            =   2580
         TabIndex        =   24
         Text            =   "83 F"
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Text2 
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
         Left            =   3300
         TabIndex        =   23
         Text            =   "72 psi"
         Top             =   4320
         Width           =   675
      End
      Begin VB.TextBox Text8 
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
         Left            =   3420
         TabIndex        =   15
         Text            =   "70 psi"
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Damper"
         Height          =   555
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         Left            =   900
         TabIndex        =   13
         Text            =   " 10.6 kg/s"
         Top             =   3060
         Width           =   795
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         TabIndex        =   12
         Text            =   "79 F"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Fuel Oil Service"
         Height          =   555
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   2220
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   2220
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   2
         Left            =   5460
         TabIndex        =   5
         Top             =   480
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   3
         Left            =   5460
         TabIndex        =   6
         Top             =   1260
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   4
         Left            =   4800
         TabIndex        =   7
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   5
         Left            =   4800
         TabIndex        =   8
         Top             =   1260
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   6
         Left            =   3360
         TabIndex        =   9
         Top             =   840
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   7
         Left            =   2640
         TabIndex        =   10
         Top             =   840
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   8
         Left            =   3780
         TabIndex        =   22
         Top             =   1860
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   9
         Left            =   5460
         TabIndex        =   25
         Top             =   3360
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   10
         Left            =   5460
         TabIndex        =   26
         Top             =   4140
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   11
         Left            =   4800
         TabIndex        =   27
         Top             =   3360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   12
         Left            =   4800
         TabIndex        =   28
         Top             =   4140
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   13
         Left            =   3360
         TabIndex        =   29
         Top             =   3720
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   14
         Left            =   2640
         TabIndex        =   30
         Top             =   3720
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   15
         Left            =   3780
         TabIndex        =   31
         Top             =   2820
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   16
         Left            =   4140
         TabIndex        =   39
         Top             =   840
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   17
         Left            =   4560
         TabIndex        =   40
         Top             =   1860
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   18
         Left            =   4560
         TabIndex        =   41
         Top             =   2820
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   19
         Left            =   4140
         TabIndex        =   42
         Top             =   3720
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
         ValveType       =   1
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   420
         Width           =   135
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   1260
         Width           =   135
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   3300
         Width           =   135
      End
      Begin VB.Shape Burner_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6480
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   135
      End
      Begin VB.Shape Ign_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   4
         Left            =   7020
         Shape           =   2  'Oval
         Top             =   4620
         Width           =   375
      End
      Begin VB.Shape Ign_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   7020
         Shape           =   2  'Oval
         Top             =   3900
         Width           =   375
      End
      Begin VB.Shape Ign_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   7020
         Shape           =   2  'Oval
         Top             =   1800
         Width           =   375
      End
      Begin VB.Shape Ign_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   7020
         Shape           =   2  'Oval
         Top             =   1020
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   4
         Left            =   5160
         Shape           =   2  'Oval
         Top             =   4740
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   5100
         Shape           =   2  'Oval
         Top             =   3900
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   5100
         Shape           =   2  'Oval
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         Caption         =   "Air Registers Stbd"
         Height          =   255
         Index           =   5
         Left            =   8940
         TabIndex        =   66
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         Caption         =   "Air Registers Port"
         Height          =   255
         Index           =   4
         Left            =   8940
         TabIndex        =   65
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         Caption         =   "# 2"
         Height          =   255
         Index           =   3
         Left            =   10020
         TabIndex        =   64
         Top             =   4140
         Width           =   315
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         Caption         =   "# 1"
         Height          =   255
         Index           =   2
         Left            =   10020
         TabIndex        =   63
         Top             =   3480
         Width           =   315
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         Caption         =   "# 2"
         Height          =   255
         Index           =   1
         Left            =   10080
         TabIndex        =   62
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label Reg_label 
         Alignment       =   2  'Center
         Caption         =   "# 1"
         Height          =   255
         Index           =   0
         Left            =   10080
         TabIndex        =   61
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label30 
         Caption         =   "Igniter extended and enegrised"
         Height          =   255
         Left            =   900
         TabIndex        =   56
         Top             =   5820
         Width           =   2655
      End
      Begin VB.Label Label29 
         Caption         =   "Igniter retracted and de-enegrised"
         Height          =   255
         Left            =   900
         TabIndex        =   55
         Top             =   5460
         Width           =   2655
      End
      Begin VB.Shape Shape22 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   300
         Shape           =   2  'Oval
         Top             =   5940
         Width           =   375
      End
      Begin VB.Shape Shape21 
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   300
         Shape           =   2  'Oval
         Top             =   5520
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   5100
         Shape           =   2  'Oval
         Top             =   1080
         Width           =   375
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   23
         X1              =   6000
         X2              =   6180
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   22
         X1              =   6180
         X2              =   6300
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   19
         X1              =   6060
         X2              =   6180
         Y1              =   4740
         Y2              =   4560
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   18
         X1              =   6180
         X2              =   6300
         Y1              =   4740
         Y2              =   4560
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   14
         X1              =   6120
         X2              =   6120
         Y1              =   4560
         Y2              =   4740
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   15
         X1              =   6240
         X2              =   6240
         Y1              =   4560
         Y2              =   4740
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   21
         X1              =   6000
         X2              =   6180
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   20
         X1              =   6180
         X2              =   6300
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   17
         X1              =   6060
         X2              =   6180
         Y1              =   4140
         Y2              =   4260
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   16
         X1              =   6180
         X2              =   6300
         Y1              =   4140
         Y2              =   4260
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   13
         X1              =   6240
         X2              =   6240
         Y1              =   4080
         Y2              =   4260
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   12
         X1              =   6120
         X2              =   6120
         Y1              =   4080
         Y2              =   4260
      End
      Begin VB.Line Port_damper_line 
         Index           =   15
         X1              =   6240
         X2              =   6240
         Y1              =   1680
         Y2              =   1860
      End
      Begin VB.Line Port_damper_line 
         Index           =   14
         X1              =   6120
         X2              =   6120
         Y1              =   1680
         Y2              =   1860
      End
      Begin VB.Line Port_damper_line 
         Index           =   19
         X1              =   6060
         X2              =   6180
         Y1              =   1860
         Y2              =   1680
      End
      Begin VB.Line Port_damper_line 
         Index           =   18
         X1              =   6180
         X2              =   6300
         Y1              =   1860
         Y2              =   1680
      End
      Begin VB.Line Port_damper_line 
         Index           =   22
         X1              =   6180
         X2              =   6300
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Port_damper_line 
         Index           =   23
         X1              =   6000
         X2              =   6180
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Port_damper_line 
         Index           =   21
         X1              =   6000
         X2              =   6180
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Port_damper_line 
         Index           =   20
         X1              =   6120
         X2              =   6300
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Port_damper_line 
         Index           =   13
         X1              =   6240
         X2              =   6240
         Y1              =   1200
         Y2              =   1380
      End
      Begin VB.Line Port_damper_line 
         Index           =   12
         X1              =   6120
         X2              =   6120
         Y1              =   1200
         Y2              =   1380
      End
      Begin VB.Line Port_damper_line 
         Index           =   17
         X1              =   6180
         X2              =   6300
         Y1              =   1260
         Y2              =   1380
      End
      Begin VB.Line Port_damper_line 
         Index           =   16
         X1              =   6060
         X2              =   6180
         Y1              =   1260
         Y2              =   1380
      End
      Begin VB.Label Label28 
         Height          =   135
         Left            =   6340
         TabIndex        =   54
         Top             =   4560
         Width           =   195
      End
      Begin VB.Label Label27 
         Height          =   135
         Left            =   6340
         TabIndex        =   53
         Top             =   4200
         Width           =   195
      End
      Begin VB.Label Label26 
         Height          =   135
         Left            =   6340
         TabIndex        =   52
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label25 
         Height          =   135
         Left            =   6340
         TabIndex        =   51
         Top             =   1680
         Width           =   195
      End
      Begin VB.Label Label24 
         Height          =   135
         Left            =   6340
         TabIndex        =   50
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label Label23 
         Height          =   135
         Left            =   6340
         TabIndex        =   49
         Top             =   540
         Width           =   195
      End
      Begin VB.Label Label22 
         Height          =   135
         Left            =   6340
         TabIndex        =   48
         Top             =   3420
         Width           =   195
      End
      Begin VB.Label Label21 
         Height          =   135
         Left            =   6340
         TabIndex        =   47
         Top             =   3780
         Width           =   195
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   7
         X1              =   4680
         X2              =   4440
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   6
         X1              =   4680
         X2              =   4680
         Y1              =   4440
         Y2              =   3120
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   5
         X1              =   5940
         X2              =   4680
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   4
         X1              =   5940
         X2              =   4680
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   3
         X1              =   4680
         X2              =   4440
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   2
         X1              =   4680
         X2              =   4680
         Y1              =   1920
         Y2              =   780
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   1
         X1              =   5940
         X2              =   4680
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label20 
         Caption         =   "FO re-circ"
         Height          =   255
         Left            =   2520
         TabIndex        =   46
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "From DO Settling Tank"
         Height          =   255
         Left            =   6420
         TabIndex        =   45
         Top             =   2460
         Width           =   2235
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   6120
         Picture         =   "FO_system.frx":0000
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   4200
         TabIndex        =   44
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   600
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   4680
         X2              =   4680
         Y1              =   1920
         Y2              =   3060
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   5
         X1              =   6240
         X2              =   4680
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Image Arrow_left 
         Height          =   255
         Left            =   3300
         Picture         =   "FO_system.frx":030A
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   3600
         Width           =   795
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   3600
         Y2              =   3480
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   3720
         Y2              =   3840
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   5940
         X2              =   8940
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   4380
         Width           =   795
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   4380
         Y2              =   4260
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   4500
         Y2              =   4620
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   1
         X1              =   6120
         X2              =   6120
         Y1              =   3300
         Y2              =   3480
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   0
         X1              =   6240
         X2              =   6240
         Y1              =   3300
         Y2              =   3480
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   2
         X1              =   6240
         X2              =   6240
         Y1              =   3780
         Y2              =   3960
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   3
         X1              =   6120
         X2              =   6120
         Y1              =   3780
         Y2              =   3960
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   6240
         X2              =   6420
         Y1              =   5100
         Y2              =   4800
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   4
         X1              =   6180
         X2              =   6300
         Y1              =   3360
         Y2              =   3480
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   5
         X1              =   6060
         X2              =   6180
         Y1              =   3360
         Y2              =   3480
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   6
         X1              =   6180
         X2              =   6300
         Y1              =   3960
         Y2              =   3780
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   7
         X1              =   6060
         X2              =   6180
         Y1              =   3960
         Y2              =   3780
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   8700
         X2              =   8880
         Y1              =   4800
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   9
         X1              =   4440
         X2              =   2220
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Label Label16 
         Caption         =   "Stbd Boiler Furnace"
         Height          =   255
         Left            =   6960
         TabIndex        =   38
         Top             =   2880
         Width           =   2235
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   8
         X1              =   6180
         X2              =   6300
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   9
         X1              =   6000
         X2              =   6180
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   10
         X1              =   6180
         X2              =   6300
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Stbd_damper_line 
         Index           =   11
         X1              =   6000
         X2              =   6180
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line13 
         X1              =   3120
         X2              =   3120
         Y1              =   4020
         Y2              =   4500
      End
      Begin VB.Line Line12 
         X1              =   3900
         X2              =   3900
         Y1              =   4020
         Y2              =   4440
      End
      Begin VB.Label Label15 
         Caption         =   "Air"
         Height          =   195
         Left            =   6060
         TabIndex        =   37
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label Label14 
         Caption         =   "Burner No 2"
         Height          =   195
         Left            =   6960
         TabIndex        =   36
         Top             =   3540
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "Burner No 1"
         Height          =   195
         Left            =   6960
         TabIndex        =   35
         Top             =   4320
         Width           =   915
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5520
         TabIndex        =   34
         Top             =   3900
         Width           =   255
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   5940
         X2              =   5940
         Y1              =   3060
         Y2              =   5100
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5520
         TabIndex        =   33
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   3420
         TabIndex        =   32
         Top             =   3480
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   8
         X1              =   3360
         X2              =   3900
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   7
         X1              =   3900
         X2              =   3900
         Y1              =   4020
         Y2              =   1140
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   3420
         TabIndex        =   21
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5940
         X2              =   5940
         Y1              =   180
         Y2              =   2220
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   1020
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Burner No 1"
         Height          =   195
         Left            =   6960
         TabIndex        =   18
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Burner No 2"
         Height          =   195
         Left            =   6900
         TabIndex        =   17
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Air"
         Height          =   195
         Left            =   6060
         TabIndex        =   16
         Top             =   120
         Width           =   435
      End
      Begin VB.Line Line10 
         X1              =   3900
         X2              =   3900
         Y1              =   420
         Y2              =   1140
      End
      Begin VB.Line Line9 
         X1              =   3120
         X2              =   3120
         Y1              =   420
         Y2              =   1140
      End
      Begin VB.Line Port_damper_line 
         Index           =   11
         X1              =   6000
         X2              =   6180
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Port_damper_line 
         Index           =   10
         X1              =   6180
         X2              =   6300
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Port_damper_line 
         Index           =   9
         X1              =   6000
         X2              =   6180
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Port_damper_line 
         Index           =   8
         X1              =   6180
         X2              =   6300
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Port Boiler Furnace"
         Height          =   255
         Left            =   6720
         TabIndex        =   11
         Top             =   0
         Width           =   2235
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   6
         X1              =   4380
         X2              =   2220
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Burner_line 
         BorderWidth     =   3
         Index           =   0
         X1              =   5940
         X2              =   4680
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   8700
         X2              =   8880
         Y1              =   1920
         Y2              =   2220
      End
      Begin VB.Line Port_damper_line 
         Index           =   7
         X1              =   6180
         X2              =   6300
         Y1              =   1080
         Y2              =   900
      End
      Begin VB.Line Port_damper_line 
         Index           =   6
         X1              =   6060
         X2              =   6180
         Y1              =   1080
         Y2              =   900
      End
      Begin VB.Line Port_damper_line 
         Index           =   5
         X1              =   6060
         X2              =   6180
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Port_damper_line 
         Index           =   4
         X1              =   6180
         X2              =   6300
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   6240
         X2              =   6420
         Y1              =   2220
         Y2              =   1920
      End
      Begin VB.Line Port_damper_line 
         Index           =   3
         X1              =   6120
         X2              =   6120
         Y1              =   900
         Y2              =   1080
      End
      Begin VB.Line Port_damper_line 
         Index           =   2
         X1              =   6240
         X2              =   6240
         Y1              =   900
         Y2              =   1080
      End
      Begin VB.Line Port_damper_line 
         Index           =   1
         X1              =   6240
         X2              =   6240
         Y1              =   420
         Y2              =   600
      End
      Begin VB.Line Port_damper_line 
         Index           =   0
         X1              =   6120
         X2              =   6120
         Y1              =   420
         Y2              =   600
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   1620
         Y2              =   1740
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   1500
         Y2              =   1380
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   1500
         Width           =   795
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   5940
         X2              =   8940
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   6720
         X2              =   6840
         Y1              =   840
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   6720
         X2              =   6840
         Y1              =   720
         Y2              =   600
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5940
         Top             =   720
         Width           =   795
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   2
         X1              =   2220
         X2              =   2220
         Y1              =   1140
         Y2              =   4020
      End
      Begin VB.Label Label2 
         Caption         =   "Flowmeter"
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   2400
         Width           =   75
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1080
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   315
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   240
         Picture         =   "FO_system.frx":0614
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   360
         X2              =   360
         Y1              =   2160
         Y2              =   2520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H000040C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   360
         X2              =   2220
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   1695
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   300
         Width           =   2415
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   1695
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   3180
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FO_system"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

FO_Service.Show

End Sub

Private Sub Command2_Click()


Damper_count = Damper_count + 1
If Damper_count > 2 Then Damper_count = 0

    Port_damper_line(0).Visible = False
    Port_damper_line(1).Visible = False
    Port_damper_line(2).Visible = False
    Port_damper_line(3).Visible = False
    
    Port_damper_line(4).Visible = False
    Port_damper_line(5).Visible = False
    Port_damper_line(6).Visible = False
    Port_damper_line(7).Visible = False
    
    Port_damper_line(8).Visible = False
    Port_damper_line(9).Visible = False
    Port_damper_line(10).Visible = False
    Port_damper_line(11).Visible = False
    
If Damper_count = 0 Then
    Port_damper_line(0).Visible = True
    Port_damper_line(1).Visible = True
    Port_damper_line(2).Visible = True
    Port_damper_line(3).Visible = True
ElseIf Damper_count = 1 Then
    Port_damper_line(4).Visible = True
    Port_damper_line(5).Visible = True
    Port_damper_line(6).Visible = True
    Port_damper_line(7).Visible = True
Else
    Port_damper_line(8).Visible = True
    Port_damper_line(9).Visible = True
    Port_damper_line(10).Visible = True
    Port_damper_line(11).Visible = True
End If


For i = 0 To 11
    Port_damper_line(i + 12).Visible = Port_damper_line(i).Visible
Next i

For i = 0 To 23
    Stbd_damper_line(i).Visible = Port_damper_line(i).Visible
Next i

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idFO_system)

    Port_damper_line(0).Visible = False
    Port_damper_line(1).Visible = False
    Port_damper_line(2).Visible = False
    Port_damper_line(3).Visible = False
    
    Port_damper_line(4).Visible = True
    Port_damper_line(5).Visible = True
    Port_damper_line(6).Visible = True
    Port_damper_line(7).Visible = True
    
    Port_damper_line(8).Visible = False
    Port_damper_line(9).Visible = False
    Port_damper_line(10).Visible = False
    Port_damper_line(11).Visible = False
    
For i = 0 To 11
    Port_damper_line(i + 12).Visible = Port_damper_line(i).Visible
Next i
For i = 0 To 23
    Stbd_damper_line(i).Visible = Port_damper_line(i).Visible
Next i

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idFO_system)
    HavenStretch1.Visible = True
End If

End Sub
