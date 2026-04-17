VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Dtypeblr 
   Caption         =   "STARBOARD BOILER -STEAM SYSTEMS"
   ClientHeight    =   8970
   ClientLeft      =   1605
   ClientTop       =   585
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   12210
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      TabIndex        =   48
      Top             =   120
      Width           =   435
   End
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8775
      Left            =   -60
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   15478
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 2 OFF/ON"
         Height          =   495
         Index           =   2
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5160
         Width           =   1035
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1 OFF/ON"
         Height          =   495
         Index           =   1
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   6480
         Width           =   1035
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2415
         Index           =   1
         Left            =   960
         Max             =   20000
         TabIndex        =   45
         Top             =   5100
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0FFFF&
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
         Left            =   8160
         TabIndex        =   44
         Text            =   "T"
         Top             =   6840
         Width           =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   660
         Top             =   660
      End
      Begin VB.TextBox Tsteam1 
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
         Height          =   285
         Left            =   7080
         MousePointer    =   1  'Arrow
         TabIndex        =   39
         Text            =   "510 C"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Tsteam6 
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
         Height          =   285
         Left            =   9960
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         Text            =   "510 C"
         Top             =   6780
         Width           =   735
      End
      Begin VB.TextBox Psteam6 
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
         Height          =   285
         Left            =   9180
         MousePointer    =   1  'Arrow
         TabIndex        =   32
         Text            =   "60.3 Bar"
         Top             =   6780
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   11
         Left            =   5220
         TabIndex        =   26
         Top             =   1380
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BottomColour    =   65280
         EnabledTopLeft  =   0   'False
         ColourClosed    =   255
         BackColour      =   14745599
      End
      Begin VB.CommandButton Command3D4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -2460
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4500
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Psteam1 
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
         Height          =   285
         Left            =   7080
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Text            =   "60.3 Bar"
         Top             =   2040
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   4
         Left            =   10500
         TabIndex        =   10
         Top             =   1620
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   767
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   0
         Left            =   9240
         TabIndex        =   11
         Top             =   1620
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   5
         Left            =   4500
         TabIndex        =   12
         Top             =   300
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   6
         Left            =   9540
         TabIndex        =   13
         Top             =   2700
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   16777215
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   7
         Left            =   9660
         TabIndex        =   14
         Top             =   7680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   675
         Index           =   9
         Left            =   3120
         TabIndex        =   16
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1191
         TopColour       =   16711680
         BottomColour    =   16711680
         LineWidth       =   40
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   8
         Left            =   8760
         TabIndex        =   15
         Top             =   6840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   1
         Left            =   9960
         TabIndex        =   21
         Top             =   3750
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   2
         Left            =   10110
         TabIndex        =   22
         Top             =   4920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 HavenValveV1_22 
         Height          =   375
         Left            =   6540
         TabIndex        =   34
         Top             =   5910
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   661
         TopColour       =   65280
         BottomColour    =   65280
         ColourClosed    =   255
         BackColour      =   14745599
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   10
         Left            =   9780
         TabIndex        =   38
         Top             =   1620
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   873
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   3
         Left            =   10020
         TabIndex        =   40
         Top             =   7680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Line Line64 
         X1              =   6000
         X2              =   7140
         Y1              =   2460
         Y2              =   2280
      End
      Begin VB.Label Vattemp_POS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "39 %"
         Height          =   255
         Left            =   7260
         TabIndex        =   43
         Top             =   6420
         Width           =   495
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Attemperator"
         Height          =   255
         Left            =   7140
         TabIndex        =   42
         Top             =   6180
         Width           =   1035
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   2940
         Shape           =   2  'Oval
         Top             =   6360
         Width           =   375
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   2520
         Shape           =   2  'Oval
         Top             =   4740
         Width           =   375
      End
      Begin VB.Line Line83 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9630
         X2              =   9630
         Y1              =   2340
         Y2              =   2580
      End
      Begin VB.Line Line82 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9900
         X2              =   9900
         Y1              =   2280
         Y2              =   1260
      End
      Begin VB.Shape Shape17 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3000
         Top             =   6000
         Width           =   255
      End
      Begin VB.Shape Shape16 
         Height          =   1065
         Left            =   2550
         Shape           =   3  'Circle
         Top             =   5610
         Width           =   1155
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2670
         Shape           =   3  'Circle
         Top             =   5790
         Width           =   375
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2580
         Top             =   4380
         Width           =   255
      End
      Begin VB.Shape Shape8 
         Height          =   1065
         Left            =   2130
         Shape           =   3  'Circle
         Top             =   3990
         Width           =   1155
      End
      Begin VB.Line Line81 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5100
         X2              =   6120
         Y1              =   5250
         Y2              =   4830
      End
      Begin VB.Line Line80 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5580
         X2              =   6120
         Y1              =   5310
         Y2              =   5070
      End
      Begin VB.Line Line79 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5100
         X2              =   6120
         Y1              =   5730
         Y2              =   5310
      End
      Begin VB.Line Line78 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5580
         X2              =   6120
         Y1              =   5790
         Y2              =   5550
      End
      Begin VB.Line Line77 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5220
         X2              =   6240
         Y1              =   6240
         Y2              =   5820
      End
      Begin VB.Line Line76 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5700
         X2              =   6240
         Y1              =   6300
         Y2              =   6060
      End
      Begin VB.Line Line73 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5280
         X2              =   6300
         Y1              =   6750
         Y2              =   6330
      End
      Begin VB.Line Line72 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5760
         X2              =   6300
         Y1              =   6810
         Y2              =   6570
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5190
         X2              =   5070
         Y1              =   5910
         Y2              =   5940
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5400
         X2              =   5520
         Y1              =   5160
         Y2              =   5130
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Soot Blowers"
         Height          =   255
         Left            =   7680
         TabIndex        =   37
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   7020
         X2              =   7020
         Y1              =   2580
         Y2              =   7140
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "STEAM DRUM MOUNTINGS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9060
         TabIndex        =   36
         Tag             =   "Steam Drum"
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 1 Burner"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Top             =   6090
         Width           =   645
      End
      Begin VB.Line Line71 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5490
         X2              =   6030
         Y1              =   4830
         Y2              =   4590
      End
      Begin VB.Line Line70 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5400
         X2              =   5940
         Y1              =   4380
         Y2              =   4140
      End
      Begin VB.Line Line62 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5040
         X2              =   6060
         Y1              =   4770
         Y2              =   4350
      End
      Begin VB.Line Line47 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4980
         X2              =   6000
         Y1              =   4320
         Y2              =   3900
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   3075
         Index           =   2
         Left            =   8970
         Top             =   480
         Width           =   2355
      End
      Begin VB.Shape Shape7 
         FillStyle       =   0  'Solid
         Height          =   75
         Left            =   7980
         Top             =   7080
         Width           =   195
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6060
         Picture         =   "DtypeBlr.frx":0000
         Stretch         =   -1  'True
         Top             =   2820
         Width           =   255
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6540
         Picture         =   "DtypeBlr.frx":030A
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   6420
         X2              =   6060
         Y1              =   7020
         Y2              =   3060
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6660
         X2              =   6660
         Y1              =   7140
         Y2              =   2700
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5700
         Shape           =   3  'Circle
         Top             =   3540
         Width           =   195
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5700
         Shape           =   3  'Circle
         Top             =   3480
         Width           =   375
      End
      Begin VB.Line Line67 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   4860
         X2              =   5880
         Y1              =   3900
         Y2              =   3480
      End
      Begin VB.Line Line66 
         BorderColor     =   &H00000000&
         BorderStyle     =   2  'Dash
         X1              =   5280
         X2              =   5820
         Y1              =   3960
         Y2              =   3720
      End
      Begin VB.Line Line63 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4920
         X2              =   5010
         Y1              =   4620
         Y2              =   4590
      End
      Begin VB.Image Image27 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5880
         Picture         =   "DtypeBlr.frx":0614
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   255
      End
      Begin VB.Line Line44 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   8070
         X2              =   6870
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9900
         X2              =   9390
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5700
         X2              =   6480
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4860
         X2              =   7020
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6180
         X2              =   6180
         Y1              =   660
         Y2              =   2340
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4860
         X2              =   6180
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4860
         X2              =   4860
         Y1              =   2340
         Y2              =   2580
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   6000
         X2              =   5940
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   7200
         Picture         =   "DtypeBlr.frx":091E
         Stretch         =   -1  'True
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Economiser"
         Height          =   255
         Left            =   6600
         TabIndex        =   31
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label M 
         Caption         =   "Main Feed"
         Height          =   195
         Left            =   7920
         TabIndex        =   30
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Superheater Headers"
         Height          =   435
         Left            =   4020
         TabIndex        =   29
         Top             =   3720
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "WET STEAM"
         Height          =   255
         Left            =   2340
         TabIndex        =   28
         Top             =   540
         Width           =   1035
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3480
         Picture         =   "DtypeBlr.frx":0C28
         Stretch         =   -1  'True
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "SUPERHEATED STEAM"
         Height          =   195
         Left            =   9120
         TabIndex        =   27
         Top             =   6540
         Width           =   2115
      End
      Begin VB.Image Image15 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   9540
         Picture         =   "DtypeBlr.frx":0F32
         Stretch         =   -1  'True
         Top             =   7020
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1875
         Index           =   1
         Left            =   9000
         Top             =   3720
         Width           =   2355
      End
      Begin VB.Label Label5 
         Caption         =   "DRAIN"
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
         Left            =   9420
         TabIndex        =   25
         Top             =   5340
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "VENT"
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
         Left            =   9540
         TabIndex        =   24
         Top             =   3780
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "SUPERHEATER CONTROLS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   23
         Top             =   4620
         Width           =   2235
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9600
         X2              =   9840
         Y1              =   5040
         Y2              =   5280
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9840
         X2              =   10740
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   10680
         Picture         =   "DtypeBlr.frx":123C
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   255
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9480
         X2              =   9720
         Y1              =   4470
         Y2              =   4110
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9720
         X2              =   10560
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   10560
         Picture         =   "DtypeBlr.frx":1546
         Stretch         =   -1  'True
         Top             =   3990
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5160
         X2              =   6420
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   7860
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6420
         X2              =   6420
         Y1              =   2940
         Y2              =   5520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4560
         X2              =   5340
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4560
         X2              =   4560
         Y1              =   1200
         Y2              =   3420
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STEAM DRUM"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1260
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   2040
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line61 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5640
         X2              =   7860
         Y1              =   7980
         Y2              =   7980
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   6780
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   7380
         Shape           =   3  'Circle
         Top             =   1500
         Width           =   135
      End
      Begin VB.Line Line60 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6060
         X2              =   5880
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line59 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   7680
         X2              =   7680
         Y1              =   1800
         Y2              =   1980
      End
      Begin VB.Line Line58 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6300
         X2              =   7680
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line57 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6540
         X2              =   7680
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6540
         X2              =   6540
         Y1              =   1440
         Y2              =   1800
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   7680
         X2              =   7680
         Y1              =   1260
         Y2              =   1440
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6540
         X2              =   6540
         Y1              =   1080
         Y2              =   1260
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6540
         X2              =   7680
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6540
         X2              =   7680
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   6540
         X2              =   7980
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+9 in."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2700
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SCUMMING VALVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9120
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         Top             =   3210
         Width           =   1395
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   1380
         X2              =   1380
         Y1              =   3360
         Y2              =   7860
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   5340
         X2              =   1440
         Y1              =   7980
         Y2              =   7980
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 2 Burner"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   1860
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Top             =   4500
         Width           =   645
      End
      Begin VB.Line Line48 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   4680
         X2              =   4140
         Y1              =   7500
         Y2              =   3120
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   4140
         X2              =   4680
         Y1              =   3120
         Y2              =   2700
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   8070
         X2              =   8070
         Y1              =   7140
         Y2              =   6060
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-9 in."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   2700
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Top             =   2340
         Width           =   285
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   6480
         X2              =   6480
         Y1              =   1860
         Y2              =   660
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3360
         X2              =   4740
         Y1              =   1860
         Y2              =   2040
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3360
         X2              =   4680
         Y1              =   2400
         Y2              =   2280
      End
      Begin VB.Image Image12 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   10140
         Picture         =   "DtypeBlr.frx":1850
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   255
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   10380
         Picture         =   "DtypeBlr.frx":1B5A
         Stretch         =   -1  'True
         Top             =   7920
         Width           =   255
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9420
         X2              =   10260
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9300
         X2              =   10140
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AIR VENT"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   10680
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   1080
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SAFETY VALVES"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9420
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   870
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   3420
         X2              =   6180
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00404040&
         BorderWidth     =   4
         X1              =   1380
         X2              =   4620
         Y1              =   3360
         Y2              =   2400
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   5700
         X2              =   6420
         Y1              =   7680
         Y2              =   7020
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   5880
         X2              =   6060
         Y1              =   2880
         Y2              =   3060
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6660
         X2              =   5160
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WATER DRUM"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4980
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   8160
         Width           =   1185
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5340
         X2              =   5340
         Y1              =   1200
         Y2              =   1740
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5280
         X2              =   10080
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Shape burner2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   4170
         Width           =   375
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9180
         X2              =   9420
         Y1              =   7800
         Y2              =   8040
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9060
         X2              =   9300
         Y1              =   2820
         Y2              =   3060
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BLOWDOWN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9240
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   8340
         Width           =   945
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WATER DRUM MOUNTING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9150
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   7440
         Width           =   1980
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O/BOARD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   10500
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "O/BOARD"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   10620
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   7920
         Width           =   675
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   1215
         Index           =   0
         Left            =   8940
         Top             =   7380
         Width           =   2355
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   6420
         X2              =   5580
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   5160
         X2              =   5160
         Y1              =   2700
         Y2              =   2940
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   9360
         X2              =   9360
         Y1              =   2280
         Y2              =   1260
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   10620
         X2              =   10620
         Y1              =   1020
         Y2              =   2220
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         X1              =   4680
         X2              =   5220
         Y1              =   7500
         Y2              =   7740
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   4560
         X2              =   4800
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   7890
         X2              =   7890
         Y1              =   7980
         Y2              =   420
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   5160
         Shape           =   3  'Circle
         Top             =   7560
         Width           =   675
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00808080&
         BorderWidth     =   8
         X1              =   5220
         X2              =   5700
         Y1              =   3390
         Y2              =   6930
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         BorderWidth     =   8
         X1              =   4800
         X2              =   5280
         Y1              =   3420
         Y2              =   7140
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   1395
         Left            =   4560
         Shape           =   3  'Circle
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00C0FFFF&
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   7155
         Left            =   780
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   7995
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0FFFF&
         Caption         =   "30 %"
         Height          =   255
         Left            =   6720
         TabIndex        =   41
         Top             =   6300
         Width           =   495
      End
   End
End
Attribute VB_Name = "Dtypeblr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim set_dummy_focus As Integer



Private Sub Burner_ON_OFF_Click(Index As Integer)

Burner_ON(Index) = Not Burner_ON(Index)
If Burner_ON(Index) Then
    Burner_ON_OFF(Index).BackColor = RGB(0, &HFF, 0)
Else
    Burner_ON_OFF(Index).BackColor = RGB(&HC0, &HC0, &HC0)
End If

' SET in Steam_Sim.Dll
Call set_STEAM_logicals(Index, Burner_ON(Index))

set_dummy_focus = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idsy155)

set_dummy_focus = True

' 22/7/98
old_greenness_LSS(5) = -123    ' To ensure an update

' 22/7/98
   update_VBAI(35) = True

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy155)
    HavenStretch1.Visible = True
End If
End Sub









Private Sub Timer1_Timer()

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' update_vvalve_LSS(syxxx,nValve,index1,index2)
'   21/7/99
Call update_vvalve_LSS(Dtypeblr, 16, 11, -1)    ' VV16
Call update_vvalve_LSS(Dtypeblr, 16, 8, -1)    ' VV16


Psteam1.Text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
Tsteam1.Text = Format(LSS_reals(755) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

Psteam6.Text = Format(LSS_reals(788) * Press_mult, "##0.0") & Press_chars     ' Bar
Tsteam6.Text = Format(LSS_reals(790) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

Vattemp_pos.Caption = Format(LSS_reals(786), "##0") & " %"



End Sub

Private Sub valve_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
Select Case Index

' 20/7/99
Case 8
     Call valve_pcdi_LSS(16)
Case 11
     Call valve_pcdi_LSS(16)
    
End Select

End Sub

Private Sub VScroll1_Change(Index As Integer)

' 21/7/99
Vscroll_val(Index) = VScroll1(Index).Value

Call set_VSCROLL(Index, VScroll1(Index).Value)

End Sub
