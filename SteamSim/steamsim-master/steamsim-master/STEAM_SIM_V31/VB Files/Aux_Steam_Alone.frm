VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Object = "{479F4823-E7F1-11D0-901E-0020AF7543C2}#5.0#0"; "fxrtx50g.ocx"
Begin VB.Form Aux_Steam_Alone 
   Caption         =   "Auxiliary Steam"
   ClientHeight    =   12855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16770
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12855
   ScaleWidth      =   16770
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   12615
      Left            =   0
      Top             =   0
      Width           =   16485
      _ExtentX        =   29078
      _ExtentY        =   22251
      Begin VB.TextBox Tsteam10 
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
         Left            =   7800
         TabIndex        =   187
         Text            =   "955 F"
         Top             =   3210
         Width           =   795
      End
      Begin VB.CommandButton Command14 
         Height          =   315
         Left            =   5430
         TabIndex        =   186
         Top             =   8490
         Width           =   285
      End
      Begin VB.TextBox Psteam12 
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
         Left            =   15360
         TabIndex        =   184
         Text            =   " xx psi"
         Top             =   5310
         Width           =   675
      End
      Begin VB.TextBox Psteam11 
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
         Left            =   13950
         TabIndex        =   183
         Text            =   "xx psi"
         Top             =   5520
         Width           =   675
      End
      Begin VB.TextBox Tsteam11 
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
         Left            =   13950
         TabIndex        =   182
         Text            =   "xx F"
         Top             =   5850
         Width           =   675
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   600
         Top             =   2940
      End
      Begin VB.TextBox Psteam101 
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
         Left            =   12240
         MousePointer    =   1  'Arrow
         TabIndex        =   181
         Text            =   "xx Bar"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2790
         TabIndex        =   180
         Text            =   "P"
         Top             =   9870
         Width           =   195
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2790
         TabIndex        =   179
         Text            =   "P"
         Top             =   8520
         Width           =   195
      End
      Begin VB.CommandButton Command13 
         Height          =   315
         Left            =   4320
         TabIndex        =   178
         Top             =   12210
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   14400
         TabIndex        =   177
         Top             =   11250
         Width           =   285
      End
      Begin VB.CommandButton Command12 
         Height          =   315
         Left            =   13530
         TabIndex        =   176
         Top             =   11790
         Width           =   285
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   8130
         TabIndex        =   175
         Top             =   11250
         Width           =   285
      End
      Begin VB.CommandButton Command11 
         Height          =   315
         Left            =   6750
         TabIndex        =   174
         Top             =   11250
         Width           =   285
      End
      Begin VB.CommandButton Command10 
         Height          =   315
         Left            =   6660
         TabIndex        =   171
         Top             =   4050
         Width           =   285
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Left            =   3870
         TabIndex        =   170
         Top             =   7230
         Width           =   285
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   11460
         TabIndex        =   169
         Top             =   2340
         Width           =   285
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   2850
         TabIndex        =   168
         Top             =   2340
         Width           =   285
      End
      Begin VB.CommandButton Command6 
         Height          =   315
         Left            =   8310
         TabIndex        =   167
         Top             =   7260
         Width           =   285
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   8310
         TabIndex        =   166
         Top             =   6750
         Width           =   285
      End
      Begin VB.TextBox P_35 
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
         Left            =   5220
         TabIndex        =   165
         Text            =   "x.x Bar"
         Top             =   9840
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   3870
         TabIndex        =   164
         Top             =   8640
         Width           =   285
      End
      Begin VB.CommandButton SP_HP 
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9000
         TabIndex        =   162
         Top             =   4440
         Width           =   375
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   23
         Left            =   9750
         TabIndex        =   53
         Top             =   5670
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin VB.TextBox Psteam6 
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
         Left            =   7800
         TabIndex        =   139
         Text            =   "879 psi"
         Top             =   2880
         Width           =   795
      End
      Begin VB.HScrollBar Hscroll_HP_Bypass 
         Height          =   135
         Left            =   9360
         Max             =   1000
         TabIndex        =   120
         Top             =   6150
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   1335
         Left            =   5580
         Max             =   1000
         TabIndex        =   118
         Top             =   10530
         Width           =   135
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4800
         TabIndex        =   107
         Text            =   "P"
         Top             =   10290
         Width           =   195
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1335
         Left            =   14280
         Max             =   1000
         TabIndex        =   106
         Top             =   270
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox P_IP_reg_IN 
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
         Left            =   11160
         TabIndex        =   105
         Text            =   "x.x Bar"
         Top             =   9840
         Width           =   795
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   1335
         Left            =   14670
         Max             =   1000
         TabIndex        =   104
         Top             =   210
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
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
         Left            =   10860
         TabIndex        =   88
         Text            =   "P"
         Top             =   9870
         Width           =   195
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   9420
         Max             =   1000
         TabIndex        =   87
         Top             =   10980
         Width           =   1215
      End
      Begin VB.CommandButton SP_IP 
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9270
         TabIndex        =   86
         Top             =   9120
         Width           =   375
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   1260
         Max             =   1000
         TabIndex        =   74
         Top             =   10800
         Width           =   1215
      End
      Begin VB.CommandButton SP_875_35 
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   73
         Top             =   9300
         Width           =   375
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8790
         TabIndex        =   68
         Text            =   "P"
         Top             =   5100
         Width           =   195
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1155
         Left            =   10590
         Max             =   1000
         TabIndex        =   51
         Top             =   6090
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox P_HP_reg_IN 
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
         Left            =   10350
         TabIndex        =   48
         Text            =   "x.x Bar"
         Top             =   4890
         Width           =   795
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   2820
         TabIndex        =   21
         Text            =   "P"
         Top             =   5100
         Width           =   195
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2820
         TabIndex        =   20
         Text            =   "P"
         Top             =   7020
         Width           =   195
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   1260
         Max             =   1000
         TabIndex        =   19
         Top             =   6120
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   135
         Left            =   1260
         Max             =   1000
         TabIndex        =   18
         Top             =   8040
         Width           =   1215
      End
      Begin VB.TextBox P_150 
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
         Left            =   5220
         TabIndex        =   17
         Text            =   "x.x Bar"
         Top             =   5040
         Width           =   795
      End
      Begin VB.TextBox P_875_desup_OUT 
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
         Left            =   7230
         TabIndex        =   16
         Text            =   "x.x Bar"
         Top             =   1590
         Width           =   795
      End
      Begin VB.TextBox P_Atom 
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
         Left            =   3660
         TabIndex        =   15
         Text            =   "x.x Bar"
         Top             =   6810
         Width           =   795
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   8310
         TabIndex        =   14
         Top             =   6240
         Width           =   285
      End
      Begin VB.CommandButton SP_875_150 
         Caption         =   "SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   13
         Top             =   4440
         Width           =   375
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
         Height          =   315
         Left            =   1740
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Text            =   "xx Bar"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Drum_water_level_stbd 
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
         Left            =   1740
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Text            =   "-xx in"
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox Drum_water_level_port 
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
         Left            =   12240
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Text            =   "xx in"
         Top             =   1260
         Width           =   735
      End
      Begin FXRTX50GLib.FXRtext FXRtext2 
         Height          =   1875
         Left            =   13260
         TabIndex        =   7
         Top             =   8340
         Width           =   315
         _Version        =   327680
         _ExtentX        =   556
         _ExtentY        =   3307
         _StockProps     =   70
         Caption         =   "Astern Element"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Persistence     =   -1  'True
         _StdProps       =   -4088998
         ErrStr          =   "JAMHAOMAAPLFGLLCMMBPNGHFBINOJIJMLGCFPBGKNDEIKL"
         ErrCode         =   1713846826
         ErrInfo         =   -718973410
         Angle           =   90
         BackColor       =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   8
         Left            =   690
         TabIndex        =   22
         Top             =   4860
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   88
         Left            =   2025
         TabIndex        =   23
         Top             =   5145
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   89
         Left            =   1260
         TabIndex        =   24
         Top             =   5145
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   20
         Left            =   1620
         TabIndex        =   25
         Top             =   5610
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   106
         Left            =   5880
         TabIndex        =   26
         Top             =   6150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   115
         Left            =   2010
         TabIndex        =   27
         Top             =   7065
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   116
         Left            =   1260
         TabIndex        =   28
         Top             =   7065
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   47
         Left            =   1620
         TabIndex        =   29
         Top             =   7530
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   53
         Left            =   780
         TabIndex        =   30
         Top             =   6750
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   126
         Left            =   5880
         TabIndex        =   31
         Top             =   5670
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
         Index           =   76
         Left            =   10080
         TabIndex        =   49
         Top             =   5130
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   92
         Left            =   9330
         TabIndex        =   50
         Top             =   5130
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   79
         Left            =   11940
         TabIndex        =   52
         Top             =   7170
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
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
         Index           =   94
         Left            =   10830
         TabIndex        =   54
         Top             =   5130
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   28
         Left            =   2850
         TabIndex        =   75
         Top             =   9450
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   49152
         BackColour      =   12632256
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   107
         Left            =   1920
         TabIndex        =   76
         Top             =   9900
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   108
         Left            =   1260
         TabIndex        =   77
         Top             =   9900
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   39
         Left            =   1620
         TabIndex        =   78
         Top             =   10320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   40
         Left            =   10320
         TabIndex        =   89
         Top             =   9795
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   49152
         BackColour      =   12632256
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   111
         Left            =   9300
         TabIndex        =   90
         Top             =   9900
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   112
         Left            =   9960
         TabIndex        =   91
         Top             =   9930
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   43
         Left            =   9720
         TabIndex        =   92
         Top             =   10395
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   74
         Left            =   12540
         TabIndex        =   103
         Top             =   7170
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   77
         Left            =   4350
         TabIndex        =   108
         Top             =   10560
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   81
         Left            =   4350
         TabIndex        =   109
         Top             =   11220
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   21
         Left            =   4950
         TabIndex        =   110
         Top             =   10980
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   50
         Left            =   3990
         TabIndex        =   111
         Top             =   11400
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   49152
         BottomColour    =   49152
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   11
         ColourOpen      =   49152
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   438
         Left            =   3570
         TabIndex        =   121
         Top             =   2490
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
         Index           =   468
         Left            =   10590
         TabIndex        =   122
         Top             =   2490
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   283
         Left            =   8940
         TabIndex        =   124
         Top             =   2490
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   284
         Left            =   5580
         TabIndex        =   125
         Top             =   2490
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
         Height          =   255
         Index           =   293
         Left            =   12210
         TabIndex        =   127
         Top             =   3090
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   0
         Left            =   12090
         TabIndex        =   128
         Top             =   3330
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   873
         TopColour       =   16384
         BottomColour    =   16384
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16384
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   9
         Left            =   12150
         TabIndex        =   129
         Top             =   3750
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   435
         Left            =   3570
         TabIndex        =   135
         Top             =   1200
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
         Index           =   465
         Left            =   10590
         TabIndex        =   137
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   577
         Left            =   1260
         TabIndex        =   147
         Top             =   8550
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   104
         Left            =   5880
         TabIndex        =   157
         Top             =   6630
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
         Index           =   105
         Left            =   5880
         TabIndex        =   159
         Top             =   7110
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
      Begin VB.Label Reg_875_225_pos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "xxx %"
         Height          =   255
         Left            =   1650
         TabIndex        =   190
         Top             =   8970
         Width           =   495
      End
      Begin VB.Label REG_120_pos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "xxx %"
         Height          =   255
         Left            =   2970
         TabIndex        =   189
         Top             =   8970
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "H.P. Heater"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   185
         Top             =   8250
         Width           =   1575
      End
      Begin VB.Image Arrow_875 
         Height          =   255
         Index           =   4
         Left            =   5430
         Picture         =   "Aux_Steam_Alone.frx":0000
         Stretch         =   -1  'True
         Top             =   7950
         Width           =   255
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   22
         X1              =   4860
         X2              =   4860
         Y1              =   10590
         Y2              =   11100
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   21
         X1              =   2880
         X2              =   2640
         Y1              =   10095
         Y2              =   9720
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   20
         X1              =   1770
         X2              =   1770
         Y1              =   9750
         Y2              =   9900
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   19
         X1              =   2610
         X2              =   1770
         Y1              =   9720
         Y2              =   9720
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   18
         X1              =   3330
         X2              =   3330
         Y1              =   8310
         Y2              =   8460
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   16
         X1              =   4170
         X2              =   3330
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   15
         X1              =   2910
         X2              =   2670
         Y1              =   8685
         Y2              =   8310
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   13
         X1              =   1800
         X2              =   1800
         Y1              =   8340
         Y2              =   8490
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   12
         X1              =   2640
         X2              =   1800
         Y1              =   8310
         Y2              =   8310
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   11
         X1              =   2910
         X2              =   2670
         Y1              =   7245
         Y2              =   6870
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   10
         X1              =   1800
         X2              =   1800
         Y1              =   6900
         Y2              =   7050
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   2640
         X2              =   1800
         Y1              =   6870
         Y2              =   6870
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   14
         X1              =   9090
         X2              =   8910
         Y1              =   4950
         Y2              =   5220
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   8
         X1              =   9900
         X2              =   9900
         Y1              =   4950
         Y2              =   5070
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   9870
         X2              =   9090
         Y1              =   4950
         Y2              =   4950
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   2910
         X2              =   2670
         Y1              =   5325
         Y2              =   4950
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   1800
         X2              =   1800
         Y1              =   4980
         Y2              =   5130
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   2640
         X2              =   1800
         Y1              =   4950
         Y2              =   4950
      End
      Begin VB.Label Warming_label 
         BackStyle       =   0  'Transparent
         Caption         =   " Warming through valve"
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
         Left            =   12540
         TabIndex        =   131
         Top             =   3030
         Width           =   2025
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   4
         Left            =   11190
         Picture         =   "Aux_Steam_Alone.frx":030A
         Stretch         =   -1  'True
         Top             =   3570
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   0
         Left            =   1860
         Picture         =   "Aux_Steam_Alone.frx":0614
         Stretch         =   -1  'True
         Top             =   3570
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   5
         X1              =   3720
         X2              =   3720
         Y1              =   3690
         Y2              =   4530
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   10
         Left            =   5160
         Picture         =   "Aux_Steam_Alone.frx":091E
         Stretch         =   -1  'True
         Top             =   4410
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   9
         Left            =   5160
         Picture         =   "Aux_Steam_Alone.frx":0C28
         Stretch         =   -1  'True
         Top             =   4050
         Width           =   255
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To Cargo Pumps"
         Height          =   195
         Index           =   3
         Left            =   5460
         TabIndex        =   173
         Top             =   4440
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   2
         X1              =   3720
         X2              =   5250
         Y1              =   4530
         Y2              =   4530
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To Feed Pumps"
         Height          =   195
         Index           =   2
         Left            =   5460
         TabIndex        =   172
         Top             =   4080
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   0
         X1              =   3750
         X2              =   5370
         Y1              =   4170
         Y2              =   4170
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   1
         Left            =   13560
         Picture         =   "Aux_Steam_Alone.frx":0F32
         Stretch         =   -1  'True
         Top             =   11280
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   4
         Left            =   15120
         Picture         =   "Aux_Steam_Alone.frx":123C
         Stretch         =   -1  'True
         Top             =   7830
         Width           =   255
      End
      Begin VB.Image Arrow_875 
         Height          =   255
         Index           =   3
         Left            =   8130
         Picture         =   "Aux_Steam_Alone.frx":1546
         Stretch         =   -1  'True
         Top             =   10710
         Width           =   255
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "To DFT Tank"
         Height          =   195
         Index           =   2
         Left            =   8010
         TabIndex        =   163
         Top             =   11040
         Width           =   1185
      End
      Begin VB.Image Image13 
         Height          =   255
         Index           =   4
         Left            =   6750
         Picture         =   "Aux_Steam_Alone.frx":1850
         Stretch         =   -1  'True
         Top             =   10710
         Width           =   255
      End
      Begin VB.Image Arrow_875 
         Height          =   255
         Index           =   2
         Left            =   4320
         Picture         =   "Aux_Steam_Alone.frx":1B5A
         Stretch         =   -1  'True
         Top             =   10290
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   1
         Left            =   3150
         Picture         =   "Aux_Steam_Alone.frx":1E64
         Stretch         =   -1  'True
         Top             =   10110
         Width           =   255
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "From Feed Pump Exhaust"
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   161
         Top             =   11040
         Width           =   2205
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   12
         Left            =   6840
         Top             =   10260
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   555
         Index           =   1
         Left            =   8220
         Top             =   10230
         Width           =   75
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   6420
         Picture         =   "Aux_Steam_Alone.frx":216E
         Stretch         =   -1  'True
         Top             =   7290
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   5580
         X2              =   6630
         Y1              =   7410
         Y2              =   7410
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Distil Air Ejectors"
         Height          =   255
         Index           =   0
         Left            =   6720
         TabIndex        =   160
         Top             =   7320
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Main Air Ejectors"
         Height          =   315
         Left            =   6720
         TabIndex        =   158
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   6420
         Picture         =   "Aux_Steam_Alone.frx":2478
         Stretch         =   -1  'True
         Top             =   6840
         Width           =   255
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   11
         Left            =   9090
         Top             =   10680
         Width           =   675
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   555
         Index           =   10
         Left            =   9030
         Top             =   10200
         Width           =   75
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contaminated Steam Evaporator"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   156
         Top             =   9060
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   6
         Left            =   3570
         Picture         =   "Aux_Steam_Alone.frx":2782
         Stretch         =   -1  'True
         Top             =   8730
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   660
         Picture         =   "Aux_Steam_Alone.frx":2A8C
         Stretch         =   -1  'True
         Top             =   8760
         Width           =   255
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":2D96
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   155
         Top             =   8430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":ABAE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   154
         Top             =   8430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":12DC6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   153
         Top             =   8430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_875_225 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":1B7DE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   152
         Top             =   8430
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   3
         Left            =   3120
         OleObjectBlob   =   "Aux_Steam_Alone.frx":23DF6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   151
         Top             =   8430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   0
         Left            =   3120
         OleObjectBlob   =   "Aux_Steam_Alone.frx":2BC0E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   150
         Top             =   8430
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   1
         Left            =   3120
         OleObjectBlob   =   "Aux_Steam_Alone.frx":34226
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   149
         Top             =   8430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Reg_120 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Index           =   2
         Left            =   3120
         OleObjectBlob   =   "Aux_Steam_Alone.frx":3CC3E
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   148
         Top             =   8430
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         X1              =   450
         X2              =   1350
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   10020
         X2              =   11490
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   3
         X1              =   9810
         X2              =   8940
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   7
         Left            =   5010
         Top             =   10260
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   405
         Index           =   3
         Left            =   4410
         Top             =   10230
         Width           =   75
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "35 # Dump Regulator"
         Height          =   195
         Index           =   3
         Left            =   2370
         TabIndex        =   146
         Top             =   11160
         Width           =   1875
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   6
         Left            =   13530
         Picture         =   "Aux_Steam_Alone.frx":44E56
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   11910
         Picture         =   "Aux_Steam_Alone.frx":45160
         Stretch         =   -1  'True
         Top             =   6060
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   5
         Left            =   11610
         Picture         =   "Aux_Steam_Alone.frx":4546A
         Stretch         =   -1  'True
         Top             =   5310
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Condenser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   145
         Top             =   12000
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   3
         Left            =   6420
         Picture         =   "Aux_Steam_Alone.frx":45774
         Stretch         =   -1  'True
         Top             =   6330
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   14
         X1              =   5550
         X2              =   6630
         Y1              =   6450
         Y2              =   6450
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   13
         X1              =   5550
         X2              =   5550
         Y1              =   8100
         Y2              =   5460
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   5
         Left            =   3150
         Picture         =   "Aux_Steam_Alone.frx":45A7E
         Stretch         =   -1  'True
         Top             =   5310
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   2
         Left            =   6420
         Picture         =   "Aux_Steam_Alone.frx":45D88
         Stretch         =   -1  'True
         Top             =   5850
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   4
         Left            =   13530
         Picture         =   "Aux_Steam_Alone.frx":46092
         Stretch         =   -1  'True
         Top             =   5310
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Condenser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   12840
         TabIndex        =   144
         Top             =   11550
         Width           =   1665
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   4
         X1              =   8940
         X2              =   8940
         Y1              =   5940
         Y2              =   5430
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   4
         Left            =   690
         Picture         =   "Aux_Steam_Alone.frx":4639C
         Stretch         =   -1  'True
         Top             =   10080
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   3
         Left            =   690
         Picture         =   "Aux_Steam_Alone.frx":466A6
         Stretch         =   -1  'True
         Top             =   7230
         Width           =   255
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   6
         Left            =   2250
         Top             =   10200
         Width           =   7095
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   5
         Left            =   2580
         Top             =   10200
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   4
         Left            =   1950
         Top             =   10590
         Width           =   675
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   27
         X1              =   2520
         X2              =   2520
         Y1              =   7830
         Y2              =   7350
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   2
         X1              =   2520
         X2              =   1920
         Y1              =   7830
         Y2              =   7830
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 # / 150 # Atomising Steam Regulator"
         Height          =   195
         Index           =   3
         Left            =   1170
         TabIndex        =   143
         Top             =   6540
         Width           =   3165
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Throttle (Ahd.)"
         Height          =   255
         Index           =   1
         Left            =   11730
         TabIndex        =   142
         Top             =   4200
         Width           =   1635
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   3
         Left            =   5460
         Picture         =   "Aux_Steam_Alone.frx":469B0
         Stretch         =   -1  'True
         Top             =   3570
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   12
         X1              =   480
         X2              =   480
         Y1              =   3690
         Y2              =   10200
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   12360
         X2              =   13140
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   2
         Left            =   6300
         Picture         =   "Aux_Steam_Alone.frx":46CBA
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   2
         Left            =   8250
         Picture         =   "Aux_Steam_Alone.frx":46FC4
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   255
      End
      Begin VB.Image Arrow_875 
         Height          =   255
         Index           =   1
         Left            =   7020
         Picture         =   "Aux_Steam_Alone.frx":472CE
         Stretch         =   -1  'True
         Top             =   3150
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   6
         X1              =   480
         X2              =   7140
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   3
         X1              =   7140
         X2              =   7140
         Y1              =   2910
         Y2              =   3690
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   2
         X1              =   7140
         X2              =   7140
         Y1              =   1500
         Y2              =   2700
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   7650
         Top             =   720
         Width           =   675
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   8
         X1              =   8310
         X2              =   7650
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   5010
         X2              =   5670
         Y1              =   210
         Y2              =   210
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   14
         X1              =   7650
         X2              =   8310
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DeSuperheated Steam  Stop v/v"
         Height          =   255
         Index           =   7
         Left            =   8670
         TabIndex        =   138
         Top             =   1020
         Width           =   2535
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   8
         X1              =   2730
         X2              =   12000
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "DeSuperheated Steam  Stop v/v"
         Height          =   255
         Index           =   2
         Left            =   3390
         TabIndex        =   136
         Top             =   1020
         Width           =   2535
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  35 # Regulated"
         Height          =   195
         Index           =   2
         Left            =   8490
         TabIndex        =   134
         Top             =   690
         Width           =   1335
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   1
         Left            =   8910
         Picture         =   "Aux_Steam_Alone.frx":475D8
         Stretch         =   -1  'True
         Top             =   3570
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bulkhead v/v"
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
         Index           =   2
         Left            =   8070
         TabIndex        =   133
         Top             =   2310
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Main Steam Stop v/v"
         Height          =   255
         Index           =   1
         Left            =   9480
         TabIndex        =   132
         Top             =   2310
         Width           =   1695
      End
      Begin VB.Image Image9 
         Height          =   255
         Index           =   1
         Left            =   8220
         Picture         =   "Aux_Steam_Alone.frx":478E2
         Stretch         =   -1  'True
         Top             =   2670
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   1
         Left            =   6270
         Picture         =   "Aux_Steam_Alone.frx":47BEC
         Stretch         =   -1  'True
         Top             =   2670
         Width           =   255
      End
      Begin VB.Image Arrow_875 
         Height          =   255
         Index           =   0
         Left            =   7350
         Picture         =   "Aux_Steam_Alone.frx":47EF6
         Stretch         =   -1  'True
         Top             =   3150
         Width           =   255
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   12840
         X2              =   12840
         Y1              =   4050
         Y2              =   3690
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   0
         X1              =   11880
         X2              =   11880
         Y1              =   3690
         Y2              =   4050
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   0
         Left            =   13290
         Picture         =   "Aux_Steam_Alone.frx":48200
         Stretch         =   -1  'True
         Top             =   3570
         Width           =   255
      End
      Begin VB.Label VV_pos_throttle_ahead 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12510
         TabIndex        =   130
         Top             =   3390
         Width           =   555
      End
      Begin VB.Line Warm_line 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   11880
         X2              =   12240
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   12420
         X2              =   13740
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   1
         X1              =   11880
         X2              =   12240
         Y1              =   4050
         Y2              =   4050
      End
      Begin VB.Line Warm_line 
         BorderColor     =   &H00E0E0E0&
         Index           =   1
         X1              =   11880
         X2              =   11880
         Y1              =   3270
         Y2              =   3690
      End
      Begin VB.Line Warm_line 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   13140
         X2              =   13140
         Y1              =   3270
         Y2              =   3720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulkhead v/v"
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
         Index           =   1
         Left            =   5160
         TabIndex        =   126
         Top             =   2310
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Steam Stop v/v"
         Height          =   255
         Index           =   4
         Left            =   3330
         TabIndex        =   123
         Top             =   2310
         Width           =   1725
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   11490
         X2              =   11490
         Y1              =   5970
         Y2              =   5430
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gland Seal Regulator"
         Height          =   315
         Left            =   6720
         TabIndex        =   119
         Top             =   6330
         Width           =   1905
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   10560
         X2              =   10560
         Y1              =   10695
         Y2              =   10215
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   10650
         X2              =   9810
         Y1              =   9720
         Y2              =   9720
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   4350
         Picture         =   "Aux_Steam_Alone.frx":4850A
         Stretch         =   -1  'True
         Top             =   11760
         Width           =   255
      End
      Begin VB.Label BYPASS_35_dump 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5070
         TabIndex        =   112
         Top             =   10740
         Width           =   555
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   4350
         OleObjectBlob   =   "Aux_Steam_Alone.frx":48814
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   117
         Top             =   10920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4350
         OleObjectBlob   =   "Aux_Steam_Alone.frx":5062C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   116
         Top             =   10920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   4350
         OleObjectBlob   =   "Aux_Steam_Alone.frx":58644
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   115
         Top             =   10920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OLE Dump_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4350
         OleObjectBlob   =   "Aux_Steam_Alone.frx":6065C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   114
         Top             =   10920
         Width           =   375
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   37
         X1              =   4470
         X2              =   4470
         Y1              =   11520
         Y2              =   11820
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   39
         X1              =   4470
         X2              =   5010
         Y1              =   11760
         Y2              =   11760
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   17
         X1              =   4830
         X2              =   4650
         Y1              =   11100
         Y2              =   11100
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   5040
         X2              =   5040
         Y1              =   11280
         Y2              =   11760
      End
      Begin VB.Label VV_pos_dump_reg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3750
         TabIndex        =   113
         Top             =   10980
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   12030
         X2              =   12030
         Y1              =   5430
         Y2              =   10200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   12630
         X2              =   12630
         Y1              =   6840
         Y2              =   10230
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cascade v/v"
         Height          =   195
         Index           =   2
         Left            =   10740
         TabIndex        =   102
         Top             =   7290
         Width           =   1155
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IP Regulator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9750
         TabIndex        =   101
         Top             =   9150
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   9660
         OleObjectBlob   =   "Aux_Steam_Alone.frx":68E74
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   100
         Top             =   9855
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   9660
         OleObjectBlob   =   "Aux_Steam_Alone.frx":70C8C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   99
         Top             =   9855
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   9660
         OleObjectBlob   =   "Aux_Steam_Alone.frx":78EA4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   98
         Top             =   9855
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_IP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   9660
         OleObjectBlob   =   "Aux_Steam_Alone.frx":818BC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   97
         Top             =   9855
         Width           =   315
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   15450
         TabIndex        =   96
         Top             =   1650
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "35 #"
         Height          =   315
         Left            =   4440
         TabIndex        =   95
         Top             =   9930
         Width           =   375
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   9810
         X2              =   9810
         Y1              =   9750
         Y2              =   9900
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   10920
         X2              =   10680
         Y1              =   10095
         Y2              =   9720
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   9780
         X2              =   12630
         Y1              =   10230
         Y2              =   10230
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   10560
         X2              =   10080
         Y1              =   10695
         Y2              =   10695
      End
      Begin VB.Label VV_pos_IP_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9600
         TabIndex        =   94
         Top             =   9435
         Width           =   555
      End
      Begin VB.Label BYPASS_IP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10020
         TabIndex        =   93
         Top             =   10395
         Width           =   555
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 # / 35 # Regulator"
         Height          =   195
         Index           =   0
         Left            =   1230
         TabIndex        =   85
         Top             =   9300
         Width           =   1875
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":89ED4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   84
         Top             =   9840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":91CEC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   83
         Top             =   9840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":99F04
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   82
         Top             =   9840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_875_35 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":A291C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   81
         Top             =   9840
         Width           =   315
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   21
         X1              =   1080
         X2              =   1080
         Y1              =   10620
         Y2              =   10200
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   15
         X1              =   5580
         X2              =   6630
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   16
         X1              =   1080
         X2              =   1800
         Y1              =   10620
         Y2              =   10620
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   10
         X1              =   480
         X2              =   1320
         Y1              =   10200
         Y2              =   10200
      End
      Begin VB.Label VV_pos_875_35_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1620
         TabIndex        =   80
         Top             =   9510
         Width           =   555
      End
      Begin VB.Label BYPASS_875_35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   79
         Top             =   10320
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   22
         X1              =   6600
         X2              =   5580
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "L.P. Bleed"
         Height          =   255
         Index           =   0
         Left            =   14100
         TabIndex        =   72
         Top             =   10980
         Width           =   915
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   3
         Left            =   14400
         Picture         =   "Aux_Steam_Alone.frx":AAF34
         Stretch         =   -1  'True
         Top             =   10680
         Width           =   255
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Un-Regulated"
         Height          =   195
         Index           =   1
         Left            =   5820
         TabIndex        =   71
         Top             =   150
         Width           =   1995
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   0
         Left            =   14280
         Picture         =   "Aux_Steam_Alone.frx":AB23E
         Stretch         =   -1  'True
         Top             =   5130
         Width           =   255
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 #  Regulated"
         Height          =   195
         Left            =   8490
         TabIndex        =   70
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0C0&
         Caption         =   "150 # Regulated"
         Height          =   195
         Index           =   0
         Left            =   8490
         TabIndex        =   69
         Top             =   420
         Width           =   1335
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   0
         Left            =   3570
         Picture         =   "Aux_Steam_Alone.frx":AB548
         Stretch         =   -1  'True
         Top             =   7230
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   9
         X1              =   2520
         X2              =   2520
         Y1              =   5910
         Y2              =   5430
      End
      Begin VB.OLE Reg_HP_new 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   3
         Left            =   9690
         OleObjectBlob   =   "Aux_Steam_Alone.frx":AB852
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   65
         Top             =   5040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_HP_new 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   9690
         OleObjectBlob   =   "Aux_Steam_Alone.frx":B386A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   64
         Top             =   5040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "H.P. Bleed"
         Height          =   315
         Index           =   1
         Left            =   13320
         TabIndex        =   63
         Top             =   5100
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.P. Bleed v/v"
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
         Left            =   13020
         TabIndex        =   62
         Top             =   7260
         Width           =   1035
      End
      Begin VB.Label BYPASS_HP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10140
         TabIndex        =   57
         Top             =   5700
         Width           =   555
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   15150
         OleObjectBlob   =   "Aux_Steam_Alone.frx":BBA82
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   61
         Top             =   2070
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   2
         Left            =   14580
         OleObjectBlob   =   "Aux_Steam_Alone.frx":C389A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   60
         Top             =   2070
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   1
         Left            =   13980
         OleObjectBlob   =   "Aux_Steam_Alone.frx":CB8B2
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   59
         Top             =   2070
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OLE Reg_HP 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   0
         Left            =   13470
         OleObjectBlob   =   "Aux_Steam_Alone.frx":D40CA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   58
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label VV_pos_HP_reg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9570
         TabIndex        =   56
         Top             =   4680
         Width           =   555
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "HP Bleed Regulator"
         Height          =   195
         Left            =   9120
         TabIndex        =   55
         Top             =   4440
         Width           =   2025
      End
      Begin VB.OLE Reg_Atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   0
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":DC0E2
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   44
         Top             =   6990
         Width           =   405
      End
      Begin VB.OLE Reg_875_150 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   3
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":E46FA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   43
         Top             =   5070
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_150 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":EC712
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   42
         Top             =   5070
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_150 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   1
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":F492A
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   41
         Top             =   5070
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_875_150 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   0
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":FD342
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   40
         Top             =   5070
         Width           =   405
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   4
         X1              =   1080
         X2              =   1080
         Y1              =   5910
         Y2              =   5430
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   10
         X1              =   1680
         X2              =   1080
         Y1              =   5910
         Y2              =   5910
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   25
         X1              =   1320
         X2              =   510
         Y1              =   7350
         Y2              =   7350
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   26
         X1              =   1080
         X2              =   1080
         Y1              =   7830
         Y2              =   7350
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   28
         X1              =   1740
         X2              =   1080
         Y1              =   7830
         Y2              =   7830
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 # Desup Steam"
         Height          =   255
         Index           =   0
         Left            =   8070
         TabIndex        =   37
         Top             =   1650
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image14 
         Height          =   255
         Index           =   0
         Left            =   690
         Picture         =   "Aux_Steam_Alone.frx":10595A
         Stretch         =   -1  'True
         Top             =   5310
         Width           =   255
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   1860
         X2              =   1800
         Y1              =   5070
         Y2              =   5130
      End
      Begin VB.Label VV_pos_875_150_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1650
         TabIndex        =   35
         Top             =   4680
         Width           =   555
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hotel Services"
         Height          =   195
         Index           =   0
         Left            =   6720
         TabIndex        =   34
         Top             =   5880
         Width           =   1305
      End
      Begin VB.Image Image25 
         Height          =   255
         Left            =   7890
         Picture         =   "Aux_Steam_Alone.frx":105C64
         Stretch         =   -1  'True
         Top             =   5310
         Width           =   255
      End
      Begin VB.Label BYPASS_atom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   32
         Top             =   7590
         Width           =   555
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   0
         X1              =   2520
         X2              =   1980
         Y1              =   5910
         Y2              =   5910
      End
      Begin VB.Image Image10 
         Height          =   255
         Index           =   2
         Left            =   15120
         Picture         =   "Aux_Steam_Alone.frx":105F6E
         Stretch         =   -1  'True
         Top             =   5430
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cross Over Pipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   15300
         TabIndex        =   12
         Top             =   5850
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   1
         X1              =   13680
         X2              =   13680
         Y1              =   10860
         Y2              =   11400
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   14400
         X2              =   14400
         Y1              =   5040
         Y2              =   5430
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   14520
         X2              =   14520
         Y1              =   10500
         Y2              =   10800
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Height          =   255
         Left            =   13920
         TabIndex        =   10
         Top             =   9060
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Left            =   13800
         TabIndex        =   8
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00004080&
         BorderWidth     =   3
         Height          =   1215
         Index           =   0
         Left            =   12000
         Shape           =   3  'Circle
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Port Boiler"
         Height          =   255
         Index           =   0
         Left            =   12420
         TabIndex        =   6
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Drum"
         Height          =   255
         Index           =   3
         Left            =   12180
         TabIndex        =   5
         Top             =   1020
         Width           =   915
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00004080&
         BorderWidth     =   3
         Height          =   1215
         Index           =   1
         Left            =   1500
         Shape           =   3  'Circle
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Boiler"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   4
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Steam Drum"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   3
         Top             =   1020
         Width           =   915
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   2775
         Index           =   1
         Left            =   1110
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   2175
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4005
         Left            =   13140
         OleObjectBlob   =   "Aux_Steam_Alone.frx":106278
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   11
         Top             =   7440
         Width           =   2295
      End
      Begin VB.Line LP_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   15240
         X2              =   15240
         Y1              =   5280
         Y2              =   8580
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   3720
         X2              =   2280
         Y1              =   7350
         Y2              =   7350
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 # / 150 # Regulator"
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   39
         Top             =   4440
         Width           =   1965
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Atomising Steam to  burners"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   38
         Top             =   7560
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.OLE Reg_Atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   2
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":10DC90
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   47
         Top             =   6990
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_Atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   3
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":115EA8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   46
         Top             =   6990
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_Atomising 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   1
         Left            =   1620
         OleObjectBlob   =   "Aux_Steam_Alone.frx":11DEC0
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   45
         Top             =   6990
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_HP_new 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   0
         Left            =   9690
         OleObjectBlob   =   "Aux_Steam_Alone.frx":1268D8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   67
         Top             =   5040
         Width           =   405
      End
      Begin VB.OLE Reg_HP_new 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Index           =   1
         Left            =   9690
         OleObjectBlob   =   "Aux_Steam_Alone.frx":12EEF0
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   66
         Top             =   5040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   8
         X1              =   9420
         X2              =   2310
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   4
         X1              =   12630
         X2              =   15240
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   10470
         X2              =   14400
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   7
         X1              =   3150
         X2              =   11340
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.OLE HP_OLE 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1965
         Left            =   13680
         OleObjectBlob   =   "Aux_Steam_Alone.frx":137908
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\My Documents\visiotbne.vsd"
         TabIndex        =   9
         Top             =   3360
         Width           =   1620
      End
      Begin VB.Line LINE_875 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   13740
         X2              =   13740
         Y1              =   3720
         Y2              =   3900
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   1
         X1              =   480
         X2              =   1380
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   1
         X1              =   7470
         X2              =   7470
         Y1              =   2820
         Y2              =   3690
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   7
         X1              =   7470
         X2              =   12120
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Steam_pipe 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   12480
         X2              =   12840
         Y1              =   4050
         Y2              =   4050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 # SH Steam"
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
         Left            =   9960
         TabIndex        =   141
         Top             =   3450
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "875 # Desup Steam"
         Height          =   255
         Index           =   1
         Left            =   2430
         TabIndex        =   140
         Top             =   3450
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Label BYPASS_875_150 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   33
         Top             =   5670
         Width           =   555
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00A0A0A0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   8
         Left            =   1950
         Top             =   8820
         Width           =   1275
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00A0A0A0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00A0A0A0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   9
         Left            =   3450
         Top             =   8820
         Width           =   225
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "150 #"
         Height          =   255
         Left            =   4440
         TabIndex        =   36
         Top             =   5190
         Width           =   525
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   2775
         Index           =   0
         Left            =   11280
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1620
      TabIndex        =   188
      Top             =   90
      Width           =   435
   End
End
Attribute VB_Name = "Aux_Steam_Alone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Feed_Water_Alone.ZOrder


End Sub

Private Sub Command10_Click()

Feed_Water_Alone.ZOrder

End Sub

Private Sub Command11_Click()

Feed_Water_Alone.ZOrder

End Sub

Private Sub Command12_Click()

Main_Condenser.ZOrder

End Sub

Private Sub Command13_Click()

Condensate.ZOrder

End Sub

Private Sub Command14_Click()

Feed_Water_Alone.ZOrder

End Sub

Private Sub Command2_Click()

Stm_Stm_Generator.ZOrder

End Sub

Private Sub Command3_Click()

If LSS_reals(80) > 25# Then
  Gland_Steam_Full_away.Show
Else
  Gland_Steam.Show
End If

End Sub

Private Sub Command4_Click()


' 13/07/09
'Main_Turbines.zoreder
Main_Turbines.ZOrder


End Sub

Private Sub Command5_Click()

Condensate.ZOrder

End Sub

Private Sub Command6_Click()

Distil.ZOrder

End Sub

Private Sub Command7_Click()

Stbd_Boiler.ZOrder

End Sub

Private Sub Command8_Click()

Port_Boiler.ZOrder

End Sub

Private Sub Command9_Click()

' 16/07/09
'Burner_FO_system.ZOrder
Burner_FO_system_NEW.ZOrder

End Sub

Private Sub Drum_water_level_port_Click()

    Unload Alarm_set_point   ' Do this first
    
    nAlarm_HI = 52
    nAlarm_LO = 60  ' 6/1/00

    Alarm_set_point.Show

End Sub

Private Sub Drum_water_level_stbd_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 5
    
    nAlarm_LO = 13  ' 6/1/00
    
    Alarm_set_point.Show

End Sub

Private Sub Form_Load()



' 14/07/09 Removed so that objects which do NOT exist can be detected....
'On Error Resume Next

' 18/06/09
'Call JYAScreenInit(Me, idFeed_water)
Call JYAScreenInit(Me, idAux_Steam_Alone)


'====== From Main Turbine code....====================
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    'For i = 0 To 6
    For i = 0 To 3
        Steam_pipe(i).Visible = False
    Next i
        Steam_pipe(6).Visible = False
    valve(9).Visible = False        ' Manual Throttle Ahead valve
    'Valve(10).Visible = False       ' Manual Throttle Astern valve
Else
    'For i = 0 To 6
    For i = 0 To 3
        Steam_pipe(i).Visible = True
    Next i
        Steam_pipe(6).Visible = True
    valve(9).Visible = True        ' Manual Throttle Ahead valve
    'Valve(10).Visible = True       ' Manual Throttle Astern valve
End If
' === 13/12/02 17.062 end ===


' === 25/09/02 17.044 ======
If (TARGET_this_sim = CALIFORNIA) Then
    If I_AM_INSTRUCTOR Then
        'Valve(90).Visible = True            ' Warming through valve
        valve(293).Visible = True            ' Warming through valve
        Warm_line(0).Visible = True
        Warm_line(1).Visible = True
        Warm_line(2).Visible = True
        Warming_label.Visible = True
    Else
        'Valve(90).Visible = False           ' Warming through valve
         valve(293).Visible = False          ' Warming through valve
        Warm_line(0).Visible = False
        Warm_line(1).Visible = False
        Warm_line(2).Visible = False
        Warming_label.Visible = False
    End If
End If
' === 25/09/02 17.044 end ===
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    VV_pos_throttle_ahead.Visible = False        ' 0 to 100% ahead label
    'VV_pos_throttle_astern.Visible = False       ' 0 to 100% astern label
Else
    VV_pos_throttle_ahead.Visible = True        ' 0 to 100% ahead label
    'VV_pos_throttle_astern.Visible = True       ' 0 to 100% astern label
End If
'=================================================

' === 18/11/02 ===
update_VSCROLL(23) = True
update_VSCROLL(25) = True
update_VSCROLL(28) = True


' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' === 27/09/02 14.114 ===
    'RECIRC_label(1).Visible = False
    'RECIRC_label(2).Visible = False
    'OPEN_VV(1).Visible = False
    'OPEN_VV(2).Visible = False
    'CLOSE_VV(1).Visible = False
    'CLOSE_VV(2).Visible = False
    'AUTO_VV(1).Visible = False
    'AUTO_VV(2).Visible = False
    ' =======================
    
        ' === 10/04/02 S Green e-mail 30/3/02 point 2b ===
        'REMOTE_feed_control(1).Visible = False
       ' REMOTE_feed_control(2).Visible = False
    
    
     ' 11/12/00
     BYPASS_IP.Visible = False
     BYPASS_875_35.Visible = False
     BYPASS_HP.Visible = False
     BYPASS_atom.Visible = False
     BYPASS_875_150.Visible = False
     BYPASS_35_dump.Visible = False
     
     'BYPASS_feed_1.Visible = False
     'BYPASS_feed_2.Visible = False
   
   
    ' 11/12/00
    HScroll1.Visible = False
    
    
    ' === 18/9/00  Point 8.126 ===
    
    ' === 01/03/01 Point 10.2b NOT the Elec Pump ===
    'ONpump(0).Visible = False
    'OFFpump(0).Visible = False
    'sbyPUMP(0).Visible = False

    'ONpump(3).Visible = False
    'OFFpump(3).Visible = False
    'sbyPUMP(3).Visible = False
    
    'ONpump(4).Visible = False
    'OFFpump(4).Visible = False
    'sbyPUMP(4).Visible = False
    
    'AUX_LO_label.Visible = False
    'AUX_1.Visible = False
    'AUX_2.Visible = False
    
    ' 7/9/00
    'TLO_main_feed(1).Visible = False
    'TLO_main_feed(2).Visible = False
    
    ' 10/12/00
    'PLO_total_feed_pump(1).Visible = False
    'PLO_total_feed_pump(2).Visible = False
    
    ' 4/5/00
    SP_HP.Visible = False
    SP_875_150.Visible = False
    SP_IP.Visible = False
    SP_875_35.Visible = False
    
    
    ' === 21/9/00 ===
         VV_pos_875_35_reg.Visible = False
        VV_pos_875_150_reg.Visible = False
             VV_pos_IP_reg.Visible = False
             VV_pos_HP_reg.Visible = False
           VV_pos_dump_reg.Visible = False
        
        'VV_pos_feed(1).Visible = False
        'VV_pos_feed(2).Visible = False
        
        'VV_feed_aux_pos(1).Visible = False
        'VV_feed_aux_pos(2).Visible = False
        
    
Else

    ' === 27/09/02 14.114 ===
    'RECIRC_label(1).Visible = True
    'RECIRC_label(2).Visible = True
    'OPEN_VV(1).Visible = True
    'OPEN_VV(2).Visible = True
    'CLOSE_VV(1).Visible = True
    'CLOSE_VV(2).Visible = True
    'AUTO_VV(1).Visible = True
    'AUTO_VV(2).Visible = True
    ' =======================

        ' === 10/04/02 S Green e-mail 30/3/02 point 2b ===
        'REMOTE_feed_control(1).Visible = True
        'REMOTE_feed_control(2).Visible = True

     ' 11/12/00
     BYPASS_IP.Visible = True
     BYPASS_875_35.Visible = True
     BYPASS_HP.Visible = True
     BYPASS_atom.Visible = True
     BYPASS_875_150.Visible = True
     BYPASS_35_dump.Visible = True
     
     'BYPASS_feed_1.Visible = True
     'BYPASS_feed_2.Visible = True



    ' 11/12/00
    HScroll1.Visible = True

    ' === 18/9/00  Point 8.126 ===
    'ONpump(0).Visible = True
    'OFFpump(0).Visible = True
    'sbyPUMP(0).Visible = True


    'ONpump(3).Visible = True
    'OFFpump(3).Visible = True
    'sbyPUMP(3).Visible = True
    
    'ONpump(4).Visible = True
    'OFFpump(4).Visible = True
    'sbyPUMP(4).Visible = True
    
    'AUX_LO_label.Visible = True
    'AUX_1.Visible = True
    'AUX_2.Visible = True
    
    ' 7/9/00
    'TLO_main_feed(1).Visible = True
    'TLO_main_feed(2).Visible = True

    ' 10/12/00
    'PLO_total_feed_pump(1).Visible = True
    'PLO_total_feed_pump(2).Visible = True

    
    ' 4/5/00
    SP_HP.Visible = True
    SP_875_150.Visible = True
    SP_IP.Visible = True
    SP_875_35.Visible = True
    
    
    ' === 21/9/00 ===
         VV_pos_875_35_reg.Visible = True
        VV_pos_875_150_reg.Visible = True
             VV_pos_IP_reg.Visible = True
             VV_pos_HP_reg.Visible = True
           VV_pos_dump_reg.Visible = True
        
        'VV_pos_feed(1).Visible = True
        'VV_pos_feed(2).Visible = True
        
        'VV_feed_aux_pos(1).Visible = True
        'VV_feed_aux_pos(2).Visible = True


End If


' 19/06/09 REMOVED...(Note: Vscroll1 is now Hscroll_HP_Bypass)
'VScroll1.Visible = PTT
' 14/07/09 ADDED...
Hscroll_HP_Bypass.Visible = PTT


HScroll2.Visible = PTT
HScroll3.Visible = PTT

' 14/07/09 removed
'HScroll5.Visible = PTT
'HScroll4.Visible = PTT
'HScroll9.Visible = PTT
'HScroll7.Visible = PTT
'HScroll6.Visible = PTT

'update_VBAI(6) = True


' === 13/11/02 ===
update_VBAI(6) = True
update_VBAI(7) = True
update_VBAI(8) = True

' === 13/11/02 ===
update_VBAI(12) = True
update_VBAI(13) = True
update_VBAI(14) = True
update_VBAI(15) = True
update_VBAI(16) = True



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then

    HavenStretch1.Visible = False
    
    
    Call JYAScreenResize(Me, idAux_Steam_Alone)
    
    
    HavenStretch1.Visible = True
    
    
    
   
End If

End Sub

Private Sub Hscroll_HP_Bypass_Change()

Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  HP_bypass

' 1/12/99    Call set_VSCROLL(24, VScroll1.Value)

' 22/06/09
'AIvalue = VScroll1.Value / 4
AIvalue = Hscroll_HP_Bypass.Value / 4

Call setAI(6, AIvalue)




End Sub

Private Sub HScroll1_Change()


'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  IP_bypass

' 11/12/00
'Call set_VSCROLL(25, HScroll1.Value)
Call setVscroll(25, HScroll1.Value)

End Sub

Private Sub HScroll2_Change()

Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  Eight75_150_bypass

' 1/12/99
'Call set_VSCROLL(21, HScroll2.Value)

' 1/12/99
AIvalue = HScroll2.Value / 4
Call setAI(7, AIvalue)


End Sub

Private Sub HScroll3_Change()

Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  Eight75_35_bypass
' 1/12/99
'Call set_VSCROLL(22, HScroll3.Value)

' 1/12/99
AIvalue = HScroll3.Value / 4
Call setAI(8, AIvalue)

End Sub

Private Sub HScroll8_Change()


'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)

' 6/11/99  Atom_bypass

' 7/12/00
'Call set_VSCROLL(28, HScroll8.Value)
Call setVscroll(28, HScroll8.Value)


End Sub

Private Sub Psteam1_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 4
    nAlarm_LO = 12  ' 6/1/00
    Alarm_set_point.Show

End Sub

Private Sub Psteam101_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 53
    nAlarm_LO = 61  ' 6/1/99
    Alarm_set_point.Show

End Sub

Private Sub SP_875_150_Click()


set_dummy_focus = True

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 257

Alarm_set_point.Show

set_dummy_focus = True

End Sub

Private Sub SP_875_35_Click()

set_dummy_focus = True

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 259

Alarm_set_point.Show

set_dummy_focus = True

End Sub

Private Sub SP_HP_Click()

set_dummy_focus = True

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 256

Alarm_set_point.Show

set_dummy_focus = True

End Sub

Private Sub SP_IP_Click()

set_dummy_focus = True

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 258

Alarm_set_point.Show

set_dummy_focus = True

End Sub

Private Sub Timer1_Timer()


Dim i As Integer
Dim DFT_level As Integer
Dim green_colour As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' === 18/11/02 ===
' === 18/11/02 ==============
'equivalence(STEAM_integers(46),Vscroll_value(1))

If update_VSCROLL(25) Then
   update_VSCROLL(25) = False
   
   ' === 19/02/03 ===
   'HScroll1.Value = VB_steam_integers(70)
   Call update_the_slider(HScroll1, VB_steam_integers(70))
   
End If
If update_VSCROLL(23) Then
   update_VSCROLL(23) = False
   
   ' === 19/02/03 ===
   'VScroll5.Value = VB_steam_integers(68)
   Call update_the_slider(VScroll5, VB_steam_integers(68))
     ' TESTING...
     'i = VScroll5.Value
   
End If
If update_VSCROLL(28) Then
   update_VSCROLL(28) = False
   
   ' === 19/02/03 ===
   'HScroll8.Value = VB_steam_integers(73)
   Call update_the_slider(HScroll8, VB_steam_integers(73))
   
End If


' === 13/11/02 =======
If update_VBAI(6) Then
     update_VBAI(6) = False
     
     ' === 19/02/03 ===
     'VScroll1.Value = VB_AI(6) * 4
     
     ' 13/07/09
     'Call update_the_slider(VScroll1, VB_AI(6) * 4)
     Call update_the_slider(Hscroll_HP_Bypass, VB_AI(6) * 4)
     
     
End If
If update_VBAI(7) Then
     update_VBAI(7) = False
     
     'HScroll2.Value = VB_AI(7) * 4
     Call update_the_slider(HScroll2, VB_AI(7) * 4)
     
End If
If update_VBAI(8) Then
     update_VBAI(8) = False
     
     'HScroll3.Value = VB_AI(8) * 4
     Call update_the_slider(HScroll3, VB_AI(8) * 4)
     
End If


' 22/06/09 Removed...
'If update_VBAI(12) Then
'     update_VBAI(12) = False
'     Call update_the_slider(HScroll5, VB_AI(12) * 4)
'End If
'If update_VBAI(13) Then
'     update_VBAI(13) = False
'     Call update_the_slider(HScroll4, VB_AI(13) * 4)
'End If


'If update_VBAI(14) Then
'     update_VBAI(14) = False
'     Call update_the_slider(HScroll9, VB_AI(14) * 4)
'End If

'If update_VBAI(15) Then
'     update_VBAI(15) = False
'     Call update_the_slider(HScroll7, VB_AI(15) * 4)
'End If

'If update_VBAI(16) Then
'     update_VBAI(16) = False
'     ' === 19/02/03 ===
'     'HScroll6.Value = VB_AI(16) * 4
'     Call update_the_slider(HScroll6, VB_AI(16) * 4)
'     ' TESTING...
'     'i = HScroll6.Value
'End If
' === 13/11/02 END ===


' 17/8/00
' 22/06/09 Removed...
'If VB_steam_logicals(123) Then
'    REMOTE_feed_control(1).BackColor = &HFFFF&          ' Yellow
'Else
'    REMOTE_feed_control(1).BackColor = &HC0C0C0         ' Grey
'End If
'If VB_steam_logicals(124) Then
'    REMOTE_feed_control(2).BackColor = &HFFFF&          ' Yellow
'Else
'    REMOTE_feed_control(2).BackColor = &HC0C0C0         ' Grey
'End If


' 18/9?00 Point 8.126
' 22/06/09 Removed...
'If Not I_AM_INSTRUCTOR And Not PTT Then
'    ' I MUST be an ACT Terminal (Full Mission) so
'    Tripped_lbl(1).Visible = False
'    Tripped_lbl(2).Visible = False
'Else
'  If VB_steam_logicals(125) Then
'    Tripped_lbl(1).Visible = True
'  Else
'    Tripped_lbl(1).Visible = False
'  End If
'  If VB_steam_logicals(126) Then
'    Tripped_lbl(2).Visible = True
'  Else
'    Tripped_lbl(2).Visible = False
'  End If
'End If


'P_main_out(1).text = Format(LSS_reals(344) * Press_mult, "##0.0") & Press_chars
'P_main_out(2).text = Format(LSS_reals(345) * Press_mult, "##0.0") & Press_chars


' 19/7/00....

'If LSS_LOGICALS(61) Then
'    AUTO_feed_pump(1).BackColor = &HFFFF&         '  YELLOW
'Else
'    AUTO_feed_pump(1).BackColor = &HC0C0C0     'GREY
'End If

'If LSS_LOGICALS(62) Then
'    AUTO_feed_pump(2).BackColor = &HFFFF&         '  YELLOW
'Else
'    AUTO_feed_pump(2).BackColor = &HC0C0C0     'GREY
'End If

'   2/5/00
'' LO pump for Feed pump # 1
'Call update_pump_LSS(Me, 3, 6)      ' Feed pump AUX LO  # 1
'' LO pump for Feed pump # 1
'Call update_pump_LSS(Me, 4, 18)     ' Feed pump AUX LO  # 2


' 27/8/00
' 22/06/09 Removed...

'Call paint_alarm(Me.PLO_total_feed_pump(1), 78, 0, 0, 0)
'Call paint_alarm(Me.PLO_total_feed_pump(2), 79, 0, 0, 0)

'Call paint_alarm(Me.P_MF, 75, 0, 0, 0)
'Call paint_alarm(Me.Z_DFT, 76, 81, 0, 0)

' 10/1/00
'    PC_REALS(398) = F_feed_pump(1)
'    PC_REALS(399) = F_feed_pump(2)
'If (LSS_reals(398) > 0.02) Then
'  HavenPump1(1).PumpOn = True
'Else
'  HavenPump1(1).PumpOn = False
'End If
'If (LSS_reals(399) > 0.02) Then
'  HavenPump1(2).PumpOn = True
'Else
'  HavenPump1(2).PumpOn = False
'End If


'  1/12/99
'If Not PTT Then
'
'    If update_VBAI(6) Then
'       update_VBAI(6) = False
'        VScroll1.Value = VB_AI(6) * 4  ' VB_AI(6) = AI(6) = HP bypass
'    End If
'End If



'          PC_REALS(342) = P_MF
'          PC_REALS(343) = P_AF
'          PC_REALS(344) = P_main_out(1)
'          PC_REALS(345) = P_main_out(2)
'          PC_REALS(346) = P_MAIN(1)
'          PC_REALS(347) = P_MAIN(2)
'          PC_REALS(348) = F_MAIN(1)
'          PC_REALS(349) = F_MAIN(2)
'          PC_REALS(350) = F_MF(1)
'          PC_REALS(351) = F_MF(2)
'          PC_REALS(352) = F_MF_recirc(1)
'          PC_REALS(353) = F_MF_recirc(2)
'          PC_REALS(354) = F_MF_aux(1)
'          PC_REALS(355) = F_MF_aux(2)
'          PC_REALS(356) = VV_feed_aux_pos(1) * 100#
'          PC_REALS(357) = VV_feed_aux_pos(2) * 100#
'          PC_REALS(358) = F_main_feed(1)
'          PC_REALS(359) = F_main_feed(2)
'          PC_REALS(360) = F_aux_feed(1)
'          PC_REALS(361) = F_aux_feed(2)

'          PC_REALS(362) =  P_aux_feed_OUT

'          PC_REALS(309) = T_DFT
'          PC_REALS(310) = T_HP_htr
'          PC_REALS(311) = T_DWT
'          PC_REALS(312) = T_ADT

'T_DFT.text = Format(LSS_reals(309) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'T_HP_htr.text = Format(LSS_reals(310) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
    
    
'          PC_REALS(155) = TLO_main_feed(1)
'          PC_REALS(156) = TLO_main_feed(2)
' 7/9/00
'TLO_main_feed(1).text = Format(LSS_reals(155) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'Call paint_alarm(Me.TLO_main_feed(1), 83, 0, 0, 0)
'TLO_main_feed(2).text = Format(LSS_reals(156) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'Call paint_alarm(Me.TLO_main_feed(2), 84, 0, 0, 0)
    
' 12/9/00
Tsteam11.text = Format(LSS_reals(71) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
    

    P_150.Visible = Advanced_mode
    P_875_desup_OUT.Visible = Advanced_mode
    P_HP_reg_IN.Visible = Advanced_mode
    P_IP_reg_IN.Visible = Advanced_mode
    P_35.Visible = Advanced_mode
    P_Atom.Visible = Advanced_mode
    
   ' 16/8/00
   
    'P_main(1).Visible = Advanced_mode
    'P_main(2).Visible = Advanced_mode
    
    'F_main(1).Visible = False
    'F_main(2).Visible = False
    'F_MF(1).Visible = False
    'F_MF(2).Visible = False
    
    'F_MF_recirc(1).Visible = Advanced_mode
    'F_MF_recirc(2).Visible = Advanced_mode
    
    'F_MF_aux(1).Visible = Advanced_mode
    'F_MF_aux(2).Visible = Advanced_mode
    
    'F_main_feed(1).Visible = Advanced_mode
    'F_main_feed(2).Visible = Advanced_mode
    
    'F_aux_feed(1).Visible = Advanced_mode
    'F_aux_feed(2).Visible = Advanced_mode
    

'  === 6/11/99 Regulator and Bypass Valves ===

'    PC_REALS(364) = Eight75_150_reg * 100#
'    PC_REALS(365) = Eight75_150_bypass * 100#
'
'    PC_REALS(366) = Eight75_35_reg * 100#
'    PC_REALS(367) = Eight75_35_bypass * 100#
'
'    PC_REALS(368) = Dump_35_reg * 100#
'    PC_REALS(369) = Dump_35_bypass * 100#
'
'    PC_REALS(370) = HP_reg * 100#
'    PC_REALS(371) = HP_bypass * 100#
'
'    PC_REALS(372) = IP_reg * 100#
'    PC_REALS(373) = IP_bypass * 100#
'
'''' PC_REALS(374) = Feed_pump_reg(1) * 100#
'    PC_REALS(374) =  VV_pos_feed_pump(1) * 100.0

'    PC_REALS(375) = Feed_pump_bypass(1) * 100#
'
'''' PC_REALS(376) = Feed_pump_reg(2) * 100#
'    PC_REALS(376) =  VV_pos_feed_pump(2) * 100.0

'    PC_REALS(377) = Feed_pump_bypass(2) * 100#
'
'    PC_REALS(378) = Atom_reg * 100#
'    PC_REALS(379) = Atom_bypass * 100#

'    PC_REALS(380) = P_150
'    PC_REALS(381) = P_875_desup_OUT
'    PC_REALS(382) = P_HP_reg_IN
'    PC_REALS(383) = P_IP_reg_IN
'    PC_REALS(384) = P_35
'    PC_REALS(385) = P_Atom


VV_pos_875_150_reg.Caption = Format(LSS_reals(364), "##0") & " %"

VV_pos_875_35_reg.Caption = Format(LSS_reals(366), "##0") & " %"

VV_pos_dump_reg.Caption = Format(LSS_reals(368), "##0") & " %"

VV_pos_HP_reg.Caption = Format(LSS_reals(370), "##0") & " %"

VV_pos_IP_reg.Caption = Format(LSS_reals(372), "##0") & " %"


' 22/06/09 REMOVED...
'VV_pos_feed(1).Caption = Format(LSS_reals(374), "##0") & " %"
'VV_pos_feed(2).Caption = Format(LSS_reals(376), "##0") & " %"


'  === 8/11/99 Regulator  Valves ===
Call update_regulator(364, Reg_875_150(0), Reg_875_150(1), Reg_875_150(2), Reg_875_150(3))
Call update_regulator(366, Reg_875_35(0), Reg_875_35(1), Reg_875_35(2), Reg_875_35(3))
Call update_regulator(368, Dump_35(0), Dump_35(1), Dump_35(2), Dump_35(3))


' 22/06/09  The HP regulator is now horizontal...
'Call update_regulator(370, Reg_HP(0), Reg_HP(1), Reg_HP(2), Reg_HP(3))
Call update_regulator(370, Reg_HP_new(0), Reg_HP_new(1), Reg_HP_new(2), Reg_HP_new(3))

Call update_regulator(372, Reg_IP(0), Reg_IP(1), Reg_IP(2), Reg_IP(3))

' 22/06/09
'Call update_regulator(374, Reg_pump_1(0), Reg_pump_1(1), Reg_pump_1(2), Reg_pump_1(3))
'Call update_regulator(376, Reg_pump_2(0), Reg_pump_2(1), Reg_pump_2(2), Reg_pump_2(3))

Call update_regulator(378, Reg_Atomising(0), Reg_Atomising(1), Reg_Atomising(2), Reg_Atomising(3))


' 26/8/00
'       PC_REALS(117) = VV_main_feed_pos(2) * 100#

' 22/06/09
'Call update_green_VV(117, Valve(5))


'  === 6/11/99 Bypass  Valves ===

'    PC_REALS(365) = Eight75_150_bypass * 100#
Call update_green_VV(365, valve(20))
' 11/12/00
BYPASS_875_150.Caption = Format(LSS_reals(365), "##0") & " %"


'    PC_REALS(367) = Eight75_35_bypass * 100#
Call update_green_VV(367, valve(39))
' 11/12/00
BYPASS_875_35.Caption = Format(LSS_reals(367), "##0") & " %"

' PC_REALS(369) = Dump_35_bypass * 100#
Call update_green_VV(369, valve(21))
' 11/12/00
BYPASS_35_dump.Caption = Format(LSS_reals(369), "##0") & " %"


'    PC_REALS(371) = HP_bypass * 100#
Call update_green_VV(371, valve(23))
' 11/12/00
BYPASS_HP.Caption = Format(LSS_reals(371), "##0") & " %"

'    PC_REALS(373) = IP_bypass * 100#
Call update_green_VV(373, valve(43))
' 11/12/00
BYPASS_IP.Caption = Format(LSS_reals(373), "##0") & " %"



'        PC_REALS(375) = Feed_pump_bypass(1) * 100#
'Call update_green_VV(375, Valve(32))
' 11/12/00
'BYPASS_feed_1.Caption = Format(LSS_reals(375), "##0") & " %"

'        PC_REALS(377) = Feed_pump_bypass(2) * 100#
'Call update_green_VV(377, Valve(33))
' 11/12/00
'BYPASS_feed_2.Caption = Format(LSS_reals(377), "##0") & " %"



'        PC_REALS(379) = Atom_bypass * 100#
Call update_green_VV(379, valve(47))
' 11/12/00
BYPASS_atom.Caption = Format(LSS_reals(379), "##0") & " %"


' Main Feed #1 to AUX feed valve...
'VV_feed_aux_pos(1).Caption = Format(LSS_reals(356), "##0") & " %"
'Call update_green_VV(356, Valve(12))

' Main Feed #2 to AUX feed valve...
'VV_feed_aux_pos(2).Caption = Format(LSS_reals(357), "##0") & " %"
'Call update_green_VV(357, Valve(51))




If SI_units Then
    
    'PLO_total_feed_pump(1).text = Format(LSS_reals(486) * Press_mult, "##0.0") & Press_chars
    'PLO_total_feed_pump(2).text = Format(LSS_reals(487) * Press_mult, "##0.0") & Press_chars

    'Z_DFT.text = Format(LSS_reals(451) * 25.4, "###0.0") & " mm"

    'P_aux_feed_out.text = Format(LSS_reals(362) * Press_mult, "##0.0") & Press_chars
  
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0.0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0.0") & Press_chars
    
    'P_MF.text = Format(LSS_reals(342) * Press_mult, "##0.0") & Press_chars
    'P_AF.text = Format(LSS_reals(343) * Press_mult, "##0.0") & Press_chars

  If (Advanced_mode) Then
  
    P_150.text = Format(LSS_reals(380) * Press_mult, "##0.0") & Press_chars
    P_875_desup_OUT.text = Format(LSS_reals(381) * Press_mult, "##0.0") & Press_chars
    P_HP_reg_IN.text = Format(LSS_reals(382) * Press_mult, "##0.0") & Press_chars
    P_IP_reg_IN.text = Format(LSS_reals(383) * Press_mult, "##0.0") & Press_chars
    P_35.text = Format(LSS_reals(384) * Press_mult, "##0.0") & Press_chars
    
    P_Atom.text = Format(LSS_reals(385) * Press_mult, "##0.0") & Press_chars
    
    'P_main(1).text = Format(LSS_reals(346) * Press_mult, "##0.0") & Press_chars
    'P_main(2).text = Format(LSS_reals(347) * Press_mult, "##0.0") & Press_chars
    'F_main(1).text = Format(LSS_reals(348) * Flow_mult, "###0.0") & Flow_chars
    'F_main(2).text = Format(LSS_reals(349) * Flow_mult, "###0.0") & Flow_chars
    
    'F_MF(1).text = Format(LSS_reals(350) * Flow_mult, "###0.0") & Flow_chars
    'F_MF(2).text = Format(LSS_reals(351) * Flow_mult, "###0.0") & Flow_chars
    
    'F_MF_recirc(1).text = Format(LSS_reals(352) * Flow_mult, "###0.0") & Flow_chars
    'F_MF_recirc(2).text = Format(LSS_reals(353) * Flow_mult, "###0.0") & Flow_chars
    
    'F_MF_aux(1).text = Format(LSS_reals(354) * Flow_mult, "###0.0") & Flow_chars
    'F_MF_aux(2).text = Format(LSS_reals(355) * Flow_mult, "###0.0") & Flow_chars
    
    'F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult, "###0.0") & Flow_chars
    'F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult, "###0.0") & Flow_chars
    
    'F_aux_feed(1).text = Format(LSS_reals(360) * Flow_mult, "###0.0") & Flow_chars
    'F_aux_feed(2).text = Format(LSS_reals(361) * Flow_mult, "###0.0") & Flow_chars
  End If
  
Else

    'PLO_total_feed_pump(1).text = Format(LSS_reals(486) * Press_mult, "##0") & Press_chars
    'PLO_total_feed_pump(2).text = Format(LSS_reals(487) * Press_mult, "##0") & Press_chars


    'Z_DFT.text = Format(LSS_reals(451), "###0.0") & " in"
    
    'P_aux_feed_out.text = Format(LSS_reals(362) * Press_mult, "##0") & Press_chars
    
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0") & Press_chars
    
    'P_MF.text = Format(LSS_reals(342) * Press_mult, "##0") & Press_chars
    'P_AF.text = Format(LSS_reals(343) * Press_mult, "##0") & Press_chars
    
  If (Advanced_mode) Then
  
    P_150.text = Format(LSS_reals(380) * Press_mult, "##0") & Press_chars
    P_875_desup_OUT.text = Format(LSS_reals(381) * Press_mult, "##0") & Press_chars
    P_HP_reg_IN.text = Format(LSS_reals(382) * Press_mult, "##0") & Press_chars
    P_IP_reg_IN.text = Format(LSS_reals(383) * Press_mult, "##0") & Press_chars
    P_35.text = Format(LSS_reals(384) * Press_mult, "##0") & Press_chars
    P_Atom.text = Format(LSS_reals(385) * Press_mult, "##0") & Press_chars
    
    'P_main(1).text = Format(LSS_reals(346) * Press_mult, "##0") & Press_chars
    'P_main(2).text = Format(LSS_reals(347) * Press_mult, "##0") & Press_chars
    
    'F_main(1).text = Format(LSS_reals(348) * Flow_mult / 3600#, "###0.0") & " lb/s"
    'F_main(2).text = Format(LSS_reals(349) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    'F_MF(1).text = Format(LSS_reals(350) * Flow_mult / 3600#, "###0.0") & " lb/s"
    'F_MF(2).text = Format(LSS_reals(351) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    'F_MF_recirc(1).text = Format(LSS_reals(352) * Flow_mult / 3600#, "###0.0") & " lb/s"
    'F_MF_recirc(2).text = Format(LSS_reals(353) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    'F_MF_aux(1).text = Format(LSS_reals(354) * Flow_mult / 3600#, "###0.0") & " lb/s"
    'F_MF_aux(2).text = Format(LSS_reals(355) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    'F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult / 3600#, "###0.0") & " lb/s"
    'F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    'F_aux_feed(1).text = Format(LSS_reals(360) * Flow_mult / 3600#, "###0.0") & " lb/s"
    'F_aux_feed(2).text = Format(LSS_reals(361) * Flow_mult / 3600#, "###0.0") & " lb/s"
  End If
  
End If

''  DFT Tank  ( -35 inches to +35 inches )
'DFT_level = (LSS_reals(451) + 35#) * 100# / 70#
'If DFT_level > 99 Then
'  Valve(44).LineWidth = 100
'ElseIf DFT_level < 1 Then
'  Valve(44).LineWidth = 0
'Else
'  Valve(44).LineWidth = DFT_level
'End If


' === 21/11/00  Feed Pump re-circ valves===
 '  Call update_vvalve_LSS(Me, 85, 85, -1)
 '  Call update_vvalve_LSS(Me, 86, 86, -1)

' OPEN re-circ valves
' ==== 24/11/00 ====
'If (D_OUTPUT(145)) Then

'If vv_status_LSS(85) = 3 Then
'  OPEN_VV(1).BackColor = &HFF00&            ' GREEN
'    Valve(54).OpenBtmRight = True
'    Valve(54).OpenTopLeft = True
'Else
'    OPEN_VV(1).BackColor = &HC0C0C0           ' GREY
'End If


' ==== 24/11/00 ====
'If (D_OUTPUT(493)) Then
'If vv_status_LSS(86) = 3 Then
'  OPEN_VV(2).BackColor = &HFF00&            ' GREEN
'    Valve(55).OpenBtmRight = True
'    Valve(55).OpenTopLeft = True
'Else
'  OPEN_VV(2).BackColor = &HC0C0C0           ' GREY
'End If

' CLOSE re-circ valves
' ==== 24/11/00 ====
'If (D_OUTPUT(149)) Then
'If vv_status_LSS(85) = 0 Then
'  CLOSE_VV(1).BackColor = &H8080FF           ' PINK
'    Valve(54).OpenBtmRight = False
'    Valve(54).OpenTopLeft = False
'Else
'    CLOSE_VV(1).BackColor = &HC0C0C0           ' GREY
'End If


' ==== 24/11/00 ====
'If (D_OUTPUT(497)) Then
'If vv_status_LSS(86) = 0 Then
'  CLOSE_VV(2).BackColor = &H8080FF           ' PINK
'    Valve(55).OpenBtmRight = False
'    Valve(55).OpenTopLeft = False
'Else
'    CLOSE_VV(2).BackColor = &HC0C0C0           ' GREY
'End If


' Auto re-circ valves
'If (D_OUTPUT(147)) Then
'  AUTO_VV(1).BackColor = &HFFFF&            ' YELLOW
'Else
'  AUTO_VV(1).BackColor = &HC0C0C0           ' GREY
'End If
'If (D_OUTPUT(495)) Then
'  AUTO_VV(2).BackColor = &HFFFF&            ' YELLOW
'Else
'  AUTO_VV(2).BackColor = &HC0C0C0           ' GREY
'End If


' Note: update_pump_LSS(syxxx,index,nPump)
'Call update_pump_LSS(Me, 0, 9)


' 26/8/00  Valve(5) is a POT

'For i = 0 To 7
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
'For i = 0 To 4
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
 'For i = 6 To 7
'  Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
'For i = 9 To 11
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
'For i = 13 To 19
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i

'   Call update_vvalve_LSS(Me, 70 + 22, 22, -1)
  
'For i = 24 To 27
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
'For i = 29 To 31
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
'For i = 34 To 38
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
'For i = 41 To 42
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i

''For i = 45 To 47
'For i = 45 To 46
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
   
'  6/9/00   Point Hotel Services Valve
'   Call update_vvalve_LSS(Me, 126, 56, -1)


    Call update_vvalve_LSS(Me, 74, 74, -1)
    
    Call update_vvalve_LSS(Me, 76, 76, -1)
    Call update_vvalve_LSS(Me, 77, 77, -1)
    
    Call update_vvalve_LSS(Me, 79, 79, -1)
    
    Call update_vvalve_LSS(Me, 81, 81, -1)
  
    Call update_vvalve_LSS(Me, 88, 88, -1)
    Call update_vvalve_LSS(Me, 89, 89, -1)
    
    Call update_vvalve_LSS(Me, 92, 92, -1)
    
    Call update_vvalve_LSS(Me, 94, 94, -1)
    
For i = 104 To 108
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

    Call update_vvalve_LSS(Me, 111, 111, -1)
    Call update_vvalve_LSS(Me, 112, 112, -1)
    
    Call update_vvalve_LSS(Me, 115, 115, -1)
    Call update_vvalve_LSS(Me, 116, 116, -1)

' Hotel Services Valve
   Call update_vvalve_LSS(Me, 126, 126, -1)
   
   
' Stbd Boiler valves
   Call update_vvalve_LSS(Me, 435, 435, -1)
   Call update_vvalve_LSS(Me, 438, 438, -1)
   Call update_vvalve_LSS(Me, 284, 284, -1)
' Port Boiler valves
   Call update_vvalve_LSS(Me, 465, 465, -1)
   Call update_vvalve_LSS(Me, 468, 468, -1)
   Call update_vvalve_LSS(Me, 283, 283, -1)

' 13/07/09 CDT inlet valve
   Call update_vvalve_LSS(Me, 577, 577, -1)


If SI_units Then

    Psteam1.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    Drum_water_level_stbd.text = Format(LSS_reals(115) * 25.4, "###0") & " mm"
    ' PORT
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    Drum_water_level_port.text = Format(LSS_reals(116) * 25.4, "###0") & " mm"
       
   Psteam6.text = Format(LSS_reals(788) * Press_mult, "##0.0") & Press_chars
    'Psteam10.text = Format(LSS_reals(60) * Press_mult, "##0.0") & Press_chars
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0.0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0.0") & Press_chars
    'Psteam13.text = Format(LSS_reals(63) * Press_mult, "##0.00") & Press_chars
    'Psteam14.text = Format(LSS_reals(64) * Press_mult, "##0.000") & Press_chars
Else
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    Drum_water_level_stbd.text = Format(LSS_reals(115), "##0.0") & " in"
    ' PORT
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    Drum_water_level_port.text = Format(LSS_reals(116), "##0.0") & " in"
   
    Psteam6.text = Format(LSS_reals(788) * Press_mult, "##0") & Press_chars
    'Psteam10.text = Format(LSS_reals(60) * Press_mult, "##0") & Press_chars
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0") & Press_chars
    'Psteam13.text = Format(LSS_reals(63) * Press_mult, "##0.0") & Press_chars
    'Psteam14.text = Format(LSS_reals(64) * Press_mult, "##0.00") & Press_chars
End If


If SI_units Then
    ' STBD
   ' F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult, "###0.0") & Flow_chars
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    Drum_water_level_stbd.text = Format(LSS_reals(115) * 25.4, "###0") & " mm"
    ' PORT
    'F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult, "###0.0") & Flow_chars
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "##0.0") & Press_chars     ' Bar
    Drum_water_level_port.text = Format(LSS_reals(116) * 25.4, "###0") & " mm"

   
Else

     ' STBD
    'F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult / 3600#, "###0.0") & " lb/s"
    Psteam1.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    Drum_water_level_stbd.text = Format(LSS_reals(115), "##0.0") & " in"
    
    ' PORT
    'F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult / 3600#, "###0.0") & " lb/s"
    Psteam101.text = Format(LSS_reals(771) * Press_mult, "###0") & Press_chars       ' psi
    Drum_water_level_port.text = Format(LSS_reals(116), "##0.0") & " in"

End If

Tsteam10.text = Format(LSS_reals(70) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam11.text = Format(LSS_reals(71) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'Tsteam12.text = Format(LSS_reals(72) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'Tsteam13.text = Format(LSS_reals(73) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

Call paint_alarm(Me.Psteam101, 53, 61, 0, 0)
Call paint_alarm(Me.Drum_water_level_port, 52, 60, 0, 0)

Call paint_alarm(Me.Psteam1, 4, 12, 0, 0)
Call paint_alarm(Me.Drum_water_level_stbd, 5, 13, 0, 0)


' 22/06/09 Main Turbine
' PC_REALS(81) = V_ahead_hand_pos ( 0 to 100.0 )
If LSS_reals(81) > 99# Then
    green_colour = 255
ElseIf LSS_reals(81) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(81) * 2.55
End If

' 11/10/99   AHEAD HAND-WHEEL Throttling valve...
valve(9).ColourClosed = RGB(0, green_colour, 0)



'   PC_REALS(2) = V_ahead_throttle_pos   ( 0 to 100.0 )
VV_pos_throttle_ahead.Caption = Format(LSS_reals(2), "##0") & " %"
If LSS_reals(2) > 99# Then
    green_colour = 255
ElseIf LSS_reals(2) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(2) * 2.55
End If
valve(0).ColourClosed = RGB(0, green_colour, 0)
      
' Warming Through Valve
Call update_vvalve_LSS(Me, 293, 293, -1)


' 13/07/09  CDT Regulators
Reg_875_225_pos.Caption = Format(LSS_reals(146), "###0") & " %"
Call update_regulator(146, Reg_875_225(0), Reg_875_225(1), Reg_875_225(2), Reg_875_225(3))

REG_120_pos.Caption = Format(LSS_reals(148), "###0") & " %"
Call update_regulator(148, Reg_120(0), Reg_120(1), Reg_120(2), Reg_120(3))


End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
  
 ' 7/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then
   ' I MUST be an ACT Terminal (Full Mission) so
   ' 22/06/09
    ' If Index = 4 Or Index = 9 Or Index = 24 Then
   If Index = 74 Or Index = 79 Or Index = 94 Then
        Exit Sub
    End If
End If
  
If Not PTT Then    ' NON-Clickable Valves...
   ' === 21/03/02 Point 14.502 ===
   ' === 10/06/02 PUT BACK ORIGINAL LINE... ===
    ' 22/06/09
    'If Index = 13 Or Index = 14 Or Index = 30 Or Index = 31 Then
    If Index = 83 Or Index = 84 Or Index = 100 Or Index = 101 Then
        Exit Sub
    End If
End If
If Index = 5 Then
        Exit Sub
End If


' === 22/06/09 Checked Stbd Boiler and Port Boiler ===
' Do not need extra code here...
            
    ' 22/06/09
    'Call valve_pcdi_LSS(Index + 70)
    Call valve_pcdi_LSS(Index)



End Sub

Private Sub VScroll1_Change()

Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  HP_bypass

' 1/12/99    Call set_VSCROLL(24, VScroll1.Value)

' 1/12/99
AIvalue = VScroll1.Value / 4
Call setAI(6, AIvalue)


End Sub

Private Sub VScroll5_Change()

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)


' 6/11/99  Dump_35_bypass
Call set_VSCROLL(23, VScroll5.Value)


End Sub

Sub update_the_slider(Slider As Object, Pot_Value As Integer)

Dim S_max As Integer
Dim S_min As Integer

' TESTING...
S_max = Slider.max
S_min = Slider.Min

' === 19/02/03 ===
'VScroll1.Value = VB_AI(6) * 4
If Pot_Value >= Slider.max Then
    Slider.Value = Slider.max
ElseIf Pot_Value <= Slider.Min Then
    Slider.Value = Slider.Min
Else
    Slider.Value = Pot_Value
End If


End Sub
Sub update_regulator(nReals As Integer, obj_0 As Object, obj_1 As Object, obj_2 As Object, obj_3 As Object)

' 6/11/99
If LSS_reals(nReals) > 99# Then
      If (obj_3.Visible = False) Then
          obj_0.Visible = False
          obj_1.Visible = False
          obj_2.Visible = False
          obj_3.Visible = True
      End If
ElseIf LSS_reals(nReals) > 49# Then
      If (obj_2.Visible = False) Then
          obj_0.Visible = False
          obj_1.Visible = False
          obj_2.Visible = True
          obj_3.Visible = False
      End If
ElseIf LSS_reals(nReals) > 0.5 Then
      If (obj_1.Visible = False) Then
          obj_0.Visible = False
          obj_1.Visible = True
          obj_2.Visible = False
          obj_3.Visible = False
      End If
Else
      If (obj_0.Visible = False) Then
          obj_0.Visible = True
          obj_1.Visible = False
          obj_2.Visible = False
          obj_3.Visible = False
      End If
End If



End Sub
Sub update_green_VV(nReals As Integer, obj As Object)
Dim green_colour As Integer

' 6/11/99
If LSS_reals(nReals) > 99# Then
    green_colour = 255
ElseIf LSS_reals(nReals) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(nReals) * 2.55
End If

obj.ColourClosed = RGB(0, green_colour, 0)

End Sub
