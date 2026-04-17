VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Stbd_Boiler 
   Caption         =   "Stbd Boiler"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   Icon            =   "Stbd_Boiler.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   11925
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9015
      Left            =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   15901
      Begin VB.TextBox F_MAIN_stm 
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
         Height          =   285
         Left            =   9900
         MousePointer    =   1  'Arrow
         TabIndex        =   71
         Text            =   "xx kg/s"
         Top             =   5820
         Width           =   855
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
         Height          =   285
         Left            =   8580
         MousePointer    =   1  'Arrow
         TabIndex        =   70
         Text            =   "xx "
         Top             =   8580
         Visible         =   0   'False
         Width           =   915
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
         Height          =   285
         Left            =   8580
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Text            =   "xx "
         Top             =   8280
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox T_econ_OUT 
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
         Index           =   1
         Left            =   6780
         TabIndex        =   68
         Text            =   "xxx  F"
         Top             =   60
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   -660
         TabIndex        =   67
         Top             =   3780
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox F_875_desup_boiler 
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
         Height          =   285
         Left            =   10560
         MousePointer    =   1  'Arrow
         TabIndex        =   66
         Text            =   "xx kg/s"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox P_875_desup_boiler 
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
         Left            =   9360
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Text            =   "xx Bar"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox F_SSTG_stm 
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
         Height          =   285
         Left            =   9720
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Text            =   "xx kg/s"
         Top             =   7320
         Width           =   915
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   9720
         Max             =   1000
         TabIndex        =   63
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox T_main_feed 
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
         Index           =   1
         Left            =   8580
         TabIndex        =   62
         Text            =   "xxx  F"
         Top             =   1320
         Width           =   675
      End
      Begin VB.TextBox F_main_feed 
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
         Index           =   1
         Left            =   10440
         TabIndex        =   61
         Text            =   "xx.x kg/s"
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox T_econ_IN 
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
         Index           =   1
         Left            =   8280
         TabIndex        =   60
         Text            =   "xxx  F"
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox P_875_desup 
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
         Height          =   285
         Left            =   10560
         MousePointer    =   1  'Arrow
         TabIndex        =   59
         Text            =   "xx Bar"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SSTG"
         Height          =   435
         Left            =   10560
         TabIndex        =   47
         Top             =   7980
         Width           =   555
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   2700
         Top             =   7950
      End
      Begin VB.TextBox Fsteam1 
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
         Height          =   285
         Left            =   8040
         MousePointer    =   1  'Arrow
         TabIndex        =   31
         Text            =   "xx kg/s"
         Top             =   4140
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Psteam5 
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
         Left            =   9900
         MousePointer    =   1  'Arrow
         TabIndex        =   30
         Text            =   "xx Bar"
         Top             =   5520
         Width           =   855
      End
      Begin VB.TextBox Tsteam5 
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
         Left            =   8040
         MousePointer    =   1  'Arrow
         TabIndex        =   29
         Text            =   "xx C"
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox Tsteam1 
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
         Height          =   285
         Left            =   8160
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         Text            =   "xx C"
         Top             =   2520
         Width           =   795
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
         Left            =   4920
         MousePointer    =   1  'Arrow
         TabIndex        =   27
         Text            =   "xx Bar"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Drum_water_level 
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
         Left            =   4920
         MousePointer    =   1  'Arrow
         TabIndex        =   26
         Text            =   "-xx in"
         Top             =   2760
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   449
         Left            =   10080
         TabIndex        =   0
         Top             =   1380
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
         Height          =   375
         Index           =   430
         Left            =   10080
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Index           =   448
         Left            =   9060
         TabIndex        =   2
         Top             =   780
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
         Index           =   450
         Left            =   10080
         TabIndex        =   3
         Top             =   780
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
         Height          =   435
         Index           =   435
         Left            =   10200
         TabIndex        =   5
         Top             =   2100
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
         Index           =   443
         Left            =   4080
         TabIndex        =   6
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   442
         Left            =   5000
         TabIndex        =   7
         Top             =   960
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   446
         Left            =   5540
         TabIndex        =   8
         Top             =   960
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   445
         Index           =   441
         Left            =   4260
         TabIndex        =   9
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   794
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
         Index           =   437
         Left            =   4860
         TabIndex        =   10
         Top             =   7260
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   433
         Left            =   4500
         TabIndex        =   11
         Top             =   7260
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   432
         Left            =   840
         TabIndex        =   12
         Top             =   7980
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
         Index           =   431
         Left            =   6540
         TabIndex        =   13
         Top             =   5640
         Visible         =   0   'False
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
         ValveType       =   11
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   438
         Left            =   10440
         TabIndex        =   14
         Top             =   6120
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
         Index           =   436
         Left            =   7800
         TabIndex        =   15
         Top             =   780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   439
         Left            =   7800
         TabIndex        =   16
         Top             =   1620
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12648447
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   615
         Index           =   0
         Left            =   5220
         TabIndex        =   23
         Top             =   2100
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   1085
         TopColour       =   16711680
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   495
         Index           =   434
         Left            =   9900
         TabIndex        =   33
         Top             =   7860
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
         Height          =   375
         Index           =   440
         Left            =   8160
         TabIndex        =   35
         Top             =   1320
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
         Index           =   445
         Left            =   8820
         TabIndex        =   37
         Top             =   7680
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
         Index           =   451
         Left            =   8820
         TabIndex        =   39
         Top             =   5220
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   444
         Left            =   9510
         TabIndex        =   41
         Top             =   5220
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   447
         Left            =   9420
         TabIndex        =   44
         Top             =   1400
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
         Index           =   452
         Left            =   9720
         TabIndex        =   49
         Top             =   780
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
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   330
         Top             =   720
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   330
         Top             =   450
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   330
         Top             =   1020
         Width           =   195
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Removed"
         Height          =   225
         Index           =   13
         Left            =   660
         TabIndex        =   75
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "DO"
         Height          =   225
         Index           =   12
         Left            =   660
         TabIndex        =   74
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "HO"
         Height          =   225
         Index           =   11
         Left            =   660
         TabIndex        =   73
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner Tip"
         Height          =   225
         Index           =   10
         Left            =   330
         TabIndex        =   72
         Top             =   180
         Width           =   1335
      End
      Begin VB.Shape Shape4 
         Height          =   1155
         Left            =   210
         Top             =   150
         Width           =   1455
      End
      Begin VB.Line Line6 
         X1              =   8460
         X2              =   8460
         Y1              =   5400
         Y2              =   4620
      End
      Begin VB.Line Line5 
         X1              =   9600
         X2              =   9840
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Surface Blow Valve"
         Height          =   375
         Index           =   6
         Left            =   2700
         TabIndex        =   20
         Top             =   1980
         Width           =   1155
      End
      Begin VB.Shape Burner2_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   3300
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   195
      End
      Begin VB.Shape Burner1_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   2940
         Shape           =   3  'Circle
         Top             =   6000
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   3060
         Top             =   4620
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   2760
         Top             =   6240
         Width           =   195
      End
      Begin VB.Line Line4 
         X1              =   8580
         X2              =   8580
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line Advanced_line 
         X1              =   8820
         X2              =   8280
         Y1              =   1620
         Y2              =   2160
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   10740
         Picture         =   "Stbd_Boiler.frx":0CFA
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   255
      End
      Begin VB.Line Line2 
         X1              =   10200
         X2              =   10200
         Y1              =   6060
         Y2              =   6420
      End
      Begin VB.Label V_feed_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   9360
         TabIndex        =   55
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Feed Water Control"
         Height          =   255
         Index           =   9
         Left            =   8520
         TabIndex        =   54
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Feed"
         Height          =   255
         Left            =   10440
         TabIndex        =   46
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Feed"
         Height          =   255
         Left            =   10440
         TabIndex        =   45
         Top             =   1440
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         Height          =   1275
         Left            =   4630
         Shape           =   3  'Circle
         Top             =   1980
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   9600
         X2              =   10800
         Y1              =   8220
         Y2              =   8220
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Vent"
         Height          =   315
         Index           =   0
         Left            =   9420
         TabIndex        =   42
         Top             =   4800
         Width           =   375
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   9600
         X2              =   9600
         Y1              =   5040
         Y2              =   8220
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Safety"
         Height          =   315
         Index           =   1
         Left            =   8640
         TabIndex        =   40
         Top             =   4800
         Width           =   555
      End
      Begin VB.Label Vent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Drain"
         Height          =   255
         Index           =   2
         Left            =   7800
         TabIndex        =   38
         Top             =   7740
         Width           =   795
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   8940
         X2              =   8940
         Y1              =   5040
         Y2              =   8220
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   9300
         X2              =   10860
         Y1              =   6440
         Y2              =   6440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Stbd Boiler SSTG Stop"
         Height          =   555
         Left            =   9540
         TabIndex        =   34
         Top             =   8400
         Width           =   1095
      End
      Begin VB.Label V_attemp_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "xx %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   32
         Top             =   4860
         Width           =   555
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   2640
         Shape           =   2  'Oval
         Top             =   6540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape16 
         Height          =   1065
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   5790
         Width           =   1155
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 1 Burner"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         Top             =   6900
         Width           =   750
      End
      Begin VB.Shape Ign_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   3000
         Shape           =   2  'Oval
         Top             =   4890
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape8 
         Height          =   1065
         Left            =   2610
         Shape           =   3  'Circle
         Top             =   4140
         Width           =   1155
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "# 2 Burner"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   5280
         Width           =   750
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Attemperator Control"
         Height          =   495
         Index           =   8
         Left            =   7080
         TabIndex        =   22
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Blow Down Valves"
         Height          =   495
         Index           =   7
         Left            =   4020
         TabIndex        =   21
         Top             =   8100
         Width           =   1155
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Air Cock"
         Height          =   315
         Index           =   5
         Left            =   3660
         TabIndex        =   19
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Steam Stop valve"
         Height          =   495
         Index           =   4
         Left            =   9840
         TabIndex        =   18
         Top             =   6600
         Width           =   1155
      End
      Begin VB.Label Vent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeSuperheated Stop valve"
         Height          =   495
         Index           =   3
         Left            =   9120
         TabIndex        =   17
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   10140
         TabIndex        =   4
         Top             =   540
         Width           =   255
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   6900
         OleObjectBlob   =   "Stbd_Boiler.frx":1004
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   50
         Top             =   5160
         Width           =   375
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   6900
         OleObjectBlob   =   "Stbd_Boiler.frx":981C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   53
         Top             =   5160
         Width           =   375
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   6900
         OleObjectBlob   =   "Stbd_Boiler.frx":11634
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   52
         Top             =   5160
         Width           =   375
      End
      Begin VB.OLE Attemp_Reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   6900
         OleObjectBlob   =   "Stbd_Boiler.frx":1964C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   51
         Top             =   5160
         Width           =   375
      End
      Begin VB.OLE Port_Boiler_OLE 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   8115
         Left            =   780
         OleObjectBlob   =   "Stbd_Boiler.frx":21664
         SizeMode        =   1  'Stretch
         SourceDoc       =   "\\Pc01\PC01_C\STEAM_SIM\Images\Stbd_boiler_WV1.vsd"
         TabIndex        =   43
         Top             =   660
         Width           =   10290
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   5655
         Left            =   1860
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   6135
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2235
         Left            =   6060
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   1935
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   9420
         OleObjectBlob   =   "Stbd_Boiler.frx":3087C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   48
         Top             =   720
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   9420
         OleObjectBlob   =   "Stbd_Boiler.frx":38E94
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   58
         Top             =   720
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   9420
         OleObjectBlob   =   "Stbd_Boiler.frx":40CAC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   57
         Top             =   720
         Width           =   315
      End
      Begin VB.OLE Feed_reg 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   9420
         OleObjectBlob   =   "Stbd_Boiler.frx":48EC4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   56
         Top             =   720
         Width           =   315
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3660
      TabIndex        =   36
      Top             =   300
      Width           =   435
   End
End
Attribute VB_Name = "Stbd_Boiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Compressor_Updated(Code As Integer)

End Sub

Private Sub Command1_Click()

SSTG.Show

End Sub

Private Sub Drum_water_level_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 5
    
    nAlarm_LO = 13  ' 6/1/00
    
    Alarm_set_point.Show

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idsy156)



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

set_dummy_focus = True


' === 13/11/02 =======
update_VBAI(4) = True


' 5/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        F_875_desup_boiler.Visible = False
        
        '21/9/00
        V_attemp_pos(1).Visible = False

Else
        F_875_desup_boiler.Visible = True
        
        '21/9/00
        V_attemp_pos(1).Visible = True

End If



' 1/12/99
HScroll2.Visible = PTT

' 7/1/00
Feed_reg(3).ZOrder
Feed_reg(2).ZOrder
Feed_reg(1).ZOrder
Feed_reg(0).ZOrder



End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy156)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub HScroll2_Change()

Dim AIvalue As Integer

' 5/1/00   Stbd Aux Feed Valve(449)
' VALVE is now variable ( used to be clickable )

AIvalue = HScroll2.Value / 4
Call setAI(4, AIvalue)



End Sub

Private Sub P_875_desup_boiler_Click()

' 5/5/00 Added:...
    Unload Alarm_set_point   ' Do this first
    nAlarm_LO = 15

    Alarm_set_point.Show


End Sub

Private Sub P_875_desup_Click()


' 5/5/00 Moved...
'  ' 6/1/00
'    Unload Alarm_set_point   ' Do this first
'    nAlarm_LO = 15
'
'    Alarm_set_point.Show

End Sub

Private Sub Psteam1_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 4
    nAlarm_LO = 12  ' 6/1/00
    Alarm_set_point.Show

End Sub

Private Sub Psteam5_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 6
    nAlarm_LO = 14
    Alarm_set_point.Show

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim B_water_level As Integer
Dim VV_pos As Integer




If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


'c        === 19/11/02 ===
'        PC_REALS(881) = F_SSTG_stm(1)
'        PC_REALS(882) = F_SSTG_stm(2)
'        PC_REALS(883) = QAVAILABLE(1) * 0.001
'        PC_REALS(884) = QAVAILABLE(2) * 0.001
'Text1.text = Format(LSS_reals(881), "##0.00")
'Text2.text = Format(LSS_reals(883), "####0")



' === 13/11/02 =======
If update_VBAI(4) Then
     update_VBAI(4) = False
     
     ' === 26/03/03 ===
    ' HScroll2.Value = VB_AI(4) * 4
    Call update_the_slider(HScroll2, VB_AI(4) * 4)

End If


' 5/5/00
List1.clear
List1.AddItem "F_feed_pump(1) = " & LSS_reals(398)
List1.AddItem "F_feed_pump(2) = " & LSS_reals(399)
List1.AddItem "F_875_desup_out = " & LSS_reals(496)
List1.AddItem "F_atom = " & LSS_reals(497)
List1.AddItem "F_875_150 = " & LSS_reals(498)
List1.AddItem "F_875_35 = " & LSS_reals(499)
List1.AddItem "Tsteam(7) = " & LSS_reals(500)


' 10/5/00 Do NOT need these...
'Call paint_alarm(Me.Psteam1, 4, 0, 0, 0)
'Call paint_alarm(Me.Psteam5, 6, 14, 0, 0)
'Call paint_alarm(Me.Drum_water_level, 5, 0, 0, 0)
'Call paint_alarm(Me.Tsteam5, 7, 0, 0, 0)

' 7/1/00
Call paint_alarm(Me.Psteam1, 4, 12, 0, 0)
Call paint_alarm(Me.Psteam5, 6, 14, 0, 0)
Call paint_alarm(Me.Drum_water_level, 5, 13, 0, 0)
Call paint_alarm(Me.Tsteam5, 7, 0, 0, 0)


' 5/5/00
'Call paint_alarm(Me.P_875_desup, 15, 0, 0, 0)
Call paint_alarm(Me.P_875_desup_boiler, 15, 0, 0, 0)


' 26/11/99   HO/DO Nozzles
'equivalence(STEAM_logicals(15),HO_nozzle(1))

For i = 0 To 1

    ' 12/10/09 REMOVED
    'If VB_steam_logicals(15 + i) Then
    '    Burner_tip(i).FillColor = &H4080&         ' DARK BROWN
    'Else
    '    Burner_tip(i).FillColor = &H80FF&         ' ORANGE
    'End If
    
    ' 12/10/09 ADDED
    If LSS_reals(877 + i) > 1.5 Then
        'HO Nozzle fited
         Burner_tip(i).FillColor = &H4080&         ' DARK BROWN
    ElseIf LSS_reals(877 + i) > 0.5 Then
        'DO Nozzle fited
        Burner_tip(i).FillColor = &H80FF&         ' ORANGE
    Else
         Burner_tip(i).FillColor = vbWhite        ' white
    End If

    
    
Next i

'        equivalence(STEAM_logicals(1),Burner1_flame(1))
'        equivalence(STEAM_logicals(3),Burner2_flame(1))
'        logical*1 Burner1_flame(2)
'        logical*1 Burner2_flame(2)
    If VB_steam_logicals(1) Then
        Burner1_shape(1).FillColor = RGB(&HFF, 0, 0)
    Else
        Burner1_shape(1).FillColor = RGB(&HC0, &HC0, &HC0)
    End If
    If VB_steam_logicals(3) Then
        Burner2_shape(1).FillColor = RGB(&HFF, 0, 0)
    Else
        Burner2_shape(1).FillColor = RGB(&HC0, &HC0, &HC0)
    End If




'          PC_REALS(340) = VV_pos_drum(1) * 100.0
'          PC_REALS(341) = VV_pos_drum(2) * 100.0

V_feed_pos(1).Caption = Format(LSS_reals(340), "###0") & " %"
If LSS_reals(340) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(340) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(340) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Feed_reg(i).Visible = True
    Else
        Feed_reg(i).Visible = False
    End If
Next i

V_attemp_pos(1).Caption = Format(LSS_reals(786), "###0") & " %"
If LSS_reals(786) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(786) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(786) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Attemp_Reg(i).Visible = True
    Else
        Attemp_Reg(i).Visible = False
    End If
Next i

    ' 17/11/99
    F_main_feed(1).Visible = Advanced_mode
    
'        PC_REALS(386) = T_econ_IN(1)
'        PC_REALS(387) = T_econ_IN(2)
'        PC_REALS(388) = T_econ_OUT(1)
'        PC_REALS(389) = T_econ_OUT(2)
'        PC_REALS(390) = T_main_feed(1)
'        PC_REALS(391) = T_main_feed(2)

'cc 4 / 4 / 0
'    PC_REALS(127) = Psteam(5)
'    PC_REALS(128) = Psteam(105)


T_econ_IN(1).text = Format(LSS_reals(386) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

T_econ_OUT(1).text = Format(LSS_reals(388) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


T_main_feed(1).text = Format(LSS_reals(390) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

'  OV015
'          PC_REALS(363) = P_875_desup_OUT

'   3/9/99
Tsteam1.text = Format(LSS_reals(755) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Tsteam5.text = Format(LSS_reals(767) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

If SI_units Then

    ' 17/11/99
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult, "###0.0") & Flow_chars

    Psteam1.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    ' 4/5/00
    P_875_desup_boiler.text = Format(LSS_reals(492) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    ' 4/4/00
  ' Psteam5.text = Format(LSS_reals(754) * Press_mult, "##0.0") & Press_chars     ' Bar
    Psteam5.text = Format(LSS_reals(127) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    Fsteam1.text = Format(LSS_reals(758) * Flow_mult, "###0.0") & Flow_chars       ' kg/s
    
    ' === 20/11/02 ===
    '    PC_REALS(490) = F_MAIN_stm(1)
    '    PC_REALS(491) = F_MAIN_stm(2)
    F_MAIN_stm.text = Format(LSS_reals(490) * Flow_mult, "###0.0") & Flow_chars       ' kg/s

   
   ' 5/5/00
    F_875_desup_boiler.text = Format(LSS_reals(494) * Flow_mult, "###0.0") & Flow_chars
    
    'c        === 20/11/02 ===
    '        PC_REALS(881) = F_SSTG_stm(1)
    '        PC_REALS(882) = F_SSTG_stm(2)

   'F_SSTG_boiler.text = Format(LSS_reals(490) * Flow_mult, "###0.00") & Flow_chars       ' kg/s
    F_SSTG_stm.text = Format(LSS_reals(881) * Flow_mult, "###0.00") & Flow_chars       ' kg/s
    
    Drum_water_level.text = Format(LSS_reals(115) * 25.4, "###0") & " mm"

    P_875_desup.text = Format(LSS_reals(363) * Press_mult, "##0.0") & Press_chars     ' Bar

Else
     ' 17/11/99
   F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult / 3600#, "###0.0") & " lb/s"

    Psteam1.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    ' 4/5/00
    P_875_desup_boiler.text = Format(LSS_reals(492) * Press_mult, "###0") & Press_chars       ' psi
    
    ' 4/4/00
   'Psteam5.text = Format(LSS_reals(754) * Press_mult, "###0") & Press_chars       ' psi
    Psteam5.text = Format(LSS_reals(127) * Press_mult, "###0") & Press_chars       ' psi
    
    Fsteam1.text = Format(LSS_reals(758) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    ' === 20/11/02 ===
    F_MAIN_stm.text = Format(LSS_reals(490) * Flow_mult / 3600#, "###0.0") & " lb/s"
     
     
     
    ' 5/5/00
     F_875_desup_boiler.text = Format(LSS_reals(494) * Flow_mult / 3600#, "###0.0") & " lb/s"
     
    'c        === 20/11/02 ===
    '        PC_REALS(881) = F_SSTG_stm(1)
    'F_SSTG_boiler.text = Format(LSS_reals(490) * Flow_mult / 3600#, "###0.00") & " lb/s"
     F_SSTG_stm.text = Format(LSS_reals(881) * Flow_mult / 3600#, "###0.00") & " lb/s"
   
    Drum_water_level.text = Format(LSS_reals(115), "##0.0") & " in"

    P_875_desup.text = Format(LSS_reals(363) * Press_mult, "###0") & Press_chars       ' psi
End If

' Drum Water Level ( Tank )
If LSS_reals(115) > 8.9999 Then
    B_water_level = 100         ' FULL
ElseIf LSS_reals(115) < -8.9999 Then
    B_water_level = 0           ' EMPTY
Else
    B_water_level = Int((LSS_reals(115) + 9#) * 100# / 18#)
End If

' 30/11/99 VB ( or OCX bug )..
' when going from 20% to 0% the Blue water level does not change so..
If B_water_level < 1 Then
   B_water_level = 1
End If

    valve(0).LineWidth = B_water_level

' 28/10/99
'For i = 432 To 451
For i = 432 To 452

  ' 5/1/00
   'If i = 442 Or i = 446 Or i = 451 Then
    If i = 442 Or i = 446 Or i = 451 Or i = 449 Then
      ' Do NOT update safety valves here ( see code below )
    Else
        Call update_vvalve_LSS(Me, i, i, -1)  ' Update Digital Valve Status
    End If
    
Next i


'c 6 / 1 / 0
'    PC_REALS(483) = VV_aux_feed_pos(1)
'    PC_REALS(484) = VV_aux_feed_pos(2)

' 5/1/00
Call update_green_VV(483, valve(449))



'       80,81     DRUM_Safety_VV_1(1-2)
'       89,90     DRUM_Safety_VV_2(1-2)
'       91,92     SUPER_Safety_VV(1-2)



'  DRUM Safety Valve # 1
If VB_steam_logicals(80) Then
    valve(442).OpenBtmRight = True
    valve(442).OpenBtmRight = True
Else
    valve(442).OpenBtmRight = False
    valve(442).OpenBtmRight = False
End If

'  DRUM Safety Valve # 2
If VB_steam_logicals(89) Then
    valve(446).OpenBtmRight = True
    valve(446).OpenBtmRight = True
Else
    valve(446).OpenBtmRight = False
    valve(446).OpenBtmRight = False
End If
     
'  SUPER-HEAT Safety Valve
If VB_steam_logicals(91) Then
    valve(451).OpenBtmRight = True
    valve(451).OpenBtmRight = True
Else
    valve(451).OpenBtmRight = False
    valve(451).OpenBtmRight = False
End If
 


End Sub

Private Sub Tsteam5_Click()

    Unload Alarm_set_point   ' Do this first
    nAlarm_HI = 7
    Alarm_set_point.Show

End Sub

Private Sub valve_Click(Index As Integer)


 ' 14/8/00 Block aux feed valve click on ACT Terminals

' === 23/09/02 Point 17 Block 448 and 450 on ACT Terminals as well ===
'If Index = 447 Then
If Index = 447 Or Index = 448 Or Index = 452 Then


  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
  
  
End If
  



    If Index = 442 Or Index = 446 Or Index = 451 Then
      ' Do NOT click safety valves
    Else
            
          Call valve_pcdi_LSS(Index)
          
    End If

End Sub
Sub update_green_VV(nReals As Integer, obj As Object)
Dim green_colour As Integer

' 5/1/00
If LSS_reals(nReals) > 99# Then
    green_colour = 255
ElseIf LSS_reals(nReals) < 1# Then
    green_colour = 0
Else
    green_colour = LSS_reals(nReals) * 2.55
End If

obj.ColourClosed = RGB(0, green_colour, 0)

End Sub
Sub update_the_slider(Slider As Object, Pot_Value As Integer)

Dim S_max As Integer
Dim S_min As Integer

' === 26/03/03 (ADDED) ===

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


