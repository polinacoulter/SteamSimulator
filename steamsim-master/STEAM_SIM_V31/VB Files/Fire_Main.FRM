VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Fire_Main 
   Caption         =   "FIRE MAIN SYSTEM"
   ClientHeight    =   12735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   14040
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   10095
      Left            =   60
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   17806
      Begin VB.CommandButton Command1 
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6300
         Width           =   255
      End
      Begin VB.CommandButton Command3D9 
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
         Left            =   6180
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   -360
         Width           =   255
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/B"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   7380
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   7680
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ON"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   7440
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   3720
         TabIndex        =   3
         Top             =   7980
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   16384
         CentrLineWidth  =   2
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   9660
         Top             =   7800
      End
      Begin VB.TextBox pfire3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Top             =   8280
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   490
         Left            =   5160
         TabIndex        =   1
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   491
         Left            =   4800
         TabIndex        =   9
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   492
         Left            =   2520
         TabIndex        =   10
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   494
         Left            =   3840
         TabIndex        =   12
         Top             =   2100
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   495
         Left            =   4440
         TabIndex        =   13
         Top             =   2100
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   496
         Left            =   5160
         TabIndex        =   14
         Top             =   1860
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   493
         Left            =   3480
         TabIndex        =   11
         Top             =   1980
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   497
         Left            =   7920
         TabIndex        =   15
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   498
         Left            =   11640
         TabIndex        =   16
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   499
         Left            =   1500
         TabIndex        =   17
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   500
         Left            =   3780
         TabIndex        =   18
         Top             =   2940
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   501
         Left            =   4980
         TabIndex        =   19
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   502
         Left            =   9540
         TabIndex        =   20
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   504
         Left            =   3840
         TabIndex        =   21
         Top             =   3780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   505
         Left            =   4440
         TabIndex        =   22
         Top             =   3780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   506
         Left            =   5340
         TabIndex        =   23
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   507
         Left            =   7140
         TabIndex        =   24
         Top             =   4200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   509
         Left            =   5520
         TabIndex        =   25
         Top             =   4680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   508
         Left            =   4620
         TabIndex        =   26
         Top             =   4860
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   503
         Left            =   2520
         TabIndex        =   27
         Top             =   4260
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   512
         Left            =   4620
         TabIndex        =   28
         Top             =   5280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   514
         Left            =   7740
         TabIndex        =   29
         Top             =   5280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   510
         Left            =   2100
         TabIndex        =   30
         Top             =   6120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   511
         Left            =   3660
         TabIndex        =   31
         Top             =   5880
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   515
         Left            =   3780
         TabIndex        =   32
         Top             =   6600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   513
         Left            =   6720
         TabIndex        =   33
         Top             =   5520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   516
         Left            =   5160
         TabIndex        =   34
         Top             =   6780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   517
         Left            =   5880
         TabIndex        =   35
         Top             =   6780
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   518
         Left            =   6840
         TabIndex        =   36
         Top             =   6540
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   520
         Left            =   1620
         TabIndex        =   37
         Top             =   7920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   521
         Left            =   2820
         TabIndex        =   38
         Top             =   7920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   519
         Left            =   8340
         TabIndex        =   39
         Top             =   6360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   522
         Left            =   5580
         TabIndex        =   40
         Top             =   8460
         Visible         =   0   'False
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   523
         Left            =   6000
         TabIndex        =   41
         Top             =   8400
         Visible         =   0   'False
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   255
         Index           =   524
         Left            =   6660
         TabIndex        =   42
         Top             =   8400
         Visible         =   0   'False
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   575
         Left            =   4320
         TabIndex        =   44
         Top             =   7680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   315
         Index           =   524
         Left            =   6540
         TabIndex        =   48
         Top             =   7320
         Width           =   255
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   315
         Index           =   523
         Left            =   5880
         TabIndex        =   47
         Top             =   7320
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   8820
         TabIndex        =   46
         Top             =   8040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line2 
         X1              =   4860
         X2              =   4860
         Y1              =   7980
         Y2              =   8280
      End
      Begin VB.Shape Green_shape 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7680
         Shape           =   3  'Circle
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape RED_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7260
         Shape           =   3  'Circle
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   315
         Index           =   522
         Left            =   5280
         TabIndex        =   43
         Top             =   7320
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   524
         Left            =   6540
         Shape           =   3  'Circle
         Top             =   7860
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   523
         Left            =   5940
         Shape           =   3  'Circle
         Top             =   7860
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   522
         Left            =   5280
         Shape           =   3  'Circle
         Top             =   7860
         Width           =   255
      End
      Begin VB.OLE Port_Boiler_OLE 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   10080
         Left            =   240
         OleObjectBlob   =   "Fire_Main.frx":0000
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Fire_main WV2.vsd"
         TabIndex        =   8
         Top             =   420
         Width           =   13470
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Fire_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim set_dummy_focus As Integer



Private Sub Command1_Click()

' 12/7/00
Bilge_OWS.Show

End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idsy137)


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy137)
    HavenStretch1.Visible = True
End If
End Sub

Private Sub OFFpump_Click(Index As Integer)


set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
' 11/7/00
 
'   Note:  PumpIO_LSS(4=OFF,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 10))    ' FIRE Pump OFF
End If


End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
' 11/7/00
 
'   Note:  PumpIO_LSS(2=ON,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 10))    ' FIRE Pump ON
End If


End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim valve_number As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 11/7/00
'For i = 490 To 521
For i = 490 To 524
   Call update_vvalve_LSS(Me, i, i, -1)
Next i

' 11/7/00 added valve(0) ( NON-return valve to Stop GS pump flow back )

' 4/9/00
'valve(0).OpenBtmRight = valve(521).OpenBtmRight
'valve(0).OpenTopLeft = valve(521).OpenTopLeft
Call update_vvalve_LSS(Me, 575, 575, -1)

' 13/8/00   8.14
'pfire3.text = Format(LSS_reals(336), "#0.0") & " bar"
If SI_units Then
    pfire3.text = Format(LSS_reals(336) * Press_mult, "##0.0") & Press_chars     ' Bar
Else
    pfire3.text = Format(LSS_reals(336) * Press_mult, "###0") & Press_chars    ' psi
End If


Call update_pump_LSS(Me, 0, 10)      ' FIRE Pump


For i = 522 To 524

     valve_number = i
   
    If vv_status_LSS(valve_number) <> OLD_vv_status_LSS(valve_number) Then
        If vv_status_LSS(valve_number) = 0 Then
            VV_shape(i).FillColor = RED_shape.FillColor
        ElseIf vv_status_LSS(valve_number) = 3 Then
            VV_shape(i).FillColor = Green_shape.FillColor
        End If
    End If
    
    If (vv_status_LSS(valve_number) = 1) Or (vv_status_LSS(valve_number) = 2) Then ' FLASH
            If VV_shape(i).FillColor = Green_shape.FillColor Then
                VV_shape(i).FillColor = RED_shape.FillColor
            Else
                VV_shape(i).FillColor = Green_shape.FillColor
            End If
    End If
    
    OLD_vv_status_LSS(valve_number) = vv_status_LSS(valve_number)
    
Next i

End Sub

Private Sub valve_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
 ' 11/7/00  OLD CODE (LSS)
'If Index = 0 Or Index = 1 Then
'    Call valve_pcdi_LSS(131)    ' BS14
'ElseIf Index > 1 And Index < 24 Then
'    Call valve_pcdi_LSS(275 - Index)  ' Fire valves
'ElseIf Index = 24 Then
'    Call valve_pcdi_LSS(237)    ' To Fire system from GS pump
'ElseIf Index = 25 Then
'    Call valve_pcdi_LSS(275)    ' 22/12/98 Emerg Fire Pump Valve
'End If

   
 ' 11/7/00 For CMA
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
            
  ' 4/9/00
   'If Index > 520 Then     ' Must be a REAL Valve ( on the MECD )
    If Index > 520 And Index <> 575 Then   ' Must be a REAL Valve ( on the MECD )
        Exit Sub
    End If
    
End If


' 4/9/00
'If Index = 0 Then
'    Call valve_pcdi_LSS(521)
'Else
    Call valve_pcdi_LSS(Index)
'End If



End Sub

Private Sub VV_label_Click(Index As Integer)


 ' 11/7/00 For CMA
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
            
    If Index > 520 Then     ' Must be a REAL Valve ( on the MECD )
        Exit Sub
    End If
    
End If

    Call valve_pcdi_LSS(Index)


End Sub
