VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Bilge_OWS 
   Caption         =   "Bilge system"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12510
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   12510
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8925
      Left            =   0
      Top             =   0
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   15743
      Begin VB.TextBox pbs52 
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
         Left            =   2400
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox oil_ppm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Left            =   6360
         MousePointer    =   1  'Arrow
         TabIndex        =   60
         Top             =   3900
         Width           =   735
      End
      Begin VB.TextBox zbs_oil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Left            =   2940
         MousePointer    =   1  'Arrow
         TabIndex        =   59
         Top             =   5700
         Width           =   615
      End
      Begin VB.TextBox pbs53 
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
         TabIndex        =   36
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
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
         Index           =   1
         Left            =   7410
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1980
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
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
         Left            =   2370
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4800
         UseMaskColor    =   -1  'True
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
         Index           =   1
         Left            =   6960
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1980
         UseMaskColor    =   -1  'True
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
         Left            =   1920
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4800
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
         Index           =   1
         Left            =   6510
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1980
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
         Left            =   1470
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4800
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.TextBox zbs36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   7050
         Width           =   675
      End
      Begin VB.TextBox zbs35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox zbs40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         Left            =   9840
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Top             =   7020
         Width           =   675
      End
      Begin VB.TextBox zbs39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
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
         Left            =   9780
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Top             =   540
         Width           =   675
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3150
         Width           =   255
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   2130
      End
      Begin VB.TextBox zbs_hold 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Left            =   8340
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Top             =   5700
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   530
         Left            =   3660
         TabIndex        =   25
         Top             =   840
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   714
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   531
         Left            =   5580
         TabIndex        =   26
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   536
         Left            =   9750
         TabIndex        =   27
         Top             =   4110
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   873
         BottomColour    =   255
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Bilge_oil_Tank 
         Height          =   495
         Left            =   3750
         TabIndex        =   20
         Top             =   5520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         TopColour       =   65535
         LineWidth       =   20
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Bilge_Holding_Tank 
         Height          =   615
         Left            =   7080
         TabIndex        =   28
         Top             =   5370
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         TopColour       =   32768
         LineWidth       =   20
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   9
         ColourOpen      =   65535
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   495
         Index           =   533
         Left            =   6240
         TabIndex        =   39
         Top             =   2280
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   873
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   7080
         TabIndex        =   29
         Top             =   2360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         LineCol         =   32768
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   535
         Left            =   8340
         TabIndex        =   45
         Top             =   2940
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   529
         Left            =   3660
         TabIndex        =   46
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   532
         Left            =   5940
         TabIndex        =   47
         Top             =   1500
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   405
         Index           =   534
         Left            =   8370
         TabIndex        =   48
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   537
         Left            =   720
         TabIndex        =   62
         Top             =   3480
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   1980
         TabIndex        =   63
         Top             =   4150
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrLineWidth  =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   526
         Left            =   1320
         TabIndex        =   64
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   527
         Left            =   9180
         TabIndex        =   65
         Top             =   480
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   538
         Left            =   1440
         TabIndex        =   66
         Top             =   4140
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   539
         Left            =   2520
         TabIndex        =   67
         Top             =   3840
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   525
         Left            =   1500
         TabIndex        =   68
         Top             =   6960
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   541
         Left            =   5880
         TabIndex        =   69
         Top             =   7440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   5
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   528
         Left            =   8880
         TabIndex        =   70
         Top             =   6960
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   540
         Left            =   6300
         TabIndex        =   71
         Top             =   5100
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
         ValveType       =   5
      End
      Begin VB.Line OWS_line 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   6480
         X2              =   6480
         Y1              =   4440
         Y2              =   5100
      End
      Begin VB.Line OWS_line 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   6480
         X2              =   5580
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line OWS_line 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   6420
         X2              =   6420
         Y1              =   7020
         Y2              =   7620
      End
      Begin VB.Line OWS_line 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   5580
         X2              =   6420
         Y1              =   7020
         Y2              =   7020
      End
      Begin VB.Line OWS_line 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   5580
         X2              =   5580
         Y1              =   6840
         Y2              =   7020
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1695
         Left            =   3030
         OleObjectBlob   =   "Bilge.frx":0000
         SizeMode        =   1  'Stretch
         TabIndex        =   44
         Top             =   3060
         Width           =   1035
      End
      Begin VB.Line Line21 
         X1              =   3000
         X2              =   3000
         Y1              =   3420
         Y2              =   4200
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   6240
         X2              =   6420
         Y1              =   7620
         Y2              =   7620
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5880
         Picture         =   "Bilge.frx":6618
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   255
      End
      Begin VB.Line Hold_line 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Index           =   0
         X1              =   6000
         X2              =   6000
         Y1              =   6960
         Y2              =   6840
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   2280
         TabIndex        =   73
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   1920
         TabIndex        =   72
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   7980
         Picture         =   "Bilge.frx":6922
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   255
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   11040
         X2              =   30
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Image Image28 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2910
         Picture         =   "Bilge.frx":6C2C
         Stretch         =   -1  'True
         Top             =   1210
         Width           =   255
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4470
         Picture         =   "Bilge.frx":6F36
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   255
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1080
         Picture         =   "Bilge.frx":7240
         Stretch         =   -1  'True
         Top             =   4770
         Width           =   255
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1080
         Picture         =   "Bilge.frx":754A
         Stretch         =   -1  'True
         Top             =   2970
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   5415
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   8295
      End
      Begin VB.OLE OLE25 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   2940
         OleObjectBlob   =   "Bilge.frx":7854
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   52
         Top             =   6960
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   7500
         X2              =   8460
         Y1              =   2390
         Y2              =   2390
      End
      Begin VB.OLE OLE37 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   3240
         OleObjectBlob   =   "Bilge.frx":F46C
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   55
         Top             =   2040
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   1200
         X2              =   2220
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "STARBOARD SIDE"
         Height          =   285
         Left            =   4500
         TabIndex        =   58
         Top             =   8070
         Width           =   3885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PORT SIDE"
         Height          =   195
         Left            =   5280
         TabIndex        =   57
         Top             =   60
         Width           =   1095
      End
      Begin VB.OLE OLE39 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   3720
         OleObjectBlob   =   "Bilge.frx":17084
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   56
         Top             =   2040
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE OLE36 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   7620
         OleObjectBlob   =   "Bilge.frx":1EC9C
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   54
         Top             =   4500
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE OLE35 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   7680
         OleObjectBlob   =   "Bilge.frx":268B4
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   53
         Top             =   6960
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OLE OLE34 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   10020
         OleObjectBlob   =   "Bilge.frx":2E4CC
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   51
         Top             =   2100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   11010
         X2              =   0
         Y1              =   7980
         Y2              =   7980
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         Picture         =   "Bilge.frx":360E4
         Stretch         =   -1  'True
         Top             =   3660
         Width           =   255
      End
      Begin VB.OLE OLE33 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Left            =   4440
         OleObjectBlob   =   "Bilge.frx":363EE
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisvlveMLBR.vsd"
         TabIndex        =   50
         Top             =   7260
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   2040
         X2              =   2040
         Y1              =   6720
         Y2              =   7290
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   555
         Left            =   2040
         OleObjectBlob   =   "Bilge.frx":3E006
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisVlveMUBR.vsd"
         TabIndex        =   49
         Top             =   1500
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line OWS_line 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   5580
         X2              =   5580
         Y1              =   4260
         Y2              =   6540
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " S "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5880
         TabIndex        =   3
         Top             =   3870
         Width           =   240
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   6000
         X2              =   6000
         Y1              =   3510
         Y2              =   5430
      End
      Begin VB.Line Line50 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3810
         X2              =   4200
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line49 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   4230
         X2              =   4230
         Y1              =   3090
         Y2              =   3240
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   9510
         Picture         =   "Bilge.frx":45C1E
         Stretch         =   -1  'True
         Top             =   7140
         Width           =   255
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   9510
         Picture         =   "Bilge.frx":45F28
         Stretch         =   -1  'True
         Top             =   660
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   5700
         X2              =   5700
         Y1              =   1440
         Y2              =   2580
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   1680
         TabIndex        =   43
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   315
         Left            =   1080
         TabIndex        =   42
         Top             =   6660
         Width           =   315
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   8970
         TabIndex        =   41
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   285
         Left            =   9240
         TabIndex        =   40
         Top             =   6720
         Width           =   315
      End
      Begin VB.Image Image26 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8340
         Picture         =   "Bilge.frx":46232
         Stretch         =   -1  'True
         Top             =   1740
         Width           =   255
      End
      Begin VB.Line Line55 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   8460
         X2              =   8460
         Y1              =   3240
         Y2              =   3780
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   6060
         X2              =   6060
         Y1              =   1350
         Y2              =   2580
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   2340
         X2              =   3120
         Y1              =   4190
         Y2              =   4190
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ER PORT FOR'D (OIL BILGE)"
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
         Left            =   9360
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Top             =   900
         Width           =   1140
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AFT PEAK"
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
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   38
         Top             =   3930
         Width           =   705
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FORE PEAK"
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
         Left            =   9750
         MousePointer    =   1  'Arrow
         TabIndex        =   37
         Top             =   4650
         Width           =   810
      End
      Begin VB.Line Hold_line 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Index           =   5
         X1              =   6000
         X2              =   6000
         Y1              =   7980
         Y2              =   7140
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00000000&
         X1              =   8160
         X2              =   8460
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00000000&
         X1              =   3390
         X2              =   3750
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Image Image8 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4110
         Picture         =   "Bilge.frx":4653C
         Stretch         =   -1  'True
         Top             =   4530
         Width           =   255
      End
      Begin VB.Image Image16 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5880
         Picture         =   "Bilge.frx":46846
         Stretch         =   -1  'True
         Top             =   7110
         Width           =   255
      End
      Begin VB.Image Image29 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6750
         Picture         =   "Bilge.frx":46B50
         Stretch         =   -1  'True
         Top             =   5310
         Width           =   255
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   810
         Picture         =   "Bilge.frx":46E5A
         Stretch         =   -1  'True
         Top             =   7170
         Width           =   405
      End
      Begin VB.Image Image9 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   930
         Picture         =   "Bilge.frx":47164
         Stretch         =   -1  'True
         Top             =   660
         Width           =   255
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   900
         X2              =   2040
         Y1              =   7290
         Y2              =   7290
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   1170
         X2              =   1920
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   8730
         X2              =   8730
         Y1              =   6720
         Y2              =   7230
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   8730
         X2              =   9900
         Y1              =   7260
         Y2              =   7260
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   9000
         X2              =   9000
         Y1              =   780
         Y2              =   1350
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   1920
         X2              =   1920
         Y1              =   810
         Y2              =   1320
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   9030
         X2              =   9780
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SW SUCTION"
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
         Left            =   6120
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Top             =   390
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   90
         X2              =   1170
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   3780
         X2              =   3780
         Y1              =   330
         Y2              =   1320
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ER STBD FOR'D"
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
         Left            =   9000
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   7740
         Width           =   1080
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ER PORT AFT"
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
         Left            =   300
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ER STBD AFT"
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
         Left            =   420
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   7500
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BILGE HOLDING TANK"
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
         Left            =   6840
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   6060
         Width           =   1725
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BILGE OIL TANK"
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
         Left            =   3570
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   6060
         Width           =   1305
      End
      Begin VB.Image Image15 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8880
         Picture         =   "Bilge.frx":4746E
         Stretch         =   -1  'True
         Top             =   3690
         Width           =   255
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SW SUCTION"
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
         Left            =   2790
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BILGE PUMP"
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
         Left            =   1380
         TabIndex        =   10
         Top             =   3900
         Width           =   1170
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GS PUMP"
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
         Left            =   6900
         TabIndex        =   9
         Top             =   1650
         Width           =   660
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   5700
         X2              =   7140
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   5700
         X2              =   5700
         Y1              =   690
         Y2              =   1230
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   8460
         X2              =   8460
         Y1              =   360
         Y2              =   1230
      End
      Begin VB.Line Line20 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   4230
         X2              =   4230
         Y1              =   3510
         Y2              =   5490
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   3930
         X2              =   4530
         Y1              =   3390
         Y2              =   3390
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   4410
         X2              =   4770
         Y1              =   3750
         Y2              =   4110
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4410
         Top             =   4110
         Width           =   375
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   4650
         X2              =   5010
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   4890
         X2              =   5250
         Y1              =   3750
         Y2              =   4110
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4890
         Top             =   4110
         Width           =   375
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   5160
         X2              =   6000
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   7110
         X2              =   6030
         Y1              =   5430
         Y2              =   5430
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   6000
         X2              =   6000
         Y1              =   5430
         Y2              =   6600
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   9510
         X2              =   10440
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BILGE OVERBOARD DISCHARGE"
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
         Height          =   495
         Left            =   6480
         TabIndex        =   8
         Top             =   7380
         Width           =   990
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   8460
         X2              =   8460
         Y1              =   2340
         Y2              =   2940
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SEPARATOR"
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
         Left            =   3150
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   5010
         Width           =   900
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COELESCENT"
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
         Left            =   4410
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   4500
         Width           =   945
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FILTERS"
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
         Left            =   4560
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   4680
         Width           =   570
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OILY WATER"
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
         Left            =   3150
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   4830
         Width           =   840
      End
      Begin VB.Shape oil_ppm_shape 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5520
         Shape           =   3  'Circle
         Top             =   3960
         Width           =   195
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00000000&
         X1              =   5790
         X2              =   5910
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OIL / WATER MONITOR ppm"
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
         Height          =   600
         Left            =   6180
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   3390
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line47 
         BorderColor     =   &H00000000&
         Visible         =   0   'False
         X1              =   4860
         X2              =   4860
         Y1              =   2760
         Y2              =   3120
      End
      Begin VB.Line Line48 
         BorderColor     =   &H00000000&
         X1              =   6120
         X2              =   6360
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FIRE MAIN"
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
         Left            =   8550
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   3540
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         X1              =   8460
         X2              =   8460
         Y1              =   1440
         Y2              =   2340
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   5460
         Top             =   3900
         Width           =   315
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4410
         Top             =   3750
         Width           =   375
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4890
         Top             =   3750
         Width           =   375
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   5130
         X2              =   5130
         Y1              =   3510
         Y2              =   3750
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   4650
         X2              =   4650
         Y1              =   3510
         Y2              =   3750
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   5010
         X2              =   5010
         Y1              =   3510
         Y2              =   3750
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         X1              =   4530
         X2              =   4530
         Y1              =   3390
         Y2              =   3750
      End
      Begin VB.Line Line56 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   8460
         X2              =   9000
         Y1              =   3810
         Y2              =   3810
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1020
      TabIndex        =   74
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Bilge_OWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim set_dummy_focus As Integer

Private Sub Command3D10_Click()

' 1/9/99

'Dim i As Integer
'
'For i = 0 To 40
'  Black_line(i).BorderColor = &HFF&         ' RED
'Next i'


End Sub

Private Sub Command3D2_Click()


End Sub

Private Sub Command3D3_Click()


End Sub

Private Sub Command3D9_Click()


' 4/4/00
'sy137_LSS.Show
Fire_Main.Show

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idsy131)


' === 20/11/02 17.014 ===
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      ONpump(0).Visible = False
      OFFpump(0).Visible = False
      sbyPUMP(0).Visible = False
End If
 



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy131)
    HavenStretch1.Visible = True
End If
End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
  If Index = 0 Then
      Exit Sub        ' USE the MECD controls for the bilge pump
  End If
End If
   
   
' 12/7/00
 
'   Note:  PumpIO_LSS(4=OFF,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 11))    ' Bilge Pump OFF
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 37))    ' GS Pump OFF
End If


End Sub

Private Sub oil_ppm_Click()

' 13/7/98
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 275
Alarm_set_point.Show

End Sub

Private Sub ONpump_Click(Index As Integer)

    
    
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
' 12/7/00

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
  If Index = 0 Then
      Exit Sub        ' USE the MECD controls for the bilge pump
  End If
End If

 
'   Note:  PumpIO_LSS(2=ON,nPump)
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 11))    ' Bilge Pump ON
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 37))    ' GS Pump ON
End If

    
    
    
End Sub

Private Sub sbyPUMP_Click(Index As Integer)


set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   

If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If

'   NEW Steam_Sim Pump numbers
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 11))    ' Bilge pump AUTO
End If



End Sub

Private Sub Timer1_Timer()
Dim i As Integer



If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 1/9/99 ZBS_EMERG.Text = Format(LSS_reals(329), "##0") & " %"


oil_ppm.text = Format(LSS_reals(339), "##0") & " ppm"
Call paint_alarm(Me.oil_ppm, 275, 0, 0, 0)


pbs52.text = Format(LSS_reals(335), "#0.0") & " bar"
'pbs53.Text = Format(LSS_reals(?), "#0.0") & " bar"
'pbs62.Text = Format(LSS_reals(330), "#0.0") & " bar"


' 11/9/00
' ==== alarms 260,264,266,270 changed to 114 to 117 ===

' ER Port Forward Bilge Ring Main
zbs39.text = Format(LSS_reals(333), "##0") & " %"
If VB_steam_logicals(133) Then
   zbs39.BackColor = &HFF&                   ' RED  ( HIGH-HIGH )
   zbs39.ForeColor = &HFFFFFF                ' WHITE
ElseIf alarm(114) Then
   zbs39.BackColor = &H80FF&                 '  ORANGE ( HIGH )
   zbs39.ForeColor = &H0&                    ' BLACK
Else
   zbs39.BackColor = &HC000&                 ' DULL GREEN ( NORMAL )
   zbs39.ForeColor = &H0&                    ' BLACK
End If

' ER Starboard Forward Bilge Ring Main
zbs40.text = Format(LSS_reals(334), "##0") & " %"
If VB_steam_logicals(135) Then
    zbs40.BackColor = &HFF&                   ' RED  ( HIGH-HIGH )
    zbs40.ForeColor = &HFFFFFF                ' WHITE
ElseIf alarm(116) Then
    zbs40.BackColor = &H80FF&                 '  ORANGE ( HIGH )
    zbs40.ForeColor = &H0&                    ' BLACK
Else
    zbs40.BackColor = &HC000&                 ' DULL GREEN ( NORMAL )
    zbs40.ForeColor = &H0&                    ' BLACK
End If

' Port Aft Bilge Tank
zbs35.text = Format(LSS_reals(331), "##0") & " %"
If VB_steam_logicals(134) Then
   zbs35.BackColor = &HFF&                  ' RED  ( HIGH-HIGH )
   zbs35.ForeColor = &HFFFFFF               ' WHITE
ElseIf alarm(115) Then
   zbs35.BackColor = &H80FF&                '  ORANGE ( HIGH )
   zbs35.ForeColor = &H0&                   ' BLACK
Else
   zbs35.BackColor = &HC000&                ' DULL GREEN ( NORMAL )
   zbs35.ForeColor = &H0&                   ' BLACK
End If

' Starboard Aft Bilge Ring Main
zbs36.text = Format(LSS_reals(332), "##0") & " %"
If VB_steam_logicals(136) Then
   zbs36.BackColor = &HFF&                   ' RED  ( HIGH-HIGH )
   zbs36.ForeColor = &HFFFFFF                ' WHITE
ElseIf alarm(117) Then
   zbs36.BackColor = &H80FF&                 '  ORANGE ( HIGH )
   zbs36.ForeColor = &H0&                    ' BLACK
Else
   zbs36.BackColor = &HC000&                 ' DULL GREEN ( NORMAL )
   zbs36.ForeColor = &H0&                    ' BLACK
End If



' Bilge Oil Tank
zbs_oil.text = Format(LSS_reals(337), "##0") & " %"
Bilge_oil_Tank.LineWidth = LSS_reals(337)
Call paint_alarm(Me.zbs_oil, 276, 0, 0, 0)


' Bilge Holding Tank
zbs_hold.text = Format(LSS_reals(338), "##0") & " %"
Bilge_Holding_Tank.LineWidth = LSS_reals(338)
Call paint_alarm(Me.zbs_hold, 277, 0, 0, 0)


' 13/7/00  Oil ppm HIGH
Call paint_alarm(Me.oil_ppm, 275, 0, 0, 0)
If alarm(275) Then
  oil_ppm_shape.FillColor = &HFF&            ' Red
Else
  oil_ppm_shape.FillColor = &HFF00&       ' Green
End If


' PUMPS
' Note: update_pump_LSS(syxxx,index,nPump)

' 12/7/00
Call update_pump_LSS(Me, 0, 11)      ' Bilge Pump
Call update_pump_LSS(Me, 1, 37)      ' Genral Service Pump

' 13/7/00
If pp_status_LSS(11) = 1 Or pp_status_LSS(11) = 2 Then
  ' Bilge pump is in AUTO
  For i = 0 To 5
    OWS_line(i).Visible = True
  Next i
Else
  ' Bilge pump is in NOT in AUTO
  For i = 0 To 5
    OWS_line(i).Visible = False
  Next i
End If



' 12/7/00
For i = 525 To 541
   Call update_vvalve_LSS(Me, i, i, -1)
Next i


End Sub

Private Sub valve_Click(Index As Integer)


 ' 12/7/00 For CMA
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
            
    If Index = 538 Or Index = 539 Then  ' Must be a REAL Valve ( on the MECD )
        Exit Sub
    End If
    
End If

    Call valve_pcdi_LSS(Index)


End Sub

Private Sub zbs_hold_Click()

' 13/7/98
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 277
Alarm_set_point.Show

End Sub

Private Sub zbs_oil_Click()

' 13/7/98
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 276
Alarm_set_point.Show

End Sub

Private Sub zbs35_Click()

Unload Alarm_set_point   ' Do this first

' 11/9/00
'nAlarm_HI = 264
nAlarm_HI = 115
nAlarm_HIHI = 265
Alarm_set_point.Show

End Sub

Private Sub zbs36_Click()

Unload Alarm_set_point   ' Do this first
' 11/9/00
'nAlarm_HI = 270
nAlarm_HI = 117
nAlarm_HIHI = 271
Alarm_set_point.Show

End Sub

Private Sub zbs37_Click()

'Unload Alarm_set_point   ' Do this first
'nAlarm_HI = 248
'nAlarm_HIHI = 249
'Alarm_set_point.Show

End Sub

Private Sub zbs38_Click()

'Unload Alarm_set_point   ' Do this first
'nAlarm_HI = 250
'nAlarm_HIHI = 251
'Alarm_set_point.Show

End Sub

Private Sub zbs39_Click()

Unload Alarm_set_point   ' Do this first
' 11/9/00
'nAlarm_HI = 260
nAlarm_HI = 114
nAlarm_HIHI = 261
Alarm_set_point.Show

End Sub

Private Sub zbs40_Click()

Unload Alarm_set_point   ' Do this first
' 11/9/00
'nAlarm_HI = 266
nAlarm_HI = 116
nAlarm_HIHI = 267
Alarm_set_point.Show

End Sub
