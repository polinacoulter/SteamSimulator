VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{A6D9C597-48C2-11D1-8915-000021240899}#2.0#0"; "Haven3WayValve v1.41.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form CMA_purifier 
   Caption         =   "LO STORAGE & SEPARATOR SYSTEM"
   ClientHeight    =   8370
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   12150
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   7140
      Left            =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12594
      Begin VB.CommandButton Command1 
         Caption         =   "LO system"
         Height          =   495
         Left            =   240
         TabIndex        =   58
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox PLS02 
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
         Height          =   290
         Left            =   6540
         MousePointer    =   1  'Arrow
         TabIndex        =   57
         Top             =   3960
         Width           =   810
      End
      Begin VB.TextBox fls20 
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
         Height          =   290
         Left            =   2820
         MousePointer    =   1  'Arrow
         TabIndex        =   56
         Top             =   6480
         Width           =   735
      End
      Begin VB.TextBox fls80 
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
         Height          =   290
         Left            =   9420
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         Top             =   3480
         Width           =   795
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
         Left            =   4935
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   5700
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
         Left            =   4890
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AUTO"
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
         Left            =   4335
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin Haven3WayValvePrj.Haven3WayValve Haven3WayValve1 
         Height          =   450
         Left            =   4935
         TabIndex        =   22
         Top             =   3075
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   794
         OpenTopTri      =   0   'False
         OpenSideTri     =   0   'False
         OpenBtmTri      =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         Direction       =   3
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   9660
         Top             =   4920
      End
      Begin VB.TextBox zls03 
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
         Height          =   290
         Left            =   3120
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   1099
         Width           =   675
      End
      Begin VB.TextBox zls80 
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
         Height          =   290
         Left            =   8775
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   810
         Width           =   735
      End
      Begin VB.TextBox fls10 
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
         Height          =   290
         Left            =   7320
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   3176
         Width           =   735
      End
      Begin VB.TextBox ZLO_sump 
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
         Height          =   290
         Left            =   960
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   6600
         Width           =   675
      End
      Begin VB.TextBox tls04 
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
         Height          =   290
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox fls02 
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
         Height          =   290
         Left            =   5760
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   3960
         Width           =   810
      End
      Begin VB.TextBox fls90 
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
         Height          =   290
         Left            =   9600
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   420
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox fls03 
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
         Height          =   290
         Left            =   5520
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox Zlwater 
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
         Height          =   290
         Left            =   975
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   2565
         Width           =   615
      End
      Begin VB.TextBox zsludge 
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
         Height          =   290
         Left            =   6990
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   4695
         Width           =   615
      End
      Begin VB.TextBox fls04 
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
         Height          =   290
         Left            =   1140
         MousePointer    =   1  'Arrow
         TabIndex        =   0
         Top             =   3705
         Visible         =   0   'False
         Width           =   675
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   0
         Left            =   1020
         TabIndex        =   23
         Top             =   930
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   900
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   1
         Left            =   3240
         TabIndex        =   24
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   3
         Left            =   1200
         TabIndex        =   25
         Top             =   1830
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   4
         Left            =   5160
         TabIndex        =   26
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
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
         Height          =   510
         Index           =   5
         Left            =   8520
         TabIndex        =   27
         Top             =   945
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   6
         Left            =   8730
         TabIndex        =   28
         Top             =   255
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
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
         Height          =   510
         Index           =   7
         Left            =   2220
         TabIndex        =   29
         Top             =   6015
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
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
         Height          =   510
         Index           =   8
         Left            =   2700
         TabIndex        =   30
         Top             =   4770
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
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
         Index           =   9
         Left            =   6375
         TabIndex        =   31
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   10
         Left            =   6960
         TabIndex        =   32
         Top             =   3390
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
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
         Index           =   12
         Left            =   3780
         TabIndex        =   34
         Top             =   3660
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   11
         Left            =   3240
         TabIndex        =   33
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   5
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   13
         Left            =   3780
         TabIndex        =   35
         Top             =   4020
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   14
         Left            =   3780
         TabIndex        =   36
         Top             =   4380
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   975
         Index           =   15
         Left            =   1380
         TabIndex        =   37
         Top             =   480
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1720
         TopColour       =   65535
         BottomColour    =   65535
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   975
         Index           =   16
         Left            =   6870
         TabIndex        =   38
         Top             =   465
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1720
         TopColour       =   65535
         BottomColour    =   65535
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   675
         Index           =   17
         Left            =   405
         TabIndex        =   39
         Top             =   5850
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1191
         TopColour       =   65535
         BottomColour    =   65535
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   855
         Index           =   18
         Left            =   1680
         TabIndex        =   40
         Top             =   2100
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1508
         TopColour       =   16776960
         BottomColour    =   16776960
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   16776960
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   675
         Index           =   19
         Left            =   6165
         TabIndex        =   41
         Top             =   4560
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1191
         TopColour       =   49344
         BottomColour    =   49344
         LineWidth       =   10
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   49344
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   375
         Index           =   0
         Left            =   3855
         TabIndex        =   46
         Top             =   5280
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   375
         Index           =   20
         Left            =   2490
         TabIndex        =   47
         Top             =   5385
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   510
         Index           =   2
         Left            =   9780
         TabIndex        =   54
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   900
         TopColour       =   255
         BottomColour    =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   4
         Left            =   2760
         Picture         =   "CMA_purifier.frx":0000
         Stretch         =   -1  'True
         Top             =   6270
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   2760
         Picture         =   "CMA_purifier.frx":030A
         Stretch         =   -1  'True
         Top             =   5860
         Width           =   255
      End
      Begin VB.OLE OLE6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   7620
         OleObjectBlob   =   "CMA_purifier.frx":0614
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         TabIndex        =   53
         Top             =   2280
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   6720
         OleObjectBlob   =   "CMA_purifier.frx":262C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         TabIndex        =   52
         Top             =   4230
         Width           =   165
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   6870
         OleObjectBlob   =   "CMA_purifier.frx":4644
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         TabIndex        =   51
         Top             =   15
         Width           =   165
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   2940
         OleObjectBlob   =   "CMA_purifier.frx":665C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         TabIndex        =   50
         Top             =   30
         Width           =   165
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   2670
         OleObjectBlob   =   "CMA_purifier.frx":8674
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         TabIndex        =   49
         Top             =   1650
         Width           =   165
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   525
         Left            =   2055
         OleObjectBlob   =   "CMA_purifier.frx":A68C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         TabIndex        =   48
         Top             =   5355
         Width           =   165
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   2190
         X2              =   3285
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3840
         X2              =   3840
         Y1              =   2040
         Y2              =   2160
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   5640
         X2              =   4560
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Image Image23 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2880
         Picture         =   "CMA_purifier.frx":C6A4
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   255
      End
      Begin VB.Line Line51 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   2190
         X2              =   3135
         Y1              =   5985
         Y2              =   5985
      End
      Begin VB.Image Image17 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6120
         Picture         =   "CMA_purifier.frx":C9AE
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   255
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00008080&
         BorderWidth     =   3
         X1              =   6495
         X2              =   5580
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   3600
         X2              =   3600
         Y1              =   1987
         Y2              =   1620
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   3600
         X2              =   5760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   5535
         Top             =   5400
         Width           =   315
      End
      Begin VB.Line Line44 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   5100
         X2              =   5100
         Y1              =   3784
         Y2              =   3540
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   1530
         Index           =   0
         Left            =   4440
         Picture         =   "CMA_purifier.frx":CCB8
         Stretch         =   -1  'True
         Top             =   3660
         Width           =   1275
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00000000&
         X1              =   1320
         X2              =   1320
         Y1              =   6600
         Y2              =   6360
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   7800
         Picture         =   "CMA_purifier.frx":130FA
         Stretch         =   -1  'True
         Top             =   5100
         Width           =   255
      End
      Begin VB.Line Line57 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   7920
         X2              =   7290
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line55 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   7935
         X2              =   7935
         Y1              =   3810
         Y2              =   6000
      End
      Begin VB.Line Line54 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   5025
         X2              =   3300
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3180
         Shape           =   3  'Circle
         Top             =   2865
         Width           =   255
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " T "
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
         Left            =   5160
         TabIndex        =   15
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   7260
         Picture         =   "CMA_purifier.frx":13404
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   255
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   9180
         Picture         =   "CMA_purifier.frx":1370E
         Stretch         =   -1  'True
         Top             =   225
         Width           =   255
      End
      Begin VB.Image Image12 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   480
         Picture         =   "CMA_purifier.frx":13A18
         Stretch         =   -1  'True
         Top             =   3420
         Width           =   255
      End
      Begin VB.Image Image13 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   9150
         Picture         =   "CMA_purifier.frx":13D22
         Stretch         =   -1  'True
         Top             =   3510
         Width           =   255
      End
      Begin VB.Image Image14 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4125
         Picture         =   "CMA_purifier.frx":1402C
         Stretch         =   -1  'True
         Top             =   495
         Width           =   255
      End
      Begin VB.Image Image31 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6360
         Picture         =   "CMA_purifier.frx":14336
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   255
      End
      Begin VB.Image Image16 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6000
         Picture         =   "CMA_purifier.frx":14640
         Stretch         =   -1  'True
         Top             =   3675
         Width           =   255
      End
      Begin VB.Image Image36 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2055
         Picture         =   "CMA_purifier.frx":1494A
         Stretch         =   -1  'True
         Top             =   4995
         Width           =   255
      End
      Begin VB.Image Image28 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2790
         Picture         =   "CMA_purifier.frx":14C54
         Stretch         =   -1  'True
         Top             =   4515
         Width           =   255
      End
      Begin VB.Image Image30 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2790
         Picture         =   "CMA_purifier.frx":14F5E
         Stretch         =   -1  'True
         Top             =   4155
         Width           =   255
      End
      Begin VB.Image Image32 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2790
         Picture         =   "CMA_purifier.frx":15268
         Stretch         =   -1  'True
         Top             =   3795
         Width           =   255
      End
      Begin VB.Image Image38 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   840
         Picture         =   "CMA_purifier.frx":15572
         Stretch         =   -1  'True
         Top             =   2070
         Width           =   255
      End
      Begin VB.Line Line46 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   5160
         X2              =   4320
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line15 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   5115
         X2              =   5115
         Y1              =   3131
         Y2              =   2520
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3840
         X2              =   3000
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line26 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   600
         X2              =   1329
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3045
         X2              =   6495
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line Line27 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   8520
         X2              =   9249
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   8430
         X2              =   9285
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   600
         X2              =   3285
         Y1              =   5130
         Y2              =   5130
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   600
         X2              =   600
         Y1              =   5130
         Y2              =   1365
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LO SUMP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   1290
         TabIndex        =   21
         Top             =   5535
         Width           =   930
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   9300
         X2              =   9300
         Y1              =   225
         Y2              =   592
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   2595
         X2              =   2580
         Y1              =   6000
         Y2              =   5145
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LO STORAGE TANK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   7110
         TabIndex        =   20
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LO RENOVATING TANK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   975
         TabIndex        =   19
         Top             =   240
         Width           =   1920
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3765
         Top             =   2070
         Width           =   1110
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3765
         Top             =   2325
         Width           =   1110
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3765
         Top             =   2205
         Width           =   1110
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3765
         Top             =   2445
         Width           =   1110
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   5760
         X2              =   5760
         Y1              =   2811
         Y2              =   1650
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         X1              =   5400
         X2              =   5765
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00008080&
         BorderWidth     =   3
         X1              =   6480
         X2              =   6480
         Y1              =   4635
         Y2              =   4440
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LO SLUDGE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   6135
         TabIndex        =   18
         Top             =   5250
         Width           =   765
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LO SEPARATOR"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   4425
         TabIndex        =   17
         Top             =   5175
         Width           =   1305
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LO BUNKER LINE"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   9480
         TabIndex        =   16
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PUMP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   3480
         TabIndex        =   14
         Top             =   2910
         Width           =   435
      End
      Begin VB.Line Line8 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   7285
         X2              =   5220
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line43 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3270
         X2              =   4095
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line45 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3285
         X2              =   3270
         Y1              =   6390
         Y2              =   2280
      End
      Begin VB.Line Line47 
         X1              =   8475
         X2              =   8840
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line48 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   9255
         X2              =   9255
         Y1              =   1320
         Y2              =   6030
      End
      Begin VB.Line Line49 
         X1              =   2940
         X2              =   3183
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line11 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   6480
         X2              =   6480
         Y1              =   600
         Y2              =   3780
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   3
         X1              =   975
         X2              =   1704
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Line Line53 
         X1              =   1455
         X2              =   1698
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   2550
         X2              =   3158
         Y1              =   3930
         Y2              =   3930
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   2550
         X2              =   3158
         Y1              =   4275
         Y2              =   4275
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   2550
         X2              =   3158
         Y1              =   4635
         Y2              =   4635
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   2550
         X2              =   2550
         Y1              =   2925
         Y2              =   4635
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   4499
         X2              =   3405
         Y1              =   3930
         Y2              =   3930
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   4499
         X2              =   3405
         Y1              =   4635
         Y2              =   4635
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   3405
         X2              =   4499
         Y1              =   4275
         Y2              =   4275
      End
      Begin VB.Label Label23 
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
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   1695
         TabIndex        =   13
         Top             =   1710
         Width           =   1020
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WATER TANK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   1695
         TabIndex        =   12
         Top             =   1890
         Width           =   1110
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TANK"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   6345
         TabIndex        =   11
         Top             =   5415
         Width           =   390
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00000000&
         X1              =   7005
         X2              =   6825
         Y1              =   4845
         Y2              =   4845
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00000000&
         Visible         =   0   'False
         X1              =   615
         X2              =   1110
         Y1              =   3855
         Y2              =   3855
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00000000&
         X1              =   5880
         X2              =   5880
         Y1              =   3964
         Y2              =   3720
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00000000&
         Visible         =   0   'False
         X1              =   5760
         X2              =   5760
         Y1              =   3480
         Y2              =   3724
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00000000&
         X1              =   7680
         X2              =   7680
         Y1              =   3420
         Y2              =   3786
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00000000&
         X1              =   6240
         X2              =   6483
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00000000&
         Visible         =   0   'False
         X1              =   9600
         X2              =   9240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line10 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3375
         X2              =   9255
         Y1              =   6000
         Y2              =   6000
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1740
      TabIndex        =   42
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "CMA_purifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer



Private Sub Command1_Click()

' 7/8/00
Lube_Oil.Show


End Sub

Private Sub Command3_Click()


' 10/4/00
'sy030.Show

' 7/8/00 sy030_LSS.Show



End Sub

Private Sub Command3D14_Click()

' 7/8/00  sy064_LSS.Show   ' To DO sep
 
End Sub

Private Sub Command3D19_Click()

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
' 6/1/98
   Call setPCDI(PumpIO(4, 10))     ' ON/OFF

End Sub

Private Sub Command3D2_Click()

End Sub

Private Sub Command3D4_Click()

' 7/8/00  sy053_LSS.Show    ' To HO SEP

End Sub

Private Sub Form_Load()
On Error Resume Next
Call JYAScreenInit(Me, idsy041)

' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsy041)
    HavenStretch1.Visible = True
End If
End Sub



Private Sub Timer2_Timer()

End Sub


Private Sub Haven3WayValve1_Click()

   Control_PC_check
   If This_Control_status = False Then Exit Sub

'If VB_DO(337) Then
 If VB_lamps(1) Then
   Call setPCDI(285)    '  Set valve TO separator
Else
    Call setPCDI(287)    '  Set valve TO BY_PASS Separatorr
End If

End Sub



Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 28/7/00 CMA
'  Call setPCDI(PumpIO_LSS(2, 22))
   Call setPCDI(PumpIO_LSS(4, 22))

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

' 28/7/00 CMA
   Call setPCDI(PumpIO_LSS(2, 22))

End Sub

Private Sub valve_Click(Index As Integer)


Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
' 27/7/00  NEW CMA_purifier
If Index < 12 Then
    Call valve_pcdi_LSS(Index + 550)
ElseIf Index = 20 Then
    Call valve_pcdi_LSS(Index + 550)
End If

   


' OLD LSS STUFF
'' 28/5/98  LO SEPARATOR Valves
'Select Case Index
'  Case 0
'    Call valve_pcdi_LSS(60)
'  Case 1
'    Call valve_pcdi_LSS(58)
'  Case 2
'    Call valve_pcdi_LSS(57)
'  Case 3
'    Call valve_pcdi_LSS(62)
'  Case 4
'    Call valve_pcdi_LSS(63)
'  Case 5, 6
'    Call valve_pcdi_LSS(64)
'  Case 7
'    Call valve_pcdi_LSS(61)
'  Case 8
'    Call valve_pcdi_LSS(60)
'  Case 9
'    Call valve_pcdi_LSS(58)
'  Case 10
'    Call valve_pcdi_LSS(59)
'  Case 11
'    Call valve_pcdi_LSS(63)
'End Select


End Sub

Private Sub Timer1_Timer()
Dim single_char As String * 1
Dim i As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

'  27/7/00
For i = 0 To 11
   Call update_vvalve_LSS(Me, 550 + i, i, -1)
Next i
   
   
'  27/7/00   PD NEW valve
   Call update_vvalve_LSS(Me, 570, 20, -1)


'==========================


'' 28/5/98 VALVES
'Call update_vvalve_LSS(Me, 60, 0, -1)
'
'Call update_vvalve_LSS(Me, 58, 1, -1)
'Call update_vvalve_LSS(Me, 57, 2, -1)
'Call update_vvalve_LSS(Me, 62, 3, -1)
'Call update_vvalve_LSS(Me, 63, 4, -1)
'Call update_vvalve_LSS(Me, 64, 6, -1)
'Call update_vvalve_LSS(Me, 61, 7, -1)
'Call update_vvalve_LSS(Me, 60, 8, -1)
'Call update_vvalve_LSS(Me, 58, 9, -1)
'Call update_vvalve_LSS(Me, 59, 10, -1)
'Call update_vvalve_LSS(Me, 63, 11, -1)


' SEPARATOR WATER VALVES
' SEAL/FLUSH Valve
single_char = Mid$(PC_INT_STRING, 347, 1)
If single_char = "1" Then
   valve(12).OpenBtmRight = True
   valve(12).OpenTopLeft = True
Else
   valve(12).OpenBtmRight = False
   valve(12).OpenTopLeft = False
End If

' MAKE-UP Valve
single_char = Mid$(PC_INT_STRING, 348, 1)
If single_char = "1" Then
   valve(13).OpenBtmRight = True
   valve(13).OpenTopLeft = True
Else
   valve(13).OpenBtmRight = False
   valve(13).OpenTopLeft = False
End If
' OPERATING Valve
single_char = Mid$(PC_INT_STRING, 349, 1)
If single_char = "1" Then
   valve(14).OpenBtmRight = True
   valve(14).OpenTopLeft = True
Else
   valve(14).OpenBtmRight = False
   valve(14).OpenTopLeft = False
End If


' Separator By-Pass Valve
  Haven3WayValve1.OpenTopTri = True
'If VB_lamps(1) Then
If LSS_LOGICALS(1) Then
  Haven3WayValve1.OpenSideTri = True
  Haven3WayValve1.OpenBtmTri = False
Else
  Haven3WayValve1.OpenSideTri = False
  Haven3WayValve1.OpenBtmTri = True
End If
 
 
' SEPARATOR ON/OFF
''' single_char = Mid$(PC_INT_STRING, pump_no + 250, 1)


' 28/7/00   LO SEP is pump 22 at CMA
'single_char = Mid$(PC_INT_STRING, 260, 1)     ' LO SEP is Pump number 10
single_char = Mid$(PC_INT_STRING, 272, 1)     ' LO SEP is Pump number 22
If single_char = "4" Or single_char = "5" Then

   LOSEP_flash = Not LOSEP_flash
   If LOSEP_flash Then
       Shape1(1).FillColor = &HFF00&         ' Bright Green
    Else
       Shape1(1).FillColor = &H8000&           ' DULL GREEN
    End If
ElseIf single_char = "2" Or single_char = "3" Then
       Shape1(1).FillColor = &HFF00&         ' Bright Green
Else
       Shape1(1).FillColor = &H8000&           ' DULL GREEN
End If

' SEPARATOR AUTO    ( LSS model is ONLY AUTO )
' If VB_lamps(7) Then
'        Shape1(0).FillColor = &HFFC0FF        ' BRIGHT PURPLE
'Else
'        Shape1(0).FillColor = &H800080       ' DULL PURPLE
'End If


'   28/7/00
'         PC_REALS(736) = zls03 * 100#
'         PC_REALS(737) = zls80 * 100#
'         PC_REALS(738) = zsludge * 100#
'         PC_REALS(739) = Zlwater * 100#
'         PC_REALS(740) = fls02
'         PC_REALS(741) = fls04


' RENOVATING Tank
valve(15).LineWidth = LSS_reals(736)
zls03.text = Format(LSS_reals(736), "##0.0") & "%"
'Call paint_alarm(Me.zls03, , , 0, 0)

' STORAGE Tank
valve(16).LineWidth = LSS_reals(737)
zls80.text = Format(LSS_reals(737), "##0.0") & "%"
'Call paint_alarm(Me.zls80, , , 0, 0)

' SUMP Tank
' 28/7/00   OLD LSS
'valve(17).LineWidth = LSS_reals(169)
'zml70.text = Format(LSS_reals(169), "##0") & "%"
'Call paint_alarm(Me.zml70, 52, 53, 0, 0)

' 28/7/00   NEW CMA
valve(17).LineWidth = LSS_reals(207)
ZLO_sump.text = Format(LSS_reals(207), "##0.0") & " %"
Call paint_alarm(Me.ZLO_sump, 130, 0, 0, 0)


' SLUDGE Tank
valve(19).LineWidth = LSS_reals(738)
zsludge.text = Format(LSS_reals(738), "##0") & "%"
'Call paint_alarm(Me.zsludge, , 0, 0, 0)

' SEPARATOR WATER Tank
valve(18).LineWidth = LSS_reals(739)
Zlwater.text = Format(LSS_reals(739), "##0") & "%"



' === 22/01/01 =15.021 ==
'tls04.text = Format(LSS_reals(743), "##0") & Chr(176) & "C"
tls04.text = Format(LSS_reals(743) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


' === 22/01/01 =15.021 ==
'    PLS02.text = Format(LSS_reals(744) * Press_mult, "##0.0") & Press_chars
'
'    fls02.text = Format(LSS_reals(740), "##0.0") & " kg/s"
'    fls03.text = Format(LSS_reals(700), "##0.0") & " kg/s"
'    fls10.text = Format(LSS_reals(701), "##0.0") & " kg/s"
'    fls80.text = Format(LSS_reals(702), "##0.0") & " kg/s"
'    fls20.text = Format(LSS_reals(703), "##0.0") & " kg/s"
'    fls04.text = Format(LSS_reals(741), "##0.0") & "kg/s"
If SI_units Then
    PLS02.text = Format(LSS_reals(744) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    fls02.text = Format(LSS_reals(740) * Flow_mult, "###0.0") & Flow_chars
    fls03.text = Format(LSS_reals(700) * Flow_mult, "###0.0") & Flow_chars
    fls10.text = Format(LSS_reals(701) * Flow_mult, "###0.0") & Flow_chars
    fls80.text = Format(LSS_reals(702) * Flow_mult, "###0.0") & Flow_chars
    fls20.text = Format(LSS_reals(703) * Flow_mult, "###0.0") & Flow_chars
    fls04.text = Format(LSS_reals(741) * Flow_mult, "###0.0") & Flow_chars
Else
    PLS02.text = Format(LSS_reals(744) * Press_mult, "###0") & Press_chars     ' psi
    
    fls02.text = Format(LSS_reals(740) * Flow_mult / 3600#, "###0.0") & " lb/s"
    fls03.text = Format(LSS_reals(700) * Flow_mult / 3600#, "###0.0") & " lb/s"
    fls10.text = Format(LSS_reals(701) * Flow_mult / 3600#, "###0.0") & " lb/s"
    fls80.text = Format(LSS_reals(702) * Flow_mult / 3600#, "###0.0") & " lb/s"
    fls20.text = Format(LSS_reals(703) * Flow_mult / 3600#, "###0.0") & " lb/s"
    fls04.text = Format(LSS_reals(741) * Flow_mult / 3600#, "###0.0") & " lb/s"
End If


'================================================================
' SAMPLE from Lub_oil...
'TLO_bearings.text = Format(LSS_reals(211) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
'If SI_units Then
'    PLO_throttle.text = Format(LSS_reals(217) * Press_mult, "##0.0") & Press_chars     ' Bar
'    FLO_sight_glass.text = Format(LSS_reals(202) * Flow_mult, "###0.0") & Flow_chars
'Else
'    PLO_throttle.text = Format(LSS_reals(217) * Press_mult, "###0") & Press_chars     ' psi
'    FLO_sight_glass.text = Format(LSS_reals(202) * Flow_mult / 3600#, "###0.0") & " lb/s"
'End If
'================================================================

' === 21/11/02 17.025 ===
Call update_pump_LSS(Me, 0, 22)      ' LO Purifier



End Sub

Private Sub ZLO_sump_Click()

' 28/7/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 124
Alarm_set_point.Show

End Sub

Private Sub Zlwater_Click()

' 22/1/98  LSS does NOT have an alarm....
'Unload Alarm_set_point   ' Do this first
'nAlarm_LO = 308
'Alarm_set_point.Show

End Sub

Private Sub zml70_Click()


End Sub

