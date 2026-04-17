VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{A6D9C597-48C2-11D1-8915-000021240899}#2.0#0"; "Haven3WayValve v1.41.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Lube_Oil 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lubricating Oil System"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   12600
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8175
      Left            =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   14420
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Left            =   10740
         TabIndex        =   74
         Top             =   3540
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton Trips_button 
         Caption         =   "TRIPS"
         Height          =   255
         Left            =   9360
         TabIndex        =   70
         Top             =   780
         Width           =   675
      End
      Begin VB.TextBox TLO_gearbox 
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
         Left            =   7500
         TabIndex        =   69
         Text            =   "88 F"
         Top             =   5460
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Trip_throttle 
         BackColor       =   &H008080FF&
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9300
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox PLO_throttle 
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
         Left            =   10860
         TabIndex        =   64
         Text            =   "57 psi"
         Top             =   900
         Width           =   675
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H0000FF00&
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   8760
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   555
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
         Index           =   3
         Left            =   6480
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
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
         Index           =   3
         Left            =   7680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TLO_bearings 
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
         Left            =   7020
         TabIndex        =   54
         Text            =   "88 F"
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox TLO_cooler_out 
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
         Left            =   300
         TabIndex        =   53
         Text            =   "88 F"
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox FLO_sight_glass 
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
         Left            =   3480
         TabIndex        =   52
         Text            =   "xx kg/s"
         Top             =   4620
         Width           =   915
      End
      Begin VB.TextBox VV_position 
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
         Left            =   1920
         TabIndex        =   51
         Text            =   "20 %"
         Top             =   3030
         Width           =   615
      End
      Begin VB.TextBox FLO_header 
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
         Left            =   2820
         TabIndex        =   50
         Text            =   "xx kg/s"
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox FLO_sump_out 
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
         Left            =   60
         TabIndex        =   49
         Text            =   "xx kg/s"
         Top             =   4080
         Width           =   915
      End
      Begin VB.TextBox FLO_bearings 
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
         Left            =   4560
         TabIndex        =   48
         Text            =   "xx kg/s"
         Top             =   900
         Width           =   855
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
         Height          =   240
         Left            =   4020
         TabIndex        =   47
         Text            =   "T"
         Top             =   1020
         Width           =   195
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   300
         Top             =   3360
      End
      Begin VB.TextBox PLO_bearings 
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
         Left            =   8100
         TabIndex        =   43
         Text            =   "500 psi"
         Top             =   2880
         Width           =   855
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
         Index           =   1
         Left            =   2040
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   6300
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Left            =   2100
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4920
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Left            =   1560
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   6300
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Left            =   1620
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4920
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Left            =   1080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6300
         UseMaskColor    =   -1  'True
         Width           =   495
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
         Left            =   1140
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4920
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox PLO_out 
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
         Left            =   2040
         TabIndex        =   34
         Text            =   "500 psi"
         Top             =   3720
         Width           =   675
      End
      Begin Haven3WayValvePrj.Haven3WayValve Haven3WayValve1 
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   3120
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   873
         OpenSideTri     =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         LineControl     =   -1  'True
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "L.O. Purifier"
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6960
         Width           =   795
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Bearings"
         Height          =   495
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox TLO_sump 
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
         Left            =   6420
         TabIndex        =   25
         Text            =   "88 F"
         Top             =   6960
         Width           =   615
      End
      Begin VB.TextBox ZLO_header 
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
         Left            =   1500
         TabIndex        =   17
         Text            =   "20 %"
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox ZLO_sump 
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
         Left            =   6420
         TabIndex        =   2
         Text            =   "20 %"
         Top             =   7260
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   915
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   6780
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1614
         TopColour       =   8421504
         BottomColour    =   8421504
         LineWidth       =   15
         OrientVertical  =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   7
         ColourOpen      =   8421504
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   10
         Left            =   1440
         TabIndex        =   3
         Top             =   3900
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
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   11
         Left            =   720
         TabIndex        =   4
         Top             =   4320
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
         Index           =   12
         Left            =   2220
         TabIndex        =   5
         Top             =   4320
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
         Height          =   435
         Index           =   13
         Left            =   720
         TabIndex        =   6
         Top             =   5700
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
         Index           =   14
         Left            =   2220
         TabIndex        =   7
         Top             =   5700
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
         Height          =   435
         Index           =   16
         Left            =   1440
         TabIndex        =   8
         Top             =   5220
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
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   1
         Left            =   2640
         TabIndex        =   12
         Top             =   2580
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   2
         Left            =   1320
         TabIndex        =   13
         Top             =   1560
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   615
         Index           =   3
         Left            =   1500
         TabIndex        =   15
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   5
         Left            =   10740
         TabIndex        =   28
         Top             =   7020
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   6
         Left            =   10680
         TabIndex        =   29
         Top             =   6060
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
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   7
         Left            =   2880
         TabIndex        =   30
         Top             =   4320
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
         Index           =   8
         Left            =   3240
         TabIndex        =   31
         Top             =   4920
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   33
         Top             =   3480
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         TopColour       =   255
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   1380
         TabIndex        =   45
         Top             =   4380
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   1380
         TabIndex        =   46
         Top             =   5760
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   17
         Left            =   3240
         TabIndex        =   55
         Top             =   3660
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         ColourDisabled  =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   18
         Left            =   10140
         TabIndex        =   59
         Top             =   1800
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   9
         Left            =   9000
         TabIndex        =   60
         Top             =   1020
         Width           =   255
         _ExtentX        =   450
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
         Height          =   375
         Index           =   2
         Left            =   9480
         TabIndex        =   61
         Top             =   1140
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   15
         Left            =   10500
         TabIndex        =   62
         Top             =   1020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   767
         TopColour       =   255
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 ZLO_throttle 
         Height          =   615
         Left            =   10320
         TabIndex        =   68
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   1085
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   315
         Index           =   590
         Left            =   1800
         TabIndex        =   71
         Top             =   7800
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
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   435
         Index           =   557
         Left            =   7200
         TabIndex        =   75
         Top             =   7260
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
      Begin VB.Label Label18 
         Caption         =   "Elec"
         Height          =   195
         Left            =   9720
         TabIndex        =   73
         Top             =   1680
         Width           =   435
      End
      Begin VB.Shape Elec_problem 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   9360
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Mech"
         Height          =   195
         Left            =   8460
         TabIndex        =   72
         Top             =   1680
         Width           =   435
      End
      Begin VB.Shape Mech_problem 
         BorderWidth     =   2
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   8880
         Shape           =   3  'Circle
         Top             =   1680
         Width           =   375
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   0
         X1              =   4580
         X2              =   4580
         Y1              =   7680
         Y2              =   8040
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   12
         X1              =   1920
         X2              =   1920
         Y1              =   7680
         Y2              =   8160
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   7620
         Picture         =   "Lube_Oil.frx":0000
         Stretch         =   -1  'True
         Top             =   6780
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   4
         Left            =   7680
         Picture         =   "Lube_Oil.frx":030A
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   255
      End
      Begin VB.Line Line14 
         X1              =   6900
         X2              =   7260
         Y1              =   6540
         Y2              =   6540
      End
      Begin VB.Line Line13 
         Visible         =   0   'False
         X1              =   7380
         X2              =   7740
         Y1              =   5580
         Y2              =   5520
      End
      Begin VB.Shape Shape18 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2780
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   195
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2610
         Shape           =   3  'Circle
         Top             =   4080
         Width           =   195
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   2520
         X2              =   3000
         Y1              =   4300
         Y2              =   4300
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   2520
         X2              =   3000
         Y1              =   4100
         Y2              =   4100
      End
      Begin VB.Label Label17 
         Caption         =   "Accumulator"
         Height          =   195
         Left            =   10560
         TabIndex        =   67
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label16 
         Caption         =   "Throttle valves"
         Height          =   195
         Left            =   10800
         TabIndex        =   65
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   3
         Left            =   10860
         Picture         =   "Lube_Oil.frx":0614
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Throttle Hydraulic Pump"
         Height          =   195
         Left            =   8220
         TabIndex        =   63
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   60
         X1              =   10380
         X2              =   10380
         Y1              =   1320
         Y2              =   660
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   59
         X1              =   8340
         X2              =   10200
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   57
         X1              =   8820
         X2              =   9540
         Y1              =   -780
         Y2              =   -780
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   56
         X1              =   10200
         X2              =   10200
         Y1              =   1320
         Y2              =   6060
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   25
         X1              =   8760
         X2              =   10980
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Flange_line 
         Index           =   1
         X1              =   3420
         X2              =   3420
         Y1              =   3660
         Y2              =   3420
      End
      Begin VB.Line Flange_line 
         Index           =   2
         X1              =   3420
         X2              =   3300
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Flange_line 
         Index           =   0
         X1              =   3300
         X2              =   3300
         Y1              =   3660
         Y2              =   3420
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   2
         Left            =   4560
         Picture         =   "Lube_Oil.frx":091E
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Sight_glass_light 
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   4200
         Width           =   195
      End
      Begin VB.Line Line10 
         BorderStyle     =   3  'Dot
         X1              =   720
         X2              =   4020
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line9 
         BorderStyle     =   3  'Dot
         X1              =   1080
         X2              =   2520
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line8 
         BorderStyle     =   3  'Dot
         X1              =   1080
         X2              =   1080
         Y1              =   1200
         Y2              =   3360
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   3240
         Picture         =   "Lube_Oil.frx":0C28
         Stretch         =   -1  'True
         Top             =   5580
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   1
         Left            =   6180
         Picture         =   "Lube_Oil.frx":0F32
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   0
         Left            =   2220
         Picture         =   "Lube_Oil.frx":123C
         Stretch         =   -1  'True
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   16
         Left            =   6780
         Picture         =   "Lube_Oil.frx":1546
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   15
         Left            =   7260
         Picture         =   "Lube_Oil.frx":1850
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   14
         Left            =   6780
         Picture         =   "Lube_Oil.frx":1B5A
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   13
         Left            =   6780
         Picture         =   "Lube_Oil.frx":1E64
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   12
         Left            =   6600
         Picture         =   "Lube_Oil.frx":216E
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   11
         Left            =   5160
         Picture         =   "Lube_Oil.frx":2478
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   10
         Left            =   6480
         Picture         =   "Lube_Oil.frx":2782
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   9
         Left            =   5040
         Picture         =   "Lube_Oil.frx":2A8C
         Stretch         =   -1  'True
         Top             =   3900
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   8
         Left            =   5220
         Picture         =   "Lube_Oil.frx":2D96
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   7
         Left            =   4740
         Picture         =   "Lube_Oil.frx":30A0
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   6
         Left            =   5100
         Picture         =   "Lube_Oil.frx":33AA
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   5
         Left            =   6780
         Picture         =   "Lube_Oil.frx":36B4
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   4
         Left            =   6780
         Picture         =   "Lube_Oil.frx":39BE
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   3
         Left            =   6480
         Picture         =   "Lube_Oil.frx":3CC8
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   2
         Left            =   7260
         Picture         =   "Lube_Oil.frx":3FD2
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   1
         Left            =   7860
         Picture         =   "Lube_Oil.frx":42DC
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Index           =   0
         Left            =   8220
         Picture         =   "Lube_Oil.frx":45E6
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "Air Vent"
         Height          =   315
         Left            =   4560
         TabIndex        =   32
         Top             =   180
         Width           =   615
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   14
         X1              =   3180
         X2              =   2700
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   4440
         X2              =   4440
         Y1              =   6960
         Y2              =   120
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   53
         X1              =   3360
         X2              =   3360
         Y1              =   7680
         Y2              =   3720
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   51
         X1              =   7080
         X2              =   8040
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   49
         X1              =   3840
         X2              =   3840
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   48
         X1              =   3660
         X2              =   3840
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   44
         X1              =   7080
         X2              =   8040
         Y1              =   6900
         Y2              =   6900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Gear Box"
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
         Index           =   2
         Left            =   7200
         TabIndex        =   21
         Top             =   3120
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape13 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   3255
         Left            =   7080
         Top             =   2160
         Width           =   735
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   38
         X1              =   7380
         X2              =   7380
         Y1              =   6060
         Y2              =   1800
      End
      Begin VB.Image Arrow_up 
         Height          =   255
         Left            =   4320
         Picture         =   "Lube_Oil.frx":48F0
         Stretch         =   -1  'True
         Top             =   180
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   43
         X1              =   6900
         X2              =   6720
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   42
         X1              =   6720
         X2              =   6720
         Y1              =   5340
         Y2              =   4800
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   41
         X1              =   6900
         X2              =   5340
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   34
         X1              =   6900
         X2              =   6900
         Y1              =   6780
         Y2              =   4980
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   36
         X1              =   5280
         X2              =   8340
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Shape Shape11 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5125
         Top             =   2640
         Width           =   375
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   6480
         Top             =   2640
         Width           =   615
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   6480
         Top             =   4620
         Width           =   615
      End
      Begin VB.Shape Shape15 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5010
         Top             =   4620
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Main Thrust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7920
         TabIndex        =   24
         Top             =   3540
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape17 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   7860
         Top             =   3480
         Width           =   615
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   33
         X1              =   7980
         X2              =   7980
         Y1              =   3480
         Y2              =   1320
      End
      Begin VB.Label Label11 
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
         Height          =   435
         Index           =   0
         Left            =   5700
         TabIndex        =   23
         Top             =   4440
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   5700
         TabIndex        =   22
         Top             =   2520
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   32
         X1              =   6600
         X2              =   4860
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   31
         X1              =   6600
         X2              =   6600
         Y1              =   4620
         Y2              =   3600
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   30
         X1              =   5160
         X2              =   5160
         Y1              =   4620
         Y2              =   3600
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   29
         X1              =   4860
         X2              =   4860
         Y1              =   3600
         Y2              =   1320
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   28
         X1              =   4860
         X2              =   7380
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   27
         X1              =   6600
         X2              =   6600
         Y1              =   2640
         Y2              =   1800
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   26
         X1              =   5220
         X2              =   5220
         Y1              =   2640
         Y2              =   1800
      End
      Begin VB.Shape Shape16 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7740
         Top             =   3600
         Width           =   915
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   3480
         X2              =   3480
         Y1              =   420
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   3480
         X2              =   4440
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   3960
         X2              =   4440
         Y1              =   1800
         Y2              =   1500
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3960
         X2              =   3960
         Y1              =   2100
         Y2              =   1800
      End
      Begin VB.Label Label8 
         Caption         =   "LO Cooler"
         Height          =   195
         Left            =   1860
         TabIndex        =   20
         Top             =   1680
         Width           =   855
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   24
         X1              =   2760
         X2              =   2760
         Y1              =   3840
         Y2              =   2580
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3660
         Top             =   2100
         Width           =   375
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Index           =   21
         X1              =   3480
         X2              =   3300
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Index           =   20
         X1              =   3060
         X2              =   2880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   19
         X1              =   3180
         X2              =   3180
         Y1              =   720
         Y2              =   1320
      End
      Begin VB.Shape Sight_Glass_dark 
         BorderWidth     =   3
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   4200
         Width           =   195
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   4140
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "LO GravityTank"
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   0
         Width           =   1455
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   18
         X1              =   1440
         X2              =   1440
         Y1              =   1920
         Y2              =   1320
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   17
         X1              =   8820
         X2              =   1440
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   16
         X1              =   3180
         X2              =   3180
         Y1              =   3360
         Y2              =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Filling System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3180
         TabIndex        =   14
         Top             =   2940
         Width           =   615
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   15
         X1              =   3360
         X2              =   2760
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1320
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1320
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1320
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1320
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1320
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   13
         X1              =   2760
         X2              =   2760
         Y1              =   6000
         Y2              =   3780
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   480
         X2              =   480
         Y1              =   4920
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   12
         X1              =   600
         X2              =   2760
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   600
         X2              =   2760
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   600
         X2              =   600
         Y1              =   6000
         Y2              =   4620
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   6
         X1              =   1260
         X2              =   1980
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   8
         X1              =   1260
         X2              =   1260
         Y1              =   4620
         Y2              =   4200
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   9
         X1              =   1260
         X2              =   1980
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   10
         X1              =   1980
         X2              =   1980
         Y1              =   5520
         Y2              =   6000
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   11
         X1              =   1260
         X2              =   1260
         Y1              =   5520
         Y2              =   6000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "No 1"
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
         Left            =   600
         TabIndex        =   10
         Top             =   4920
         Width           =   435
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   7
         X1              =   1980
         X2              =   1980
         Y1              =   4200
         Y2              =   4620
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "No 2"
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
         Left            =   600
         TabIndex        =   9
         Top             =   6300
         Width           =   375
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   2
         X1              =   180
         X2              =   180
         Y1              =   5100
         Y2              =   7620
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   1
         X1              =   780
         X2              =   180
         Y1              =   7620
         Y2              =   7620
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   300
         Shape           =   3  'Circle
         Top             =   4920
         Width           =   195
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   300
         Shape           =   3  'Circle
         Top             =   5100
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   300
         X2              =   300
         Y1              =   4920
         Y2              =   5340
      End
      Begin VB.Label Label1 
         Caption         =   "LO Sump"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   7800
         Width           =   1155
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   180
         X2              =   360
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Label Label4 
         Caption         =   "Sight Glass"
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
         Left            =   4080
         TabIndex        =   18
         ToolTipText     =   "Sight Glass will be yellow if there is oil flow to the sump"
         Top             =   4080
         Width           =   435
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   35
         X1              =   8340
         X2              =   8340
         Y1              =   6060
         Y2              =   4020
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   37
         X1              =   5280
         X2              =   5280
         Y1              =   6060
         Y2              =   4800
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   39
         X1              =   5340
         X2              =   5340
         Y1              =   3420
         Y2              =   2820
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   40
         X1              =   6900
         X2              =   6900
         Y1              =   4440
         Y2              =   2820
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   23
         X1              =   3840
         X2              =   3840
         Y1              =   6720
         Y2              =   4440
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   5
         X1              =   360
         X2              =   600
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   22
         X1              =   3840
         X2              =   3840
         Y1              =   4080
         Y2              =   1440
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Air Separator"
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
         TabIndex        =   19
         Top             =   2700
         Width           =   735
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2145
         Left            =   5340
         OleObjectBlob   =   "Lube_Oil.frx":4BFA
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisLP&Astn.vsd"
         TabIndex        =   41
         Top             =   3720
         Width           =   1215
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1125
         Left            =   5460
         OleObjectBlob   =   "Lube_Oil.frx":C612
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\My Documents\visiotbne.vsd"
         TabIndex        =   42
         Top             =   2160
         Width           =   1035
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   44
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "Lube_Oil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer
' 7/12/00
Private Trip_throttleMouseDown As Boolean
Private FLASH_Trip_throttle As Boolean

Private Sub Command1_Click()

' 25/7/00
CMA_purifier.Show


End Sub

Private Sub Command2_Click()

'FrmHardwareTest.ZOrder

End Sub

Private Sub Command5_Click()

Main_Bearings.Show


End Sub

Private Sub Form_Activate()

'Gauge1.Resize 2500
'Gauge1.GaugeLabel 1, "-50"
'Gauge1.GaugeLabel 2, "40"
'Gauge1.GaugeLabel 3, "30"
'Gauge1.GaugeLabel 4, "20"
'Gauge1.GaugeLabel 5, "10"
'Gauge1.GaugeLabel 6, "0"
'
'Gauge1.ScaleText = "kV"

'Gauge1.Value = 127

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idLube_Oil)

' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' === 11/12/01 14.113 ===
    Trips_button.Visible = False
    
    ONpump(0).Visible = False
    OFFpump(0).Visible = False
    sbyPUMP(0).Visible = False
    ONpump(1).Visible = False
    OFFpump(1).Visible = False
    sbyPUMP(1).Visible = False
    
    ONpump(3).Visible = False
    Trip_throttle.Visible = False
    
    ' 7/9/00  point 8.94
    VV_position.Visible = False
    
Else

    ' === 11/12/01 14.113 ===
    Trips_button.Visible = True

    
    ONpump(0).Visible = True
    OFFpump(0).Visible = True
    sbyPUMP(0).Visible = True
    ONpump(1).Visible = True
    OFFpump(1).Visible = True
    sbyPUMP(1).Visible = True
    
    ONpump(3).Visible = True
    Trip_throttle.Visible = True
    
      ' 7/9/00  point 8.94
    VV_position.Visible = True

End If



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    'HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idLube_Oil)
   ' HavenStretch1.Visible = True
End If

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
   
If This_Control_status = False Then Exit Sub
   
   
' 15/2/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
  
  
'   Note:  PumpIO(4=OFF,nPump)
 
 ' 16/9/99
'If index = 0 Then
'   Call setPCDI(PumpIO_LSS(4, 8))
'ElseIf index = 1 Then
'   Call setPCDI(PumpIO_LSS(4, 9))
'End If

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 7))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 19))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 15/2/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
  
 ' 16/9/99
'   Note:  PumpIO_LSS(2=ON,nPump)
'If index = 0 Then
'   Call setPCDI(PumpIO_LSS(2, 8))
'ElseIf index = 1 Then
'   Call setPCDI(PumpIO_LSS(2, 9))
'End If

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 7))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 19))
ElseIf Index = 3 Then
   Call setPCDI(297)    ' Throttle Boost Pump ( Toggles )
End If

End Sub

Private Sub PLO_bearings_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 118
Alarm_set_point.Show

End Sub

Private Sub PLO_out_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 119
Alarm_set_point.Show

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
' 15/2/00
If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
End If
  
  
'   Note:  PumpIO(5=Stand-bye,nPump)
 ' 16/9/99
'If index = 0 Then
'   Call setPCDI(PumpIO_LSS(5, 8))
'ElseIf index = 1 Then
'   Call setPCDI(PumpIO_LSS(5, 9))
'End If

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 7))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(5, 19))
End If


End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim Throttle_level As Integer




If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' 8/12/01
'c         === 8/12/01 Throttle Boost Pump Mechanical problem ===
'             LSS_LOGICALS(90) = DO(1043)
'c         === 8/12/01 Throttle Boost Pump Electrical problem ===
'             LSS_LOGICALS(91) = DO(1047)

If LSS_LOGICALS(90) Then
    Mech_problem.FillColor = &H8080FF  ' RED
Else
    Mech_problem.FillColor = &HC0C0C0    ' GREY
End If
If LSS_LOGICALS(91) Then
    Elec_problem.FillColor = &H8080FF  ' RED
Else
    Elec_problem.FillColor = &HC0C0C0    ' GREY
End If



' 7/12/00
FLASH_Trip_throttle = Not FLASH_Trip_throttle

' 7/1/00
Call paint_alarm(Me.PLO_bearings, 118, 0, 0, 0)
Call paint_alarm(Me.PLO_out, 119, 0, 0, 0)


' 28/03/01  Alarm(120) is now on T_bearing(i) and NOT TLO_bearings
'Call paint_alarm(Me.TLO_bearings, 120, 0, 0, 0)

Call paint_alarm(Me.ZLO_header, 124, 0, 0, 0)
Call paint_alarm(Me.TLO_cooler_out, 126, 132, 0, 0)
Call paint_alarm(Me.ZLO_sump, 130, 0, 0, 0)



' TRIP_Throttle_DO_flash  /1040/

' 25/8/00
'             LSS_LOGICALS(30) = DO(1040)

'   23/11/99  TRIP_Throttle_DO
'  LSS_LOGICALS(28) = DO(1041)
'If LSS_LOGICALS(28) Then
If LSS_LOGICALS(28) Then

   ' 7/12/00
   ' Trip_throttle.BackColor = &H8080FF     '  Light RED
    
   If Trip_throttleMouseDown Then
        If FLASH_Trip_throttle Then
            Trip_throttle.BackColor = &H8080FF     '  Light RED
        Else
            Trip_throttle.BackColor = &HC0C0C0      'GREY
        End If
   Else
            Trip_throttle.BackColor = &H8080FF     '  Light RED
   End If
   
ElseIf LSS_LOGICALS(30) Then

        ' 7/12/00
        If FLASH_Trip_throttle Then
            Trip_throttle.BackColor = &H8080FF     '  Light RED
        Else
            Trip_throttle.BackColor = &HC0C0C0      'GREY
        End If
        
Else
    Trip_throttle.BackColor = &HC0C0C0     'GREY
End If

'   23/11/99   Throttle_pump_DO  /1039/
'             LSS_LOGICALS(29) = DO(1039)
If LSS_LOGICALS(29) Then
     ONpump(3).BackColor = &HFF00&        '  GREEN
     
     ' 6/12/00
    Havenpump1(2).PumpOn = True
    
Else
    ONpump(3).BackColor = &HC0C0C0     'GREY
    
     ' 6/12/00
    Havenpump1(2).PumpOn = False
    
End If


' 12/11/99 Throttle Boost Pump
' 6/12/00 See above ( Note: D_OUTPUT( does NOT get networked to the "ACT's" )
'HavenPump1(2).PumpOn = D_OUTPUT(1039)


VV_position.text = Format(LSS_reals(210), "##0.0") & " %"
If LSS_reals(210) > 99.8 Then
   Haven3WayValve1.OpenBtmTri = False
   Haven3WayValve1.OpenSideTri = True
ElseIf LSS_reals(210) > 0.1 Then
   Haven3WayValve1.OpenBtmTri = True
   Haven3WayValve1.OpenSideTri = True
Else
   Haven3WayValve1.OpenBtmTri = True
   Haven3WayValve1.OpenSideTri = False
End If

If LSS_reals(202) > 0.1 Then
    If Sight_glass_light.Visible Then
        Sight_glass_light.Visible = False
        Sight_Glass_dark.Visible = True
    Else
        Sight_glass_light.Visible = True
        Sight_Glass_dark.Visible = False
    End If
Else
        Sight_glass_light.Visible = False
        Sight_Glass_dark.Visible = True
End If

TLO_sump.text = Format(LSS_reals(205) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TLO_cooler_out.text = Format(LSS_reals(206) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TLO_bearings.text = Format(LSS_reals(211) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

 '         PC_REALS(217) = PLO_throttle
 '         PC_REALS(218) = ZLO_throttle

If SI_units Then

    PLO_throttle.text = Format(LSS_reals(217) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    PLO_out.text = Format(LSS_reals(200) * Press_mult, "##0.0") & Press_chars     ' Bar
    PLO_bearings.text = Format(LSS_reals(201) * Press_mult, "##0.0") & Press_chars     ' Bar
    
    FLO_bearings.text = Format(LSS_reals(203) * Flow_mult, "###0.0") & Flow_chars
    FLO_sump_out.text = Format(LSS_reals(204) * Flow_mult, "###0.0") & Flow_chars
    FLO_header.text = Format(LSS_reals(209) * Flow_mult, "###0.0") & Flow_chars
    FLO_sight_glass.text = Format(LSS_reals(202) * Flow_mult, "###0.0") & Flow_chars
Else
    PLO_throttle.text = Format(LSS_reals(217) * Press_mult, "###0") & Press_chars     ' psi
    
    PLO_out.text = Format(LSS_reals(200) * Press_mult, "###0") & Press_chars     ' psi
    PLO_bearings.text = Format(LSS_reals(201) * Press_mult, "###0") & Press_chars     ' psi
    
    FLO_bearings.text = Format(LSS_reals(203) * Flow_mult / 3600#, "###0.0") & " lb/s"
    FLO_sump_out.text = Format(LSS_reals(204) * Flow_mult / 3600#, "###0.0") & " lb/s"
    FLO_header.text = Format(LSS_reals(209) * Flow_mult / 3600#, "###0.0") & " lb/s"
    FLO_sight_glass.text = Format(LSS_reals(202) * Flow_mult / 3600#, "###0.0") & " lb/s"
End If

' 12/11/99
If (LSS_reals(218) > 99#) Then
  Throttle_level = 100
ElseIf (LSS_reals(218) < 1#) Then
 'Throttle_level = 100
  Throttle_level = 0
Else
  Throttle_level = LSS_reals(218)
End If
ZLO_throttle.LineWidth = Throttle_level

ZLO_sump.text = Format(LSS_reals(207), "##0.0") & " %"
Valve(0).LineWidth = LSS_reals(207)

ZLO_header.text = Format(LSS_reals(208), "##0.0") & " %"
Valve(3).LineWidth = LSS_reals(208)


' Note: update_pump_LSS(syxxx,index,nPump)

' OLD LSS..
'Call update_pump_LSS(Lube_Oil, 0, 8)      ' LO Pump # 1
'Call update_pump_LSS(Lube_Oil, 1, 9)      ' LO Pump # 2

Call update_pump_LSS(Lube_Oil, 0, 7)      ' LO Pump # 1
Call update_pump_LSS(Lube_Oil, 1, 19)      ' LO Pump # 2



' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)

'  16/9/99  Valve(0) & (3) are actually tanks...
For i = 1 To 2
   Call update_vvalve_LSS(Me, i + 240, i, -1)
Next i

' 3/4/00
'For i = 4 To 16
For i = 4 To 9
   Call update_vvalve_LSS(Me, i + 240, i, -1)
Next i

For i = 11 To 15
   Call update_vvalve_LSS(Me, i + 240, i, -1)
Next i

' 23/11/01  LO Sump Stem Heating Valve
   Call update_vvalve_LSS(Me, 590, 590, -1)

' === 21/11/02 17.023 ===
   Call update_vvalve_LSS(Me, 557, 557, -1)


' 3/4/00  LO Sump Filling Valve
   Call update_vvalve_LSS(Me, 257, 17, -1)

' 3/4/00  LO Filling Flange

' 20/7/00
'If Action(1, 811) Then
If Action(1, 815) Then
    Flange_line(0).Visible = False
    Flange_line(1).Visible = False
    Flange_line(2).Visible = False
Else
    Flange_line(0).Visible = True
    Flange_line(1).Visible = True
    Flange_line(2).Visible = True
End If


' 3/4/00  Pressure relief valve
'ZLO_sump.text = Format(LSS_reals(207), "##0.0") & " %"
If LSS_reals(207) > 1# Then
  If vv_status_LSS(251) > 0 And vv_status_LSS(252) = 0 And pp_status_LSS(7) > 1 Then
    ' OPEN relief V/V #1
      Valve(10).OpenBtmRight = True
      Valve(10).OpenTopLeft = True
  Else
      Valve(10).OpenBtmRight = False
      Valve(10).OpenTopLeft = False
  End If
  If vv_status_LSS(253) > 0 And vv_status_LSS(254) = 0 And pp_status_LSS(19) > 1 Then
    ' OPEN relief V/V #2
      Valve(16).OpenBtmRight = True
      Valve(16).OpenTopLeft = True
  Else
      Valve(16).OpenBtmRight = False
      Valve(16).OpenTopLeft = False
  End If
Else
      Valve(10).OpenBtmRight = False
      Valve(10).OpenTopLeft = False
      Valve(16).OpenBtmRight = False
      Valve(16).OpenTopLeft = False
End If





End Sub

Private Sub TLO_bearings_Click()

' 7/1/00
' === 10/10/01 ===
'Unload Alarm_set_point   ' Do this first
'nAlarm_HI = 120
'Alarm_set_point.Show

End Sub

Private Sub TLO_cooler_out_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 126
nAlarm_LO = 132
Alarm_set_point.Show

End Sub

Private Sub Trip_throttle_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


'  12/11/99
   Call setPCDI(298)    ' Throttle Boost Pump RESET


' 7/12/00
            Trip_throttle.BackColor = &HC0C0C0      'GREY


End Sub

Private Sub Trip_throttle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 7/12/00
Trip_throttleMouseDown = True

End Sub

Private Sub Trip_throttle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 7/12/00
Trip_throttleMouseDown = False

End Sub

Private Sub Trips_button_Click()


' 11/12/01
Throttle_Trips.ZOrder



End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
   
 ' 14/2/00 Block pump valve clicks on ACT Terminals
If Index > 10 And Index < 15 Then
  If Not I_AM_INSTRUCTOR And Not PTT Then
      ' I MUST be an ACT Terminal (Full Mission) so
      Exit Sub
  End If
End If
  
' 23/11/01  See Bilge X-Ref V8.xls
If Index = 590 Then
    Call valve_pcdi_LSS(590)
    Exit Sub
End If
   
' === 21/11/02 17.023 ===
If Index = 557 Then
    Call valve_pcdi_LSS(557)
    Exit Sub
End If

   
' 16/9/99  See Lubricating Oil System Valve X-Ref.xls
If Index <> 10 And Index <> 16 Then
    Call valve_pcdi_LSS(Index + 240)
End If


' OLD LSS Valves...
   
'Select Case index
'  Case 1
'   Call valve_pcdi_LSS(59)
'  Case 2
'   Call valve_pcdi_LSS(60)
'  Case 5
'   Call valve_pcdi_LSS(62)
'  Case 6
'   Call valve_pcdi_LSS(61)
'  Case 7
'   Call valve_pcdi_LSS(53)
'  Case 8
'   Call valve_pcdi_LSS(57)
'  Case 10
'   'Call valve_pcdi_LSS(48)
'  Case 11
'   Call valve_pcdi_LSS(47)
'  Case 12
'   Call valve_pcdi_LSS(49)
'  Case 13
'   Call valve_pcdi_LSS(50)
'  Case 14
'   Call valve_pcdi_LSS(52)
'  Case 16
'  'Call valve_pcdi_LSS(51)
'
'End Select

End Sub

Private Sub ZLO_header_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 124
Alarm_set_point.Show

End Sub

Private Sub ZLO_sump_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 130
Alarm_set_point.Show

End Sub
