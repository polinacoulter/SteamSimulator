VERSION 5.00
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.0#0"; "HavenPump v1_42.ocx"
Object = "{A6D9C597-48C2-11D1-8915-000021240899}#2.0#0"; "Haven3WayValve v1.41.ocx"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Begin VB.Form Lube_Oil 
   Caption         =   "Lubricatin Oil System"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   9165
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8115
      Left            =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   14314
      Begin VB.CommandButton Command5 
         Caption         =   "Main Bearings"
         Height          =   495
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6240
         Width           =   855
      End
      Begin VB.TextBox Text2 
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
         Left            =   7620
         TabIndex        =   33
         Text            =   "88 F"
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   25
         Text            =   "20 %"
         Top             =   540
         Width           =   615
      End
      Begin Haven3WayValvePrj.Haven3WayValve Haven3WayValve1 
         Height          =   495
         Left            =   800
         TabIndex        =   19
         Top             =   3180
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   873
         OpenSideTri     =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         Direction       =   3
      End
      Begin VB.CommandButton Pump_Start 
         BackColor       =   &H0000FF00&
         Caption         =   "Start"
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
         Index           =   0
         Left            =   1380
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton Pump_Start 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start"
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
         Left            =   1440
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6300
         Width           =   495
      End
      Begin VB.CommandButton Pump_Stop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stop"
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
         Left            =   1920
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6300
         Width           =   495
      End
      Begin VB.CommandButton Pump_Stop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stop"
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
         Left            =   1860
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton Pump_Stby 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stby"
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
         Index           =   0
         Left            =   2340
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton Pump_Stby 
         BackColor       =   &H0000FFFF&
         Caption         =   "Stby"
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
         Left            =   2400
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6300
         Width           =   495
      End
      Begin VB.TextBox Z_DOserv 
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
         Left            =   7620
         TabIndex        =   2
         Text            =   "20 %"
         Top             =   7380
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   915
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   6780
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1614
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   20
         OrientVertical  =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin PumpPrj.Pump Pump1 
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   4380
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin PumpPrj.Pump Pump1 
         Height          =   495
         Index           =   1
         Left            =   1860
         TabIndex        =   10
         Top             =   5760
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   10
         Left            =   1920
         TabIndex        =   11
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   11
         Left            =   1200
         TabIndex        =   12
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   12
         Left            =   2700
         TabIndex        =   13
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   13
         Left            =   1200
         TabIndex        =   14
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   14
         Left            =   2700
         TabIndex        =   15
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   16
         Left            =   1920
         TabIndex        =   16
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   1
         Left            =   900
         TabIndex        =   20
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   2
         Left            =   1920
         TabIndex        =   21
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   615
         Index           =   3
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   20
         OrientVertical  =   0   'False
         ColourClosed    =   65535
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   4800
         Picture         =   "Lube_Oil.frx":0000
         Stretch         =   -1  'True
         Top             =   6300
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   45
         X1              =   4740
         X2              =   5040
         Y1              =   6420
         Y2              =   6420
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "From Filling System"
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
         Left            =   5100
         TabIndex        =   35
         Top             =   6300
         Width           =   1515
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   44
         X1              =   4740
         X2              =   4740
         Y1              =   6840
         Y2              =   6420
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   7260
         Picture         =   "Lube_Oil.frx":030A
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image20 
         Height          =   255
         Left            =   7260
         Picture         =   "Lube_Oil.frx":0614
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Left            =   7200
         TabIndex        =   29
         Top             =   3120
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
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
         Left            =   4420
         Picture         =   "Lube_Oil.frx":091E
         Stretch         =   -1  'True
         Top             =   840
         Width           =   255
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   5400
         X2              =   5700
         Y1              =   5520
         Y2              =   5700
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   5400
         X2              =   5700
         Y1              =   4020
         Y2              =   3840
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   6240
         Picture         =   "Lube_Oil.frx":0C28
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   6780
         Picture         =   "Lube_Oil.frx":0F32
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   6600
         Picture         =   "Lube_Oil.frx":123C
         Stretch         =   -1  'True
         Top             =   4980
         Width           =   255
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   6780
         Picture         =   "Lube_Oil.frx":1546
         Stretch         =   -1  'True
         Top             =   4200
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
      Begin VB.Image Image15 
         Height          =   255
         Left            =   6780
         Picture         =   "Lube_Oil.frx":1850
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   255
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   5220
         Picture         =   "Lube_Oil.frx":1B5A
         Stretch         =   -1  'True
         Top             =   3060
         Width           =   255
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
      Begin VB.Image Image13 
         Height          =   255
         Left            =   6780
         Picture         =   "Lube_Oil.frx":1E64
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   6780
         Picture         =   "Lube_Oil.frx":216E
         Stretch         =   -1  'True
         Top             =   6300
         Width           =   255
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
      Begin VB.Image Image12 
         Height          =   255
         Left            =   5160
         Picture         =   "Lube_Oil.frx":2478
         Stretch         =   -1  'True
         Top             =   5580
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   8220
         Picture         =   "Lube_Oil.frx":2782
         Stretch         =   -1  'True
         Top             =   4260
         Width           =   255
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
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5100
         Top             =   2640
         Width           =   375
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   6480
         Top             =   2640
         Width           =   615
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   6480
         Top             =   4620
         Width           =   615
      End
      Begin VB.Shape Shape15 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5040
         Top             =   4620
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
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
         Left            =   7920
         TabIndex        =   32
         Top             =   3540
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape17 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   7860
         Top             =   3480
         Width           =   615
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   6480
         Picture         =   "Lube_Oil.frx":2A8C
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   5040
         Picture         =   "Lube_Oil.frx":2D96
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   7860
         Picture         =   "Lube_Oil.frx":30A0
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   4740
         Picture         =   "Lube_Oil.frx":33AA
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   5100
         Picture         =   "Lube_Oil.frx":36B4
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   6480
         Picture         =   "Lube_Oil.frx":39BE
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   255
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
         Left            =   5760
         TabIndex        =   31
         Top             =   4500
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
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
         Left            =   5700
         TabIndex        =   30
         Top             =   2460
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
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   5700
         X2              =   5700
         Y1              =   3840
         Y2              =   5700
      End
      Begin VB.Shape Shape16 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   7740
         Top             =   3600
         Width           =   915
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   5700
         X2              =   6480
         Y1              =   5700
         Y2              =   5340
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   5700
         X2              =   6480
         Y1              =   3840
         Y2              =   4080
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   6480
         X2              =   6480
         Y1              =   4080
         Y2              =   5340
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   5400
         X2              =   5400
         Y1              =   4020
         Y2              =   5520
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   6480
         X2              =   6480
         Y1              =   2040
         Y2              =   3300
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   5460
         X2              =   6480
         Y1              =   3060
         Y2              =   3300
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   5460
         X2              =   5460
         Y1              =   3060
         Y2              =   2340
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   5460
         X2              =   6480
         Y1              =   2340
         Y2              =   2040
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   2760
         Picture         =   "Lube_Oil.frx":3CC8
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   2340
         Picture         =   "Lube_Oil.frx":3FD2
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   255
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   4140
         X2              =   4140
         Y1              =   420
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   4140
         X2              =   4560
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   6960
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   4080
         X2              =   4560
         Y1              =   1740
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   4080
         X2              =   4080
         Y1              =   2100
         Y2              =   1740
      End
      Begin VB.Label Label8 
         Caption         =   "LO Cooler"
         Height          =   195
         Left            =   780
         TabIndex        =   28
         Top             =   1740
         Width           =   855
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   480
         Picture         =   "Lube_Oil.frx":42DC
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   255
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   25
         X1              =   1020
         X2              =   3240
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   24
         X1              =   1020
         X2              =   1020
         Y1              =   3840
         Y2              =   2580
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   2820
         TabIndex        =   27
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   23
         X1              =   3840
         X2              =   3840
         Y1              =   6840
         Y2              =   4140
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   3600
         Top             =   2100
         Width           =   615
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   22
         X1              =   3840
         X2              =   3840
         Y1              =   3600
         Y2              =   1320
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
      Begin VB.Shape Sight_Glass_shape 
         BorderWidth     =   3
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   3720
         Width           =   255
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   3
         Height          =   675
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   3540
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "LO Header Tank"
         Height          =   195
         Left            =   2520
         TabIndex        =   24
         Top             =   0
         Width           =   1455
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   18
         X1              =   2040
         X2              =   2040
         Y1              =   1920
         Y2              =   1320
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   17
         X1              =   7980
         X2              =   480
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   16
         X1              =   480
         X2              =   480
         Y1              =   3420
         Y2              =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "To Filling System"
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
         Left            =   60
         TabIndex        =   22
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   15
         X1              =   1080
         X2              =   660
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   14
         X1              =   840
         X2              =   480
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   720
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   720
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   720
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   720
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   720
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   13
         X1              =   3240
         X2              =   3240
         Y1              =   6000
         Y2              =   3840
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   600
         X2              =   600
         Y1              =   4920
         Y2              =   5340
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   12
         X1              =   780
         X2              =   3240
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   780
         X2              =   3240
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   4
         X1              =   780
         X2              =   780
         Y1              =   6000
         Y2              =   4620
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   5
         X1              =   600
         X2              =   780
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   6
         X1              =   1740
         X2              =   2460
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   8
         X1              =   1740
         X2              =   1740
         Y1              =   4620
         Y2              =   4200
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   9
         X1              =   1740
         X2              =   2460
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   10
         X1              =   2460
         X2              =   2460
         Y1              =   5520
         Y2              =   6000
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   11
         X1              =   1740
         X2              =   1740
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
         Left            =   900
         TabIndex        =   18
         Top             =   4920
         Width           =   435
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   7
         X1              =   2460
         X2              =   2460
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
         Left            =   960
         TabIndex        =   17
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
         Left            =   420
         Shape           =   3  'Circle
         Top             =   4920
         Width           =   195
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   3  'Circle
         Top             =   5100
         Width           =   195
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   420
         X2              =   420
         Y1              =   4920
         Y2              =   5340
      End
      Begin VB.Label Label1 
         Caption         =   "LO Sump"
         Height          =   195
         Left            =   2160
         TabIndex        =   1
         Top             =   7800
         Width           =   3435
      End
      Begin VB.Line Fuel_line 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   180
         X2              =   420
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
         Height          =   795
         Left            =   5100
         TabIndex        =   26
         ToolTipText     =   "Sight Glass will be yellow if there is oil flow to the sump"
         Top             =   360
         Width           =   1215
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
   End
End
Attribute VB_Name = "Lube_Oil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()

Main_Bearings.Show


End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idLube_Oil)

End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idLube_Oil)
    HavenStretch1.Visible = True
End If

End Sub
