VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Bunkers 
   Caption         =   "Bunkers"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   12345
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   8865
      Left            =   0
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   15637
      Begin VB.CommandButton Ballast_cmd 
         Caption         =   "Ballast"
         Height          =   675
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   10140
         TabIndex        =   79
         Top             =   4320
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F.O. Service System"
         Height          =   675
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox Pbunk_out 
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
         Left            =   3960
         TabIndex        =   57
         Text            =   "2.5 Bar"
         Top             =   4500
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Shore Supply"
         Height          =   915
         Left            =   3360
         TabIndex        =   52
         Top             =   60
         Width           =   1755
         Begin VB.TextBox TFO_fill 
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
            Left            =   960
            TabIndex        =   55
            Text            =   "77 F"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H000080FF&
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
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   540
            Width           =   315
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00004080&
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
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label21 
            Caption         =   "DO"
            Height          =   255
            Left            =   600
            TabIndex        =   81
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "HO"
            Height          =   255
            Left            =   600
            TabIndex        =   80
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.TextBox TDO_double_bottom 
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
         Left            =   6000
         TabIndex        =   47
         Text            =   "77 F"
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox TDO_double_bottom 
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
         Index           =   2
         Left            =   3780
         TabIndex        =   46
         Text            =   "77 F"
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox TDO_sett 
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
         Left            =   7140
         TabIndex        =   45
         Text            =   "77 F"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ZDO_double_bottom 
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
         Left            =   6000
         TabIndex        =   43
         Text            =   "85 %"
         Top             =   7500
         Width           =   615
      End
      Begin VB.TextBox ZDO_double_bottom 
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
         Index           =   2
         Left            =   3780
         TabIndex        =   42
         Text            =   "85 %"
         Top             =   7500
         Width           =   615
      End
      Begin VB.TextBox ZDO_sett 
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
         Left            =   7140
         TabIndex        =   41
         Text            =   "85 %"
         Top             =   2220
         Width           =   615
      End
      Begin VB.TextBox THO_sett 
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
         Index           =   1
         Left            =   5220
         TabIndex        =   40
         Text            =   "77 F"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ZHO_sett 
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
         Index           =   1
         Left            =   5220
         TabIndex        =   39
         Text            =   "85 %"
         Top             =   2220
         Width           =   615
      End
      Begin VB.TextBox THO_sett 
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
         Index           =   2
         Left            =   3540
         TabIndex        =   38
         Text            =   "77 F"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ZHO_sett 
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
         Index           =   2
         Left            =   3540
         TabIndex        =   37
         Text            =   "85 %"
         Top             =   2220
         Width           =   615
      End
      Begin VB.TextBox THO_deep 
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
         Index           =   1
         Left            =   8640
         TabIndex        =   36
         Text            =   "77 F"
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox ZHO_deep 
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
         Left            =   8640
         TabIndex        =   35
         Text            =   "85 %"
         Top             =   4620
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   0
         Top             =   5580
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
         Left            =   4380
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   6060
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
         Index           =   0
         Left            =   4380
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5220
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
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
         Left            =   3900
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   6060
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
         Left            =   3900
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5220
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
         Left            =   3420
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   6060
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
         Left            =   3420
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox THO_deep 
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
         Index           =   2
         Left            =   1140
         TabIndex        =   27
         Text            =   "77 F"
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox ZHO_deep 
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
         Index           =   2
         Left            =   1140
         TabIndex        =   26
         Text            =   "85 %"
         Top             =   4620
         Width           =   615
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   4
         Left            =   6480
         TabIndex        =   4
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump Havenpump1 
         Height          =   405
         Index           =   1
         Left            =   3690
         TabIndex        =   2
         Top             =   6420
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         CentrifugalPump =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   615
         Index           =   0
         Left            =   2540
         TabIndex        =   0
         Top             =   6720
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   1085
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
         ValveType       =   3
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   615
         Index           =   1
         Left            =   7410
         TabIndex        =   1
         Top             =   6720
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   1085
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
         ValveType       =   3
      End
      Begin PumpPrj.Pump Havenpump1 
         Height          =   405
         Index           =   0
         Left            =   3660
         TabIndex        =   3
         Top             =   5610
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         CentrifugalPump =   0   'False
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   10
         Left            =   3090
         TabIndex        =   5
         Top             =   5460
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   12
         Left            =   3150
         TabIndex        =   6
         Top             =   6300
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   9
         Left            =   4260
         TabIndex        =   7
         Top             =   5490
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   11
         Left            =   4260
         TabIndex        =   8
         Top             =   6330
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   6
         Left            =   3000
         TabIndex        =   9
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   5
         Left            =   4680
         TabIndex        =   10
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   7
         Left            =   1500
         TabIndex        =   12
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   2
         Left            =   9510
         TabIndex        =   14
         Top             =   840
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         BottomColour    =   65280
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   8
         Left            =   450
         TabIndex        =   18
         Top             =   840
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         OrientVertical  =   0   'False
         EnabledBtmRight =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   465
         Index           =   3
         Left            =   8340
         TabIndex        =   11
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   820
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   13
         Left            =   7200
         TabIndex        =   60
         Top             =   5640
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   14
         Left            =   6960
         TabIndex        =   61
         Top             =   5640
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   15
         Left            =   6720
         TabIndex        =   62
         Top             =   5640
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   16
         Left            =   6480
         TabIndex        =   63
         Top             =   5640
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   17
         Left            =   6240
         TabIndex        =   64
         Top             =   5640
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   18
         Left            =   6000
         TabIndex        =   65
         Top             =   5640
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   19
         Left            =   3960
         TabIndex        =   66
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 valve 
         Height          =   345
         Index           =   20
         Left            =   2640
         TabIndex        =   67
         Top             =   3960
         Visible         =   0   'False
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   609
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         BackColour      =   12632256
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   705
         Left            =   330
         OleObjectBlob   =   "Bunker Tanks.frx":0000
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\BilgeHull_Port_16.vsd"
         TabIndex        =   84
         Top             =   7680
         Width           =   495
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   720
         Left            =   9360
         OleObjectBlob   =   "Bunker Tanks.frx":2418
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\BilgeHull_Stbd_16.vsd"
         TabIndex        =   83
         Top             =   7680
         Width           =   690
      End
      Begin VB.Line Thick_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   36
         Index           =   1
         X1              =   9720
         X2              =   10260
         Y1              =   8640
         Y2              =   7980
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   720
         Left            =   6060
         OleObjectBlob   =   "Bunker Tanks.frx":4430
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Bilge_Hull_Stbd_16.vsd"
         TabIndex        =   78
         Top             =   8040
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   645
         Left            =   2820
         OleObjectBlob   =   "Bunker Tanks.frx":6E48
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\STEAM_SIM\Images\Bilge_Hull_port_16.vsd"
         TabIndex        =   82
         Top             =   8100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Thick_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   32
         Index           =   0
         X1              =   600
         X2              =   240
         Y1              =   8640
         Y2              =   8100
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   420
         Left            =   3120
         Shape           =   3  'Circle
         Top             =   4470
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   419
         Left            =   3480
         Shape           =   3  'Circle
         Top             =   4470
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   413
         Left            =   6180
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   255
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   420
         Left            =   3110
         TabIndex        =   76
         Top             =   4920
         Width           =   255
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   419
         Left            =   3440
         TabIndex        =   75
         Top             =   4920
         Width           =   255
      End
      Begin VB.Line Line51 
         BorderWidth     =   2
         X1              =   4440
         X2              =   4440
         Y1              =   6840
         Y2              =   7020
      End
      Begin VB.Line Line42 
         BorderWidth     =   2
         X1              =   4800
         X2              =   4800
         Y1              =   6840
         Y2              =   7020
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   418
         Left            =   4320
         TabIndex        =   74
         Top             =   6840
         Width           =   255
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   417
         Left            =   4680
         TabIndex        =   73
         Top             =   6840
         Width           =   255
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   416
         Left            =   5040
         TabIndex        =   72
         Top             =   6840
         Width           =   255
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   415
         Left            =   5480
         TabIndex        =   71
         Top             =   6840
         Width           =   255
      End
      Begin VB.Line Line39 
         BorderWidth     =   2
         X1              =   6000
         X2              =   5920
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line36 
         BorderWidth     =   2
         X1              =   6000
         X2              =   6000
         Y1              =   3840
         Y2              =   4320
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   75
         Index           =   414
         Left            =   5820
         TabIndex        =   70
         Top             =   6840
         Width           =   255
      End
      Begin VB.Line Line35 
         BorderWidth     =   2
         X1              =   5600
         X2              =   5600
         Y1              =   6840
         Y2              =   7020
      End
      Begin VB.Label VV_label 
         BackColor       =   &H00000000&
         Height          =   70
         Index           =   413
         Left            =   6170
         TabIndex        =   69
         Top             =   6840
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   418
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   417
         Left            =   4680
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   416
         Left            =   5040
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   415
         Left            =   5460
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   255
      End
      Begin VB.Shape VV_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Index           =   414
         Left            =   5820
         Shape           =   3  'Circle
         Top             =   7020
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Manifold"
         Height          =   255
         Left            =   6180
         TabIndex        =   59
         Top             =   6120
         Width           =   1395
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "**"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6780
         TabIndex        =   58
         Top             =   6780
         Width           =   435
      End
      Begin VB.Line Line50 
         X1              =   6720
         X2              =   6480
         Y1              =   6960
         Y2              =   7020
      End
      Begin VB.Line Line49 
         X1              =   6720
         X2              =   6720
         Y1              =   6960
         Y2              =   6600
      End
      Begin VB.Label Manifold_fuel 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "HO/ DO mixture"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   56
         ToolTipText     =   "Suck the fuel out before changing from HO to DO and vice versa"
         Top             =   6360
         Width           =   1395
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "HO/ DO mixture"
         Height          =   255
         Left            =   6720
         TabIndex        =   51
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "DO"
         Height          =   255
         Left            =   6720
         TabIndex        =   50
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "HO"
         Height          =   255
         Left            =   6720
         TabIndex        =   49
         Top             =   300
         Width           =   555
      End
      Begin VB.Shape DO_oil 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6000
         Top             =   780
         Width           =   675
      End
      Begin VB.Shape HO_oil 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6000
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tank Colours"
         Height          =   255
         Left            =   5760
         TabIndex        =   48
         Top             =   60
         Width           =   1095
      End
      Begin VB.Shape HO_DO_oil 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6000
         Top             =   540
         Width           =   675
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   3600
         X2              =   7080
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   7140
         X2              =   8100
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   2220
         X2              =   6960
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Shape Green_shape 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   4800
         Shape           =   3  'Circle
         Top             =   4140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape RED_shape 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   4140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line48 
         BorderWidth     =   2
         Index           =   7
         X1              =   9960
         X2              =   9960
         Y1              =   1740
         Y2              =   7440
      End
      Begin VB.Line Line48 
         BorderWidth     =   3
         Index           =   6
         X1              =   420
         X2              =   420
         Y1              =   840
         Y2              =   7500
      End
      Begin VB.Line Line48 
         BorderWidth     =   3
         Index           =   5
         X1              =   720
         X2              =   9480
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   2
         X1              =   8700
         X2              =   8820
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   1
         X1              =   8700
         X2              =   8580
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   7
         X1              =   9300
         X2              =   9300
         Y1              =   6960
         Y2              =   7380
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   6
         X1              =   9300
         X2              =   9120
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   5
         X1              =   9000
         X2              =   9120
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   3
         X1              =   9000
         X2              =   8820
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   8400
         X2              =   8580
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   4
         X1              =   8400
         X2              =   8400
         Y1              =   6960
         Y2              =   7380
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "FO Transfer Pumps"
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
         Left            =   4020
         TabIndex        =   34
         Top             =   4920
         Width           =   1755
      End
      Begin VB.Line Line40 
         BorderWidth     =   2
         X1              =   3570
         X2              =   3570
         Y1              =   4770
         Y2              =   4980
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   31
         X1              =   5460
         X2              =   5460
         Y1              =   2880
         Y2              =   3300
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   30
         X1              =   5460
         X2              =   5280
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   29
         X1              =   5160
         X2              =   5280
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   28
         X1              =   4560
         X2              =   4560
         Y1              =   2880
         Y2              =   3300
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   27
         X1              =   4560
         X2              =   4740
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   26
         X1              =   4860
         X2              =   4740
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   25
         X1              =   4860
         X2              =   4980
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   24
         X1              =   5160
         X2              =   4980
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   23
         X1              =   3840
         X2              =   3840
         Y1              =   2880
         Y2              =   3300
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   22
         X1              =   3840
         X2              =   3660
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   21
         X1              =   3540
         X2              =   3660
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   20
         X1              =   2940
         X2              =   2940
         Y1              =   2880
         Y2              =   3300
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   19
         X1              =   2940
         X2              =   3120
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   18
         X1              =   3240
         X2              =   3120
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   17
         X1              =   3240
         X2              =   3360
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   16
         X1              =   3540
         X2              =   3360
         Y1              =   2880
         Y2              =   3180
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   15
         X1              =   1500
         X2              =   1620
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   14
         X1              =   1200
         X2              =   1080
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   13
         X1              =   1800
         X2              =   1800
         Y1              =   6960
         Y2              =   7380
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   12
         X1              =   1800
         X2              =   1620
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   11
         X1              =   900
         X2              =   900
         Y1              =   6960
         Y2              =   7380
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   10
         X1              =   900
         X2              =   1080
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   9
         X1              =   1200
         X2              =   1320
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Line steam_line 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   8
         X1              =   1500
         X2              =   1320
         Y1              =   6960
         Y2              =   7260
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   4140
         Picture         =   "Bunker Tanks.frx":9660
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   255
      End
      Begin VB.Line Line32 
         BorderWidth     =   2
         X1              =   6000
         X2              =   6000
         Y1              =   1140
         Y2              =   3480
      End
      Begin VB.Line Line47 
         BorderWidth     =   2
         X1              =   9120
         X2              =   9120
         Y1              =   2040
         Y2              =   1560
      End
      Begin VB.Line Line46 
         BorderWidth     =   2
         X1              =   6240
         X2              =   6240
         Y1              =   1920
         Y2              =   1500
      End
      Begin VB.Line Line45 
         BorderWidth     =   2
         X1              =   5460
         X2              =   5460
         Y1              =   1980
         Y2              =   1560
      End
      Begin VB.Line Line44 
         BorderWidth     =   2
         X1              =   3780
         X2              =   3780
         Y1              =   1980
         Y2              =   1560
      End
      Begin VB.Line Line29 
         BorderWidth     =   2
         X1              =   1200
         X2              =   1200
         Y1              =   2040
         Y2              =   1560
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PORT Bunkering Station"
         Height          =   345
         Left            =   300
         TabIndex        =   22
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "STBD. Bunkering Station"
         Height          =   465
         Left            =   8370
         TabIndex        =   23
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "# 2"
         Height          =   315
         Left            =   3060
         TabIndex        =   25
         Top             =   6060
         Width           =   405
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "# 1"
         Height          =   315
         Left            =   3120
         TabIndex        =   24
         Top             =   5250
         Width           =   405
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   3240
         Picture         =   "Bunker Tanks.frx":996A
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   255
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   7020
         Picture         =   "Bunker Tanks.frx":9C74
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   255
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   5040
         Picture         =   "Bunker Tanks.frx":9F7E
         Stretch         =   -1  'True
         Top             =   6300
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   6780
         Picture         =   "Bunker Tanks.frx":A288
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   4740
         Picture         =   "Bunker Tanks.frx":A592
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   3480
         Picture         =   "Bunker Tanks.frx":A89C
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   255
      End
      Begin VB.Image Arrow_left 
         Height          =   255
         Left            =   7560
         Picture         =   "Bunker Tanks.frx":ABA6
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   3120
         Picture         =   "Bunker Tanks.frx":AEB0
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   5790
         Picture         =   "Bunker Tanks.frx":B1BA
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D.O.  Settling  Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   6390
         TabIndex        =   21
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STBD H.O.  Settling  Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   4500
         TabIndex        =   20
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PORT H.O.  Settling  Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   2880
         TabIndex        =   19
         Top             =   1920
         Width           =   795
      End
      Begin VB.Line Line41 
         BorderWidth     =   2
         X1              =   5940
         X2              =   5940
         Y1              =   6840
         Y2              =   7020
      End
      Begin VB.Line Line37 
         BorderWidth     =   2
         X1              =   3240
         X2              =   3240
         Y1              =   4770
         Y2              =   4980
      End
      Begin VB.Line Line38 
         BorderWidth     =   2
         X1              =   6300
         X2              =   6300
         Y1              =   6870
         Y2              =   6990
      End
      Begin VB.Line Line33 
         BorderWidth     =   2
         X1              =   5920
         X2              =   5920
         Y1              =   4320
         Y2              =   7020
      End
      Begin VB.Line Line31 
         BorderWidth     =   2
         X1              =   7080
         X2              =   7080
         Y1              =   1500
         Y2              =   3720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STBD.  D.O. Double Bottom Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   7230
         TabIndex        =   17
         Top             =   7560
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PORT   D.O. Double Bottom Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   1980
         TabIndex        =   16
         Top             =   7500
         Width           =   1605
      End
      Begin VB.Line Line28 
         BorderWidth     =   2
         X1              =   4980
         X2              =   5160
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   2640
         X2              =   2880
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   2670
         X2              =   2670
         Y1              =   6180
         Y2              =   4620
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STBD. H.O. Deep Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   8280
         TabIndex        =   15
         Top             =   2340
         Width           =   1425
      End
      Begin VB.Line Line25 
         BorderWidth     =   2
         X1              =   840
         X2              =   9570
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PORT H.O. Deep Tank"
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   780
         TabIndex        =   13
         Top             =   2340
         Width           =   1245
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   2
         Height          =   165
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   2010
         Width           =   165
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   1200
         X2              =   2190
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   2
         Height          =   165
         Left            =   9030
         Shape           =   3  'Circle
         Top             =   2010
         Width           =   165
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   8100
         X2              =   9120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape20 
         BorderWidth     =   2
         Height          =   165
         Left            =   5370
         Shape           =   3  'Circle
         Top             =   1980
         Width           =   165
      End
      Begin VB.Line Line22 
         BorderWidth     =   2
         X1              =   4440
         X2              =   5460
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape19 
         BorderWidth     =   2
         Height          =   165
         Left            =   3690
         Shape           =   3  'Circle
         Top             =   1980
         Width           =   165
      End
      Begin VB.Line Line21 
         BorderWidth     =   2
         X1              =   2760
         X2              =   3780
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape18 
         BorderWidth     =   2
         Height          =   165
         Left            =   6150
         Shape           =   3  'Circle
         Top             =   1950
         Width           =   165
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   6240
         X2              =   7080
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   8100
         X2              =   8100
         Y1              =   3570
         Y2              =   1560
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   4440
         X2              =   4440
         Y1              =   3600
         Y2              =   1560
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   2760
         X2              =   2760
         Y1              =   3570
         Y2              =   1560
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   2220
         X2              =   2220
         Y1              =   3570
         Y2              =   1560
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   3240
         X2              =   3240
         Y1              =   4470
         Y2              =   3600
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   3600
         X2              =   3600
         Y1              =   4470
         Y2              =   3720
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   2880
         Top             =   4470
         Width           =   1005
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   2880
         X2              =   4980
         Y1              =   6630
         Y2              =   6660
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   2880
         X2              =   4980
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   2640
         X2              =   2880
         Y1              =   6180
         Y2              =   6180
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   4980
         X2              =   4980
         Y1              =   5820
         Y2              =   6660
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   2880
         X2              =   2880
         Y1              =   5820
         Y2              =   6630
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   5160
         X2              =   5160
         Y1              =   7020
         Y2              =   6120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   6480
         X2              =   7410
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   2940
         X2              =   4260
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   5550
         X2              =   5550
         Y1              =   7350
         Y2              =   8190
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   4830
         X2              =   4830
         Y1              =   7320
         Y2              =   8160
      End
      Begin VB.Line Line1 
         X1              =   450
         X2              =   9960
         Y1              =   7380
         Y2              =   7380
      End
      Begin VB.Shape HOdeep 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   420
         Top             =   7200
         Width           =   2115
      End
      Begin VB.Shape HOdeep_full 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   5655
         Index           =   2
         Left            =   420
         Top             =   1740
         Width           =   2115
      End
      Begin VB.Shape DOsett 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   6190
         Top             =   3120
         Width           =   1650
      End
      Begin VB.Shape DOsett_full 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1575
         Left            =   6180
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Shape HOdeep 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   7860
         Top             =   7200
         Width           =   2100
      End
      Begin VB.Shape DObottom 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   480
         Top             =   8160
         Width           =   4695
      End
      Begin VB.Shape DObottom_full 
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   915
         Index           =   2
         Left            =   420
         Top             =   7380
         Width           =   4755
      End
      Begin VB.Shape DObottom 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   5220
         Top             =   8160
         Width           =   4695
      End
      Begin VB.Shape DObottom_full 
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   915
         Index           =   1
         Left            =   5160
         Top             =   7380
         Width           =   4800
      End
      Begin VB.Shape HOsett 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   2540
         Top             =   3120
         Width           =   1650
      End
      Begin VB.Shape HOsett_full 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   2
         Left            =   2520
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Shape HOsett 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   4215
         Top             =   3120
         Width           =   1650
      End
      Begin VB.Shape HOsett_full 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   1
         Left            =   4200
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Shape HOdeep_full 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   5655
         Index           =   1
         Left            =   7860
         Top             =   1740
         Width           =   2100
      End
      Begin VB.Label Label20 
         Height          =   255
         Left            =   8280
         TabIndex        =   68
         Top             =   6000
         Width           =   255
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   4260
         Top             =   7005
         Width           =   2265
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      TabIndex        =   44
      Top             =   240
      Width           =   435
   End
End
Attribute VB_Name = "Bunkers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub Pump_Stop_Click(Index As Integer)

End Sub

Private Sub Ballast_cmd_Click()

' === 30/05/02 ===
Ballast_BULK.Show

End Sub

Private Sub Command1_Click()

set_dummy_focus = True

status_flag = True

' === 22/11/02 LEAVE AS IS (Instructor ONLY)===
Call set_STEAM_logicals(7, status_flag)

End Sub

Private Sub Command2_Click()

set_dummy_focus = True

status_flag = False
' === 22/11/02 LEAVE AS IS (Instructor ONLY)===
Call set_STEAM_logicals(7, status_flag)

End Sub

Private Sub Command3_Click()

FO_Service.Show

End Sub

Private Sub Command4_Click()


End Sub

Private Sub Form_Load()

Dim i As Integer

On Error Resume Next
Call JYAScreenInit(Me, idBunkers)


old_HO_deep(1) = -123
old_HO_deep(2) = -123
OLD_DO_double_bottom(1) = -123
OLD_DO_double_bottom(2) = -123

old_HO_sett(1) = -123
old_HO_sett(2) = -123
old_DO_sett = -123

set_dummy_focus = True

' 26/10/99 To ensure an update...
For i = 413 To 420
  OLD_vv_status_LSS(i) = -123
Next

' === 31/05/02 ===
If Ballast_IS_included Then
    Ballast_cmd.Visible = True
End If


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)



End Sub

Private Sub Form_Resize()
Dim i As Integer

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idBunkers)
    HavenStretch1.Visible = True
End If

' Fxshape1( ) Smudges when re-sized , so ensure an up-date
For i = 164 To 171
    OLD_vv_status_LSS(i) = -1
Next i

old_HO_deep(1) = -123
old_HO_deep(2) = -123
OLD_DO_double_bottom(1) = -123
OLD_DO_double_bottom(2) = -123

old_HO_sett(1) = -123
old_HO_sett(2) = -123
old_DO_sett = -123

End Sub







Private Sub FXShape1_Click(Index As Integer)

Control_PC_check
If This_Control_status = False Then Exit Sub

   
   
' 20/9/99  See Bunker_Valves_XrefXref.xls

    Call valve_pcdi_LSS(Index)

   
' 22/9/99  OLD LSS Valves....
'Select Case Index
'  Case 0
'    Call valve_pcdi_LSS(164)
'  Case 1
'    Call valve_pcdi_LSS(165)
'  Case 2
'    Call valve_pcdi_LSS(166)
'  Case 3
'    Call valve_pcdi_LSS(167)
'  Case 4
'    Call valve_pcdi_LSS(168)
'  Case 5
'    Call valve_pcdi_LSS(169)
'  Case 6
'    Call valve_pcdi_LSS(170)
'  Case 7
'    Call valve_pcdi_LSS(171)
'End Select

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(4, 21))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(4, 22))
'End If
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(4, 4))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(4, 16))
End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(2, 21))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(2, 22))
'End If

If Index = 0 Then
   Call setPCDI(PumpIO_LSS(2, 4))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(2, 16))
End If

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(5=Stand-bye,nPump)
'If Index = 0 Then
'   Call setPCDI(PumpIO_LSS(5, 21))
'ElseIf Index = 1 Then
'   Call setPCDI(PumpIO_LSS(5, 22))
'End If
If Index = 0 Then
   Call setPCDI(PumpIO_LSS(5, 4))
ElseIf Index = 1 Then
   Call setPCDI(PumpIO_LSS(5, 16))
End If

End Sub

Private Sub THO_deep_Click(Index As Integer)

'6/1/00

    Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_HI = 26  ' Tank # 1 STBD
Else
    nAlarm_HI = 27  ' Tank # 2 PORT
End If

    Alarm_set_point.Show

End Sub

Private Sub THO_sett_Click(Index As Integer)

'6/1/00

    Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_HI = 28  ' Tank # 1 STBD
    nAlarm_LO = 36  ' Tank # 1 STBD
Else
    nAlarm_HI = 29  ' Tank # 2 PORT
    nAlarm_LO = 37  ' Tank # 2 PORT
End If

    Alarm_set_point.Show
    
End Sub

Private Sub Timer1_Timer()


Dim i As Integer
Dim single_char As String * 1
Dim valve_number As Integer



' 27/11/99

' 27/7/00
'List1.clear
'For i = 383 To 386
'    If D_INPUT(i) Then
'        List1.AddItem "DI(" & i & ") = " & "TRUE"
'    Else
'        List1.AddItem "DI(" & i & ") = " & "FALSE"
'    End If
'Next i
'For i = 400 To 403
'    If D_INPUT(i) Then
'        List1.AddItem "DI(" & i & ") = " & "TRUE"
'    Else
'        List1.AddItem "DI(" & i & ") = " & "FALSE"
'    End If
'Next i


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If


' === 21/01/02 ===
If VB_steam_logicals(7) Then
    Command1.BackColor = &H4080&         ' dark brown
    Command2.BackColor = &HC0C0C0        ' grey
Else
    Command1.BackColor = &HC0C0C0        ' grey
    Command2.BackColor = &H80FF&         ' light brown
End If


' 7/1/00
Call paint_alarm(Me.THO_deep(1), 26, 0, 0, 0)
Call paint_alarm(Me.THO_deep(2), 27, 0, 0, 0)

Call paint_alarm(Me.THO_sett(1), 28, 36, 0, 0)
Call paint_alarm(Me.THO_sett(2), 29, 37, 0, 0)

Call paint_alarm(Me.ZHO_sett(1), 44, 0, 0, 0)
Call paint_alarm(Me.ZHO_sett(2), 45, 0, 0, 0)

' 12/8/00
'Call paint_alarm(Me.ZDO_sett, 43, 0, 0, 0)
Call paint_alarm(Me.ZDO_sett, 43, 47, 0, 0)

' 27/7/99
For i = 1 To 2

    If VB_steam_integers(25 + i) = 0 Then
        HOdeep(i).FillColor = DO_oil.FillColor
    ElseIf VB_steam_integers(25 + i) = 1 Then
        HOdeep(i).FillColor = HO_oil.FillColor
    Else
        HOdeep(i).FillColor = HO_DO_oil.FillColor
    End If
    
    If VB_steam_integers(27 + i) = 0 Then
        HOsett(i).FillColor = DO_oil.FillColor
    ElseIf VB_steam_integers(27 + i) = 1 Then
        HOsett(i).FillColor = HO_oil.FillColor
    Else
        HOsett(i).FillColor = HO_DO_oil.FillColor
    End If
    
    If VB_steam_integers(29 + i) = 0 Then
        DObottom(i).FillColor = DO_oil.FillColor
    ElseIf VB_steam_integers(29 + i) = 1 Then
        DObottom(i).FillColor = HO_oil.FillColor
    Else
        DObottom(i).FillColor = HO_DO_oil.FillColor
    End If
    
Next i

    If VB_steam_integers(32) = 0 Then
        DOsett.FillColor = DO_oil.FillColor
    ElseIf VB_steam_integers(32) = 1 Then
        DOsett.FillColor = HO_oil.FillColor
    Else
        DOsett.FillColor = HO_DO_oil.FillColor
    End If
    
    If VB_steam_integers(24) = 0 Then
        Manifold_fuel.BackColor = DO_oil.FillColor
        Manifold_fuel.ForeColor = vbBlack
        Manifold_fuel.Caption = "DO"
    ElseIf VB_steam_integers(24) = 1 Then
        Manifold_fuel.BackColor = HO_oil.FillColor
        Manifold_fuel.ForeColor = vbWhite
        Manifold_fuel.Caption = "HO"
    Else
        Manifold_fuel.BackColor = HO_DO_oil.FillColor
        Manifold_fuel.ForeColor = vbWhite
       Manifold_fuel.Caption = "HO/DO mixture"
    End If
    


TFO_fill.text = Format(LSS_reals(253) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

THO_deep(1).text = Format(LSS_reals(254) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
THO_deep(2).text = Format(LSS_reals(255) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

TDO_double_bottom(1).text = Format(LSS_reals(256) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TDO_double_bottom(2).text = Format(LSS_reals(257) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

THO_sett(1).text = Format(LSS_reals(258) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
THO_sett(2).text = Format(LSS_reals(259) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

TDO_sett.text = Format(LSS_reals(260) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

Pbunk_out.text = Format(LSS_reals(261) * Press_mult, "##0.0") & Press_chars


'          PC_REALS(244) = FHO_deep_IN(1)
'          PC_REALS(245) = FHO_deep_IN(2)
'          PC_REALS(246) = FHO_sett_IN(1)
'          PC_REALS(247) = FHO_sett_IN(2)
'          PC_REALS(248) = FDO_sett_IN
'          PC_REALS(249) = ZHO_sett(1) * 100#
'          PC_REALS(250) = ZHO_sett(2) * 100#
'          PC_REALS(251) = ZDO_sett * 100#


' TESTING
'List1.Clear
'List1.AddItem "Zmanifold " & LSS_reals(237)
'List1.AddItem "Pmanifold " & LSS_reals(238)
'List1.AddItem "FFO_fill " & LSS_reals(239)
'List1.AddItem "FHO_deep(1) " & LSS_reals(242)
'List1.AddItem "FFO_double_bottom(1) " & LSS_reals(243)
'List1.AddItem " "
'List1.AddItem "FHO_deep_IN(1) " & LSS_reals(244)
'List1.AddItem "FHO_deep_IN(2) " & LSS_reals(245)
'List1.AddItem "FHO_sett_IN(1) " & LSS_reals(246)
'List1.AddItem "FHO_sett_IN(2) " & LSS_reals(247)
'List1.AddItem "FDO_sett_IN " & LSS_reals(248)
'List1.AddItem " "
'List1.AddItem "T_manifold " & LSS_reals(252)
'List1.AddItem "Fuel_manifold " & VB_steam_integers(24)
         

'If old_HO_deep(2) <> VScroll1.Value Then
'    old_HO_deep(2) = VScroll1.Value
'    Shape27.Height = VScroll1.Value * 0.01 * Shape26.Height
'    Shape27.Top = Shape26.Top + Shape26.Height - Shape27.Height
'End If


' HO Settling tanks
ZHO_sett(2).text = Format(LSS_reals(250), "##0.0") & " %"
HO_sett(2) = LSS_reals(250) * 10#
ZHO_sett(1).text = Format(LSS_reals(249), "##0.0") & " %"
HO_sett(1) = LSS_reals(249) * 10#
For i = 1 To 2
  If old_HO_sett(i) <> HO_sett(i) Then
    old_HO_sett(i) = HO_sett(i)
    HOsett(i).Height = HO_sett(i) * 0.001 * HOsett_full(i).Height
    HOsett(i).Top = HOsett_full(i).Top + HOsett_full(i).Height - HOsett(i).Height
  End If
Next i

' HO Deep tanks
ZHO_deep(2).text = Format(LSS_reals(236), "##0.0") & " %"
HO_deep(2) = LSS_reals(236) * 10#
ZHO_deep(1).text = Format(LSS_reals(235), "##0.0") & " %"
HO_deep(1) = LSS_reals(235) * 10#
For i = 1 To 2
  If old_HO_deep(i) <> HO_deep(i) Then
    old_HO_deep(i) = HO_deep(i)
    HOdeep(i).Height = HO_deep(i) * 0.001 * HOdeep_full(i).Height
    HOdeep(i).Top = HOdeep_full(i).Top + HOdeep_full(i).Height - HOdeep(i).Height
  End If
Next i


' DO Double Bottom Tanks
ZDO_double_bottom(2).text = Format(LSS_reals(241), "##0.0") & " %"
DO_double_bottom(2) = LSS_reals(241) * 10#
ZDO_double_bottom(1).text = Format(LSS_reals(240), "##0.0") & " %"
DO_double_bottom(1) = LSS_reals(240) * 10#
For i = 1 To 2
  If DO_double_bottom(i) <> OLD_DO_double_bottom(i) Then
    OLD_DO_double_bottom(i) = DO_double_bottom(i)
    DObottom(i).Height = DO_double_bottom(i) * 0.001 * DObottom_full(i).Height
    DObottom(i).Top = DObottom_full(i).Top + DObottom_full(i).Height - DObottom(i).Height
  End If
Next i

' DO Settling Tank
ZDO_sett.text = Format(LSS_reals(251), "##0.0") & " %"
DO_sett = LSS_reals(251) * 10#
If DO_sett <> old_DO_sett Then
    old_DO_sett = DO_sett
    DOsett.Height = DO_sett * 0.001 * DOsett_full.Height
    DOsett.Top = DOsett_full.Top + DOsett_full.Height - DOsett.Height
End If


' Note: update_pump_LSS(syxxx,index,nPump)
'Call update_pump_LSS(Bunkers, 0, 21)      ' FO Transfer Pump # 1
'Call update_pump_LSS(Bunkers, 1, 22)      ' FO Transfer Pump # 2
Call update_pump_LSS(Bunkers, 0, 4)      ' FO Transfer Pump # 1
Call update_pump_LSS(Bunkers, 1, 16)      ' FO Transfer Pump # 2



' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)

' OLD LSS Valves
'For i = 0 To 12
'    Call update_vvalve_LSS(Bunkers, 149 + i, i, -1)
'Next i

For i = 0 To 12
   Call update_vvalve_LSS(Me, 370 + i, i, -1)
Next i

  ' Call update_vvalve_LSS(Me, 413, 13, -1)


'For i = 0 To 7
For i = 413 To 420

  ' NEW Valve numbers
 '   valve_number = 164 + i
     valve_number = i
   
     '   single_char = Mid$(PC_INT_STRING, valve_number, 1)

    If valve_number < 251 Then

        single_char = Mid$(PC_INT_STRING, valve_number, 1)

        If single_char = "0" Then
            vv_status_LSS(valve_number) = 0
        ElseIf single_char = "1" Then
            vv_status_LSS(valve_number) = 1
        ElseIf single_char = "2" Then
            vv_status_LSS(valve_number) = 2
        Else
            vv_status_LSS(valve_number) = 3
        End If
    
    Else
            ' 17/12/98  EXTRA valves added from 251 to 600
            vv_status_LSS(valve_number) = VB_VV_extras(valve_number - 250)
    End If


    'If single_char = "0" Then
    '    vv_status_LSS(valve_number) = 0
    'ElseIf single_char = "3" Then
    '    vv_status_LSS(valve_number) = 3
    'Else
    '    vv_status_LSS(valve_number) = 1
    'End If
    
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
   
If Index < 13 Then
    Call valve_pcdi_LSS(Index + 370)
Else
    Call valve_pcdi_LSS(Index + 400)
End If

   
   
' OLD LSS  Valves  REMOVED


End Sub

Private Sub VV_lbl_Click(Index As Integer)

End Sub

Private Sub VV_label_Click(Index As Integer)

    Call valve_pcdi_LSS(Index)

End Sub

Private Sub ZDO_sett_Click()

'6/1/00

    Unload Alarm_set_point   ' Do this first

    nAlarm_HI = 43  ' DO Settling Tank
    
    ' 12/8/00
    nAlarm_LO = 47  ' DO Settling Tank

    Alarm_set_point.Show
    
    
End Sub

Private Sub ZHO_sett_Click(Index As Integer)

'6/1/00

    Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_LO = 44  ' Tank # 1 STBD
Else
    nAlarm_LO = 45  ' Tank # 2 PORT
End If

    Alarm_set_point.Show
    
End Sub
