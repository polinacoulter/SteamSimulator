VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Begin VB.Form MEcontrol 
   Caption         =   "Main Engine Control (MECD)"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10380
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6675
      Left            =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11774
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   7080
         TabIndex        =   71
         Top             =   5040
         Width           =   2115
         Begin VB.CommandButton Command3 
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
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   315
         End
         Begin VB.CommandButton Command4 
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
            Left            =   240
            TabIndex        =   72
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MAIN ENGINE"
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
            Left            =   720
            MousePointer    =   1  'Arrow
            TabIndex        =   75
            Top             =   780
            Width           =   1230
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "POWER OVERVIEW"
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
            Left            =   720
            MousePointer    =   1  'Arrow
            TabIndex        =   74
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   1680
         TabIndex        =   70
         Text            =   "Text6"
         Top             =   6300
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1740
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   6000
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1680
         TabIndex        =   68
         Text            =   "Text4"
         Top             =   5700
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clutch OUT"
         Height          =   550
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   5760
         Width           =   981
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   0
         TabIndex        =   66
         Text            =   "Text3"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Frame REV_limiter 
         Caption         =   "Rev Limiter"
         Height          =   3043
         Left            =   7860
         TabIndex        =   54
         Top             =   1500
         Width           =   1102
         Begin VB.Shape Shape3 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   180
            Top             =   900
            Width           =   495
         End
         Begin VB.Line Line12 
            BorderColor     =   &H00000000&
            X1              =   420
            X2              =   420
            Y1              =   360
            Y2              =   1560
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   32
            X1              =   180
            X2              =   660
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   31
            X1              =   180
            X2              =   660
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   30
            X1              =   180
            X2              =   660
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   29
            X1              =   180
            X2              =   660
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   28
            X1              =   180
            X2              =   660
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   27
            X1              =   180
            X2              =   660
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   26
            X1              =   180
            X2              =   660
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   25
            X1              =   180
            X2              =   660
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   24
            X1              =   180
            X2              =   660
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   23
            X1              =   180
            X2              =   660
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   22
            X1              =   180
            X2              =   660
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   65
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   64
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   780
            TabIndex        =   63
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "3"
            Height          =   195
            Index           =   3
            Left            =   780
            TabIndex        =   62
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "4"
            Height          =   195
            Index           =   4
            Left            =   780
            TabIndex        =   61
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "5"
            Height          =   255
            Index           =   5
            Left            =   780
            TabIndex        =   60
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   195
            Index           =   6
            Left            =   840
            TabIndex        =   59
            Top             =   1200
            Width           =   135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "7"
            Height          =   195
            Index           =   7
            Left            =   840
            TabIndex        =   58
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "8"
            Height          =   195
            Index           =   8
            Left            =   840
            TabIndex        =   57
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "9"
            Height          =   195
            Index           =   9
            Left            =   840
            TabIndex        =   56
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "10"
            Height          =   195
            Index           =   10
            Left            =   780
            TabIndex        =   55
            Top             =   240
            Width           =   195
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00000000&
            X1              =   420
            X2              =   420
            Y1              =   1560
            Y2              =   2760
         End
      End
      Begin VB.CommandButton BR_control 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bridge Control"
         Height          =   550
         Left            =   2355
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4936
         Width           =   860
      End
      Begin VB.CommandButton LCL_control 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Local Control"
         Height          =   550
         Left            =   1449
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4936
         Width           =   860
      End
      Begin VB.CommandButton ER_control 
         BackColor       =   &H00C0C0C0&
         Caption         =   "E.R Control"
         Height          =   550
         Left            =   543
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4936
         Width           =   860
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1800
         Top             =   540
      End
      Begin VB.CommandButton CLUTCHin 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clutch IN"
         Height          =   550
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5760
         Width           =   981
      End
      Begin VB.Frame Frame2 
         Caption         =   "      Clutch        "
         Height          =   907
         Left            =   5736
         TabIndex        =   46
         Top             =   178
         Width           =   1525
         Begin VB.Shape Clutch_in 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   360
            Shape           =   3  'Circle
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Clutch_out 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1020
            Shape           =   3  'Circle
            Top             =   360
            Width           =   195
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "In"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Out"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   47
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CommandButton EM_STOP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Emergency Stop"
         Height          =   550
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5040
         Width           =   981
      End
      Begin VB.CommandButton EM_RUN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Emergency Run"
         Height          =   550
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5040
         Width           =   981
      End
      Begin VB.Frame Frame1 
         Caption         =   "      Camshaft      "
         Height          =   907
         Left            =   3743
         TabIndex        =   41
         Top             =   178
         Width           =   1525
         Begin VB.Label Label6 
            Caption         =   "Astern"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   43
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Ahead"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   675
         End
         Begin VB.Shape Cam_astern 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1020
            Shape           =   3  'Circle
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Cam_ahead 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   360
            Shape           =   3  'Circle
            Top             =   360
            Width           =   195
         End
      End
      Begin VB.Frame Status_panel 
         Caption         =   "Status"
         Height          =   3642
         Left            =   242
         TabIndex        =   28
         Top             =   1189
         Width           =   3155
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bridge Control"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   11
            Left            =   2100
            TabIndex        =   40
            Top             =   2820
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Local Control"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   10
            Left            =   1200
            TabIndex        =   39
            Top             =   2820
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "E.R. Control"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   9
            Left            =   300
            TabIndex        =   38
            Top             =   2820
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Turning Gear Engaged"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   8
            Left            =   2100
            TabIndex        =   37
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Start Air Valve Blocked"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   7
            Left            =   1200
            TabIndex        =   36
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "M.E. Start Failure"
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
            Height          =   675
            Index           =   5
            Left            =   2100
            TabIndex        =   35
            Top             =   1260
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Wrong Way Alarm"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   4
            Left            =   1200
            TabIndex        =   34
            Top             =   1260
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Governor Limit Cancel"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   3
            Left            =   300
            TabIndex        =   33
            Top             =   1260
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Astern Valve Open"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   2
            Left            =   2100
            TabIndex        =   32
            Top             =   480
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Stop Valve Open"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   1
            Left            =   1200
            TabIndex        =   31
            Top             =   480
            Width           =   795
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ahead Valve Open"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   0
            Left            =   300
            TabIndex        =   30
            Top             =   480
            Width           =   795
         End
         Begin VB.Shape ME_status 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   11
            Left            =   2040
            Top             =   2760
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   10
            Left            =   1140
            Top             =   2760
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   9
            Left            =   240
            Top             =   2760
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   8
            Left            =   2040
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label ME_status_label 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Main Air Start Valve Open"
            ForeColor       =   &H00000000&
            Height          =   675
            Index           =   6
            Left            =   300
            TabIndex        =   29
            Top             =   2040
            Width           =   795
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   6
            Left            =   240
            Top             =   1980
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   7
            Left            =   1140
            Top             =   1980
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   5
            Left            =   2040
            Top             =   1200
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   4
            Left            =   1140
            Top             =   1200
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   3
            Left            =   240
            Top             =   1200
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   2
            Left            =   2040
            Top             =   420
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   1
            Left            =   1140
            Top             =   420
            Width           =   915
         End
         Begin VB.Shape ME_status 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   0
            Left            =   240
            Top             =   420
            Width           =   915
         End
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   6000
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Frame Pitch 
         Caption         =   "    Pitch   "
         Height          =   3043
         Left            =   5160
         TabIndex        =   9
         Top             =   1487
         Width           =   1102
         Begin VB.Shape Shape1 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   180
            Top             =   1380
            Width           =   495
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000000FF&
            X1              =   420
            X2              =   420
            Y1              =   1560
            Y2              =   2760
         End
         Begin VB.Label Label2 
            Caption         =   "8"
            Height          =   195
            Index           =   1
            Left            =   780
            TabIndex        =   27
            Top             =   2400
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "10"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   22
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "6"
            Height          =   195
            Index           =   2
            Left            =   780
            TabIndex        =   21
            Top             =   2160
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "4"
            Height          =   195
            Index           =   3
            Left            =   780
            TabIndex        =   20
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "2"
            Height          =   195
            Index           =   4
            Left            =   780
            TabIndex        =   19
            Top             =   1680
            Width           =   195
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0000FF00&
            X1              =   420
            X2              =   420
            Y1              =   360
            Y2              =   1560
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   21
            X1              =   180
            X2              =   660
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   20
            X1              =   180
            X2              =   660
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   19
            X1              =   180
            X2              =   660
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   18
            X1              =   180
            X2              =   660
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   17
            X1              =   180
            X2              =   660
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line2 
            Index           =   16
            X1              =   180
            X2              =   660
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   15
            X1              =   180
            X2              =   660
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   14
            X1              =   180
            X2              =   660
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   13
            X1              =   180
            X2              =   660
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   12
            X1              =   180
            X2              =   660
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   11
            X1              =   180
            X2              =   660
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   15
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "10"
            Height          =   195
            Index           =   10
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "2"
            Height          =   195
            Index           =   6
            Left            =   780
            TabIndex        =   13
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "8"
            Height          =   195
            Index           =   9
            Left            =   780
            TabIndex        =   12
            Top             =   480
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "4"
            Height          =   195
            Index           =   7
            Left            =   780
            TabIndex        =   11
            Top             =   960
            Width           =   195
         End
         Begin VB.Label Label2 
            Caption         =   "6"
            Height          =   195
            Index           =   8
            Left            =   780
            TabIndex        =   10
            Top             =   720
            Width           =   195
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   5700
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Frame Revs 
         Caption         =   "    Revs   "
         Height          =   3043
         Left            =   3900
         TabIndex        =   0
         Top             =   1487
         Width           =   1102
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            X1              =   420
            X2              =   420
            Y1              =   1560
            Y2              =   2760
         End
         Begin VB.Shape Shape2 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   180
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "10"
            Height          =   195
            Index           =   10
            Left            =   780
            TabIndex        =   26
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "8"
            Height          =   195
            Index           =   9
            Left            =   840
            TabIndex        =   25
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   195
            Index           =   8
            Left            =   840
            TabIndex        =   24
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "4"
            Height          =   195
            Index           =   7
            Left            =   840
            TabIndex        =   23
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "2"
            Height          =   195
            Index           =   6
            Left            =   840
            TabIndex        =   7
            Top             =   1200
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   780
            TabIndex        =   6
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "2"
            Height          =   195
            Index           =   4
            Left            =   780
            TabIndex        =   5
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "4"
            Height          =   195
            Index           =   3
            Left            =   780
            TabIndex        =   4
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "6"
            Height          =   255
            Index           =   2
            Left            =   780
            TabIndex        =   3
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   2
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "10"
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   1
            Top             =   2640
            Width           =   255
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   10
            X1              =   180
            X2              =   660
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   9
            X1              =   180
            X2              =   660
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   8
            X1              =   180
            X2              =   660
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   7
            X1              =   180
            X2              =   660
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   6
            X1              =   180
            X2              =   660
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   5
            X1              =   180
            X2              =   660
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   4
            X1              =   180
            X2              =   660
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   3
            X1              =   180
            X2              =   660
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   2
            X1              =   180
            X2              =   660
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   1
            X1              =   180
            X2              =   660
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   0
            X1              =   180
            X2              =   660
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            X1              =   420
            X2              =   420
            Y1              =   360
            Y2              =   1560
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Astern"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6762
         TabIndex        =   17
         Top             =   3508
         Width           =   675
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   6521
         X2              =   6762
         Y1              =   4222
         Y2              =   4222
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   6521
         X2              =   6762
         Y1              =   3033
         Y2              =   3033
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   6641
         X2              =   6641
         Y1              =   3033
         Y2              =   4222
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         X1              =   6521
         X2              =   6762
         Y1              =   1843
         Y2              =   1843
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         X1              =   6641
         X2              =   6641
         Y1              =   1843
         Y2              =   3033
      End
      Begin VB.Label Label4 
         Caption         =   "Ahead"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6762
         TabIndex        =   16
         Top             =   2319
         Width           =   675
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2100
      TabIndex        =   53
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "MEcontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private set_dummy_focus As Integer

Dim revs_mouse_down As Byte
Dim revs_lim_mouse_down As Byte
Dim pitch_mouse_down As Byte

Dim revs_top As Integer
Dim revs_bottom As Integer
Dim Rev_value As Single

Dim revs_lim_top As Integer
Dim revs_lim_bottom As Integer
Dim Rev_lim_value As Single

Dim pitch_top As Integer
Dim pitch_bottom As Integer
Dim Pitch_value As Single




Private Sub CLUTCHin_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
clutch_flag = True
Call setDI(406, clutch_flag)    ' CLUTCH_IN_SW
clutch_flag = False
Call setDI(398, clutch_flag)    ' CLUTCH_OUT_SW

End Sub

Private Sub Command1_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
clutch_flag = False
Call setDI(406, clutch_flag)    ' CLUTCH_IN_SW
clutch_flag = True
Call setDI(398, clutch_flag)    ' CLUTCH_OUT_SW


End Sub

Private Sub BR_control_Click()
Dim flag As Byte


set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
   
   
' 17/6/98
If Portugal_flag Then

 flag = True
 Call setDI(407, flag)    ' Bridge control = true
 flag = False
 Call setDI(414, flag)    '    LCL control = false
 flag = False
 Call setDI(393, flag)    '     ER control = false
 
Else
   ' VIGO/FERROL
   flag = False
   Call setDI(407, flag)    ' ER control = false
   flag = False
   Call setDI(393, flag)    ' LCL control = false
End If

End Sub



Private Sub Command3_Click()

' 4/12/98
sy011.Show

End Sub

Private Sub Command4_Click()

' 4/12/98
sy010.Show

End Sub

Private Sub EM_RUN_Click()
set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
Call setPCDI(394)     ' Emergency RUN

End Sub

Private Sub EM_STOP_Click()

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
Call setPCDI(392)     ' Emergency STOP

End Sub

Private Sub ER_control_Click()
Dim flag As Byte

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub


' 17/6/98
If Portugal_flag Then

 flag = True
 Call setDI(393, flag)    '     ER control = true
 flag = False
 Call setDI(414, flag)    '    LCL control = false
 flag = False
 Call setDI(407, flag)    ' Bridge control = false
 
Else
   ' VIGO/FERROL
 flag = True
 Call setDI(407, flag)    '     ER control = true
 flag = False
 Call setDI(393, flag)    '    LCL control = false
 
End If


End Sub

Private Sub form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idsyMECD)

'revs_top = 360 - Shape2.Height / 2
'revs_bottom = 2760 - Shape2.Height / 2
'pitch_top = 360 - Shape1.Height / 2
'pitch_bottom = 2760 - Shape1.Height / 2
'revs_lim_top = 360 - Shape3.Height / 2
'revs_lim_bottom = 2760 - Shape3.Height / 2

revs_top = Line1.Y1 - Shape2.Height / 2
revs_bottom = Line10.Y2 - Shape2.Height / 2
pitch_top = Line3.Y1 - Shape1.Height / 2
pitch_bottom = Line9.Y2 - Shape1.Height / 2
revs_lim_top = Line12.Y1 - Shape3.Height / 2
revs_lim_bottom = Line11.Y2 - Shape3.Height / 2


' 23/7/98
update_VBAI(1) = True
update_VBAI(2) = True
update_VBAI(3) = True

   
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idsyMECD)
    
End If
revs_top = Line1.Y1 - Shape2.Height / 2
revs_bottom = Line10.Y2 - Shape2.Height / 2
Shape2.Top = revs_top + (revs_bottom - revs_top) * (10 - Rev_value) / 10

pitch_top = Line3.Y1 - Shape1.Height / 2
pitch_bottom = Line9.Y2 - Shape1.Height / 2
Shape1.Top = pitch_top + (pitch_bottom - pitch_top) * (10 - Pitch_value) / 10

revs_lim_top = Line12.Y1 - Shape3.Height / 2
revs_lim_bottom = Line11.Y2 - Shape3.Height / 2
Shape3.Top = revs_lim_top + (revs_lim_bottom - revs_lim_top) * (10 - Rev_lim_value) / 10
HavenStretch1.Visible = True
End Sub



Private Sub Label1_Click(Index As Integer)

Shape2.Top = revs_top + (revs_bottom - revs_top) * (10 - Index) / 10
Rev_value = (revs_bottom - Shape2.Top) / (revs_bottom - revs_top) * 10#
Text1.Text = Rev_value

' 4/3/98  Adjust   AI(19) = MECD Throttle Pot
' 23/7/98
'       nCHANNEL = 19
     If Portugal_flag Then
       nCHANNEL = 29
     Else
       nCHANNEL = 19
     End If

       nVALUE = Rev_value * 25.5
       Call setAI(nCHANNEL, nVALUE)

End Sub

Private Sub Label2_Click(Index As Integer)

Shape1.Top = pitch_top + (pitch_bottom - pitch_top) * (10 - Index) / 10
Pitch_value = (pitch_bottom - Shape1.Top) / (pitch_bottom - pitch_top) * 10#
Text2.Text = Pitch_value

' 4/3/98  Adjust   AI(16) = MECD PITCH Pot
' 23/7/98
'       nCHANNEL = 16
     If Portugal_flag Then
       nCHANNEL = 25
     Else
       nCHANNEL = 16
     End If

       nVALUE = Pitch_value * 25.5
       Call setAI(nCHANNEL, nVALUE)


End Sub

Private Sub Label3_Click(Index As Integer)

Shape3.Top = revs_lim_top + (revs_lim_bottom - revs_lim_top) * (10 - Index) / 10
Rev_lim_value = (revs_lim_bottom - Shape3.Top) / (revs_lim_bottom - revs_lim_top) * 10#
Text3.Text = Rev_lim_value

' 4/3/98  Adjust   AI(18) = MECD REV Lim Pot
' 23/7/98
'      nCHANNEL = 18
     If Portugal_flag Then
       nCHANNEL = 27
     Else

       nCHANNEL = 18
     End If
 
       nVALUE = Rev_lim_value * 25.5
       Call setAI(nCHANNEL, nVALUE)

End Sub

Private Sub LCL_control_Click()
Dim flag As Byte

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   


' 17/6/98
If Portugal_flag Then

 flag = True
 Call setDI(414, flag)    '    LCL control = true
 flag = False
 Call setDI(393, flag)    '     ER control = false
 flag = False
 Call setDI(407, flag)    ' Bridge control = false
 
Else

   ' VIGO/FERROL
  flag = False
  Call setDI(407, flag)    ' ER control = false
  flag = True
  Call setDI(393, flag)    ' LCL control = true

End If

End Sub

Private Sub Pitch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

pitch_mouse_down = True

End Sub

Private Sub Pitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 10/12/98
If Not HARDWARE_flag Then

    If pitch_mouse_down Then

        If Y < pitch_top Then
            Shape1.Top = pitch_top
        ElseIf Y > pitch_bottom Then
            Shape1.Top = pitch_bottom
        Else
            Shape1.Top = Y
        End If
  
        'End If

         Pitch_value = (pitch_bottom - Shape1.Top) / (pitch_bottom - pitch_top) * 10#
        Text2.Text = Pitch_value

         ' 4/3/98  Adjust   AI(  ) = MECD PITCH Pot
            ' 17/6/98
        If Portugal_flag Then
            nCHANNEL = 25
        Else
            nCHANNEL = 16
        End If
     
       nVALUE = Pitch_value * 25.5
       Call setAI(nCHANNEL, nVALUE)
    End If

End If

End Sub

Private Sub Pitch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

pitch_mouse_down = False

End Sub

Private Sub REV_limiter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

revs_lim_mouse_down = True

End Sub

Private Sub REV_limiter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 10/12/98
If Not HARDWARE_flag Then

    If revs_lim_mouse_down Then

        If Y < revs_lim_top Then
            Shape3.Top = revs_lim_top
        ElseIf Y > revs_lim_bottom Then
            Shape3.Top = revs_lim_bottom
        Else
            Shape3.Top = Y
        End If
  
        'End If

        Rev_lim_value = (revs_lim_bottom - Shape3.Top) / (revs_lim_bottom - revs_lim_top) * 10#
        Text3.Text = Rev_lim_value


         ' 4/3/98  Adjust   AI(  ) = MECD REV LIM Pot
         ' 17/6/98
        If Portugal_flag Then
            nCHANNEL = 27
        Else
            nCHANNEL = 18
        End If
     
       nVALUE = Rev_lim_value * 25.5
       Call setAI(nCHANNEL, nVALUE)

    End If

End If

End Sub

Private Sub REV_limiter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

revs_lim_mouse_down = False

End Sub

Private Sub Revs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

revs_mouse_down = True

End Sub

Private Sub Revs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' 10/12/98
If Not HARDWARE_flag Then

    If revs_mouse_down Then

        If Y < revs_top Then
            Shape2.Top = revs_top
        ElseIf Y > revs_bottom Then
            Shape2.Top = revs_bottom
        Else
            Shape2.Top = Y
        End If
  
     'End If

        Rev_value = (revs_bottom - Shape2.Top) / (revs_bottom - revs_top) * 10#
        Text1.Text = Rev_value


        ' 4/3/98  Adjust   AI(  ) = MECD Throttle Pot
        ' 17/6/98
         If Portugal_flag Then
            nCHANNEL = 29
        Else
            nCHANNEL = 19
        End If
     
       nVALUE = Rev_value * 25.5
       Call setAI(nCHANNEL, nVALUE)

    End If

End If

End Sub

Private Sub Revs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

revs_mouse_down = False

End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim pot_channel As Integer

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

Text4.Text = "REVS    vb_pots(29) = " & VB_pots(29)
Text5.Text = "PITCH   vb_pots(25) = " & VB_pots(25)
Text6.Text = "REV LIM vb_pots(27) = " & VB_pots(27)

' 23/7/98 ( see TakeSnapForm and HOST.frm )
' REVS
If update_VBAI(1) Then
   update_VBAI(1) = False
   
'   Rev_value = (revs_bottom - Shape2.Top) / (revs_bottom - revs_top) * 10#
     If Portugal_flag Then
       pot_channel = 29
     Else
       pot_channel = 19
     End If
 
    Rev_value = VB_pots(pot_channel) / 25.5
    Shape2.Top = revs_bottom - Rev_value * (revs_bottom - revs_top) / 10#
    
End If

' REV Limiter
If update_VBAI(2) Then
   update_VBAI(2) = False
   
'   Rev_lim_value = (revs_lim_bottom - Shape3.Top) / (revs_lim_bottom - revs_lim_top) * 10#
      ' 17/6/98
     If Portugal_flag Then
       pot_channel = 27
     Else

       pot_channel = 18
     End If

    Rev_lim_value = VB_pots(pot_channel) / 25.5
    Shape3.Top = revs_lim_bottom - Rev_lim_value * (revs_lim_bottom - revs_lim_top) / 10#
    
End If


' PITCH
If update_VBAI(3) Then
   update_VBAI(3) = False
   
'   Pitch_value = (pitch_bottom - Shape1.Top) / (pitch_bottom - pitch_top) * 10#
     If Portugal_flag Then
       pot_channel = 25
     Else
       pot_channel = 16
     End If

    Pitch_value = VB_pots(pot_channel) / 25.5
    Shape1.Top = pitch_bottom - Pitch_value * (pitch_bottom - pitch_top) / 10#
    
End If



' Status Indicators
For i = 0 To 11
   If VB_lamps(36 + i) Then
        ME_status_label(i).BackColor = &HFFFFFF      ' WHITE
              ME_status(i).FillColor = &HFFFFFF     ' WHITE
   Else
        ME_status_label(i).BackColor = &HC0C0C0      ' GREY
              ME_status(i).FillColor = &HC0C0C0      ' GREY
   End If
Next i

' 3/4/98
' Emergency RUN
If VB_lamps(48) Then
    EM_RUN.BackColor = &HFF&        ' RED
Else
    EM_RUN.BackColor = &HC0C0C0     ' GREY
End If
' Emergency STOP
If VB_lamps(49) Then
    EM_STOP.BackColor = &HFF&        ' RED
Else
    EM_STOP.BackColor = &HC0C0C0     ' GREY
End If

' Clutch
If VB_lamps(52) Then
    Clutch_in.FillColor = &HFF00&       ' GREEN
   Clutch_out.FillColor = &HC0C0C0      ' GREY
Else
    Clutch_in.FillColor = &HC0C0C0      ' GREY
   Clutch_out.FillColor = &HFF&         ' RED
End If

' Camshaft AHEAD
If VB_lamps(50) Then
    Cam_ahead.FillColor = &HFF00&       ' GREEN
Else
    Cam_ahead.FillColor = &HC0C0C0      ' GREY
End If
' Camshaft ASTERN
If VB_lamps(51) Then
    Cam_astern.FillColor = &HFF&         ' RED
Else
    Cam_astern.FillColor = &HC0C0C0      ' GREY
End If

End Sub
