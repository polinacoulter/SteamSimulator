VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HAVENSTRETCH.OCX"
Begin VB.Form frm_faults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fault Entry"
   ClientHeight    =   13710
   ClientLeft      =   1920
   ClientTop       =   2250
   ClientWidth     =   10800
   FillStyle       =   0  'Solid
   Icon            =   "frm_faults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   13710
   ScaleWidth      =   10800
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   13575
      Left            =   60
      Top             =   60
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   23945
      Begin VB.Frame FaultEntry 
         Height          =   13395
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10095
         Begin VB.CommandButton FaultClearCommand 
            Caption         =   "Clear"
            Height          =   375
            Left            =   9000
            TabIndex        =   81
            Top             =   2940
            Width           =   975
         End
         Begin VB.ListBox FaultTypes 
            BackColor       =   &H00C0C0C0&
            Columns         =   1
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   4620
            Left            =   120
            TabIndex        =   73
            Top             =   4620
            Width           =   3195
         End
         Begin VB.ListBox Description 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4620
            Left            =   3420
            TabIndex        =   72
            Top             =   4620
            Width           =   6615
         End
         Begin VB.TextBox SelectedFault 
            Height          =   375
            Left            =   180
            TabIndex        =   71
            Top             =   540
            Width           =   6915
         End
         Begin VB.CommandButton FaultEnterCommand 
            Caption         =   "Enter"
            Height          =   375
            Left            =   7260
            TabIndex        =   70
            Top             =   540
            Width           =   975
         End
         Begin VB.ListBox ActiveFaultList 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1320
            Left            =   180
            TabIndex        =   69
            Top             =   2880
            Width           =   8655
         End
         Begin VB.CommandButton ClearAll 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   9000
            TabIndex        =   68
            Top             =   3480
            Width           =   975
         End
         Begin VB.CommandButton Cancel 
            Caption         =   "OK"
            Height          =   375
            Left            =   9000
            TabIndex        =   67
            Top             =   4200
            Width           =   975
         End
         Begin VB.Timer Timer1 
            Interval        =   300
            Left            =   8100
            Top             =   4260
         End
         Begin VB.TextBox FaultScheduleTime 
            Height          =   375
            Left            =   9000
            TabIndex        =   66
            Text            =   "Text1"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ListBox ScheduledFaultList 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1320
            Left            =   180
            TabIndex        =   65
            Top             =   1200
            Width           =   8655
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Clear"
            Height          =   375
            Left            =   9000
            TabIndex        =   64
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Clear All"
            Height          =   375
            Left            =   9000
            TabIndex        =   63
            Top             =   2160
            Width           =   975
         End
         Begin VB.Frame Severity_frame 
            Caption         =   "Fault Severity "
            Height          =   1335
            Left            =   120
            TabIndex        =   54
            Top             =   9240
            Width           =   9915
            Begin VB.HScrollBar Severity_scroll 
               Height          =   315
               Left            =   1740
               Max             =   100
               TabIndex        =   55
               Top             =   540
               Value           =   25
               Width           =   6615
            End
            Begin VB.Label Severity_label 
               Alignment       =   2  'Center
               Caption         =   "Severity_label"
               Height          =   255
               Left            =   1980
               TabIndex        =   61
               Top             =   1020
               Width           =   6135
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   1980
               X2              =   8100
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Line Line2 
               BorderWidth     =   2
               X1              =   8100
               X2              =   8100
               Y1              =   420
               Y2              =   240
            End
            Begin VB.Label Label6 
               Caption         =   "100%"
               Height          =   255
               Left            =   8160
               TabIndex        =   60
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label10 
               Caption         =   "0%"
               Height          =   195
               Left            =   2040
               TabIndex        =   59
               Top             =   180
               Width           =   615
            End
            Begin VB.Line Line3 
               BorderWidth     =   2
               X1              =   1980
               X2              =   1980
               Y1              =   420
               Y2              =   240
            End
            Begin VB.Line Line4 
               BorderWidth     =   2
               X1              =   5040
               X2              =   5040
               Y1              =   420
               Y2              =   240
            End
            Begin VB.Label Label11 
               Caption         =   "50%"
               Height          =   195
               Left            =   5100
               TabIndex        =   58
               Top             =   180
               Width           =   615
            End
            Begin VB.Line Line5 
               BorderWidth     =   2
               X1              =   6540
               X2              =   6540
               Y1              =   420
               Y2              =   240
            End
            Begin VB.Label Label12 
               Caption         =   "75%"
               Height          =   195
               Left            =   6600
               TabIndex        =   57
               Top             =   180
               Width           =   615
            End
            Begin VB.Line Line6 
               BorderWidth     =   2
               X1              =   3540
               X2              =   3540
               Y1              =   420
               Y2              =   240
            End
            Begin VB.Label Label13 
               Caption         =   "25%"
               Height          =   195
               Left            =   3600
               TabIndex        =   56
               Top             =   180
               Width           =   615
            End
         End
         Begin VB.Frame Bearing_frame 
            Caption         =   "Bearing Fault Severity"
            Height          =   2715
            Left            =   60
            TabIndex        =   2
            Top             =   10620
            Width           =   9855
            Begin VB.Label List_lbl 
               Caption         =   "25. Stern Tube Brg"
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
               Index           =   25
               Left            =   6840
               TabIndex        =   37
               Top             =   2160
               Width           =   1935
            End
            Begin VB.Label List_lbl 
               Caption         =   "24. Line Shaft"
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
               Index           =   24
               Left            =   6840
               TabIndex        =   36
               Top             =   1860
               Width           =   1755
            End
            Begin VB.Label List_lbl 
               Caption         =   "23. LP First Redn Pinion Aft"
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
               Index           =   23
               Left            =   6840
               TabIndex        =   35
               Top             =   1620
               Width           =   2175
            End
            Begin VB.Label List_lbl 
               Caption         =   "21. LP First Redn Aft"
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
               Index           =   21
               Left            =   6840
               TabIndex        =   33
               Top             =   1140
               Width           =   2055
            End
            Begin VB.Label List_lbl 
               Caption         =   "20. LP First Redn Fwd"
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
               Index           =   20
               Left            =   6840
               TabIndex        =   32
               Top             =   900
               Width           =   1815
            End
            Begin VB.Label List_lbl 
               Caption         =   "19. LP Pinion Aft"
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
               Index           =   19
               Left            =   6840
               TabIndex        =   31
               Top             =   660
               Width           =   2115
            End
            Begin VB.Label Bearing_severity_lbl 
               Caption         =   "100 %"
               Height          =   195
               Index           =   0
               Left            =   1200
               TabIndex        =   28
               Top             =   180
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "100 %"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   2520
               TabIndex        =   27
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   2520
               TabIndex        =   26
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   3
               Left            =   2520
               TabIndex        =   25
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   4
               Left            =   2520
               TabIndex        =   24
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   2520
               TabIndex        =   23
               Top             =   1320
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   6
               Left            =   2520
               TabIndex        =   22
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   7
               Left            =   2520
               TabIndex        =   21
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   8
               Left            =   2520
               TabIndex        =   20
               Top             =   2040
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   9
               Left            =   5760
               TabIndex        =   19
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   10
               Left            =   5760
               TabIndex        =   18
               Top             =   540
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   11
               Left            =   5760
               TabIndex        =   17
               Top             =   780
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   12
               Left            =   5760
               TabIndex        =   16
               Top             =   1020
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   13
               Left            =   5760
               TabIndex        =   15
               Top             =   1260
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   14
               Left            =   5760
               TabIndex        =   14
               Top             =   1500
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   15
               Left            =   5760
               TabIndex        =   13
               Top             =   1740
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   16
               Left            =   5760
               TabIndex        =   12
               Top             =   1980
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   17
               Left            =   9000
               TabIndex        =   11
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   18
               Left            =   9000
               TabIndex        =   10
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   19
               Left            =   9000
               TabIndex        =   9
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   20
               Left            =   9000
               TabIndex        =   8
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   21
               Left            =   9000
               TabIndex        =   7
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   22
               Left            =   9000
               TabIndex        =   6
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   23
               Left            =   9000
               TabIndex        =   5
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   24
               Left            =   9000
               TabIndex        =   4
               Top             =   1920
               Width           =   615
            End
            Begin VB.Label Bearing_severity_lbl 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "100 %"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   25
               Left            =   9000
               TabIndex        =   3
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label List_lbl 
               Caption         =   "9. HP First Redn Pinion Aft"
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
               Index           =   9
               Left            =   3780
               TabIndex        =   51
               Top             =   300
               Width           =   2475
            End
            Begin VB.Label List_lbl 
               Caption         =   "11. Bull Wheel Aft"
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
               Index           =   11
               Left            =   3720
               TabIndex        =   53
               Top             =   780
               Width           =   2595
            End
            Begin VB.Label List_lbl 
               Caption         =   "10. Bull Wheel Fwd"
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
               Index           =   10
               Left            =   3720
               TabIndex        =   52
               Top             =   540
               Width           =   2595
            End
            Begin VB.Label List_lbl 
               Caption         =   "16. LP Thrust Astern"
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
               Index           =   16
               Left            =   3660
               TabIndex        =   42
               Top             =   2040
               Width           =   2655
            End
            Begin VB.Label List_lbl 
               Caption         =   "15. LP Turbine Aft"
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
               Index           =   15
               Left            =   3660
               TabIndex        =   41
               Top             =   1740
               Width           =   2595
            End
            Begin VB.Label List_lbl 
               Caption         =   "14. LP Turbine Fwd"
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
               Index           =   14
               Left            =   3660
               TabIndex        =   40
               Top             =   1500
               Width           =   2595
            End
            Begin VB.Label List_lbl 
               Caption         =   "13. Main Thrust Brg Astern"
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
               Index           =   13
               Left            =   3660
               TabIndex        =   39
               Top             =   1260
               Width           =   2715
            End
            Begin VB.Label List_lbl 
               Caption         =   "12. Main Thrust Brg Ahead"
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
               Index           =   12
               Left            =   3660
               TabIndex        =   38
               Top             =   1020
               Width           =   2655
            End
            Begin VB.Label List_lbl 
               Caption         =   "22. LP First Redn Pinion Fwd"
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
               Index           =   22
               Left            =   6840
               TabIndex        =   34
               Top             =   1380
               Width           =   2295
            End
            Begin VB.Label List_lbl 
               Caption         =   "18. LP Pinion Fwd"
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
               Index           =   18
               Left            =   6840
               TabIndex        =   30
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label List_lbl 
               Caption         =   "17. LP Thrust Ahead"
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
               Index           =   17
               Left            =   6840
               TabIndex        =   29
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label List_lbl 
               Caption         =   "8. HP First Redn Pinion Fwd"
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
               Index           =   8
               Left            =   420
               TabIndex        =   50
               Top             =   2040
               Width           =   2415
            End
            Begin VB.Label List_lbl 
               Caption         =   "7. HP First Redn Aft"
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
               Index           =   7
               Left            =   420
               TabIndex        =   49
               Top             =   1800
               Width           =   2355
            End
            Begin VB.Label List_lbl 
               Caption         =   "6. HP First Redn Fwd."
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
               Index           =   6
               Left            =   420
               TabIndex        =   48
               Top             =   1560
               Width           =   2295
            End
            Begin VB.Label List_lbl 
               Caption         =   "5. HP Pinion Aft"
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
               Index           =   5
               Left            =   420
               TabIndex        =   47
               Top             =   1320
               Width           =   2355
            End
            Begin VB.Label List_lbl 
               Caption         =   "4. HP Pinion Fwd"
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
               Index           =   4
               Left            =   420
               TabIndex        =   46
               Top             =   1080
               Width           =   2355
            End
            Begin VB.Label List_lbl 
               Caption         =   "3. HP Turbine Aft"
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
               Index           =   3
               Left            =   420
               TabIndex        =   45
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label List_lbl 
               Caption         =   "2. HP Thrust Brg"
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
               Left            =   420
               TabIndex        =   44
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label List_lbl 
               Caption         =   "1. HP Turbine Fwd  "
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
               Left            =   420
               TabIndex        =   43
               Top             =   360
               Width           =   2295
            End
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   375
            Left            =   9600
            TabIndex        =   62
            Top             =   1200
            Width           =   195
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Fault Types"
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   4320
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Description"
            Height          =   255
            Left            =   3420
            TabIndex        =   79
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Scheduled Faults"
            Height          =   255
            Left            =   180
            TabIndex        =   78
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Selected Fault"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Entered"
            Height          =   255
            Left            =   7620
            TabIndex        =   76
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Minutes"
            Height          =   435
            Left            =   8700
            TabIndex        =   75
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Active Faults"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   2640
            Width           =   1815
         End
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Faults Entered"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frm_faults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub Fill_Active_Fault_List(form_name As Form)
Dim i As Integer
Dim j As Integer
Dim FaultInListOK As Byte
Dim ItemToRemove As Integer
Dim DummyText As String
Dim PositionQ As Integer
Dim String1 As String
Dim String2 As String
Dim TimeStr As String
Dim SingleScan As Integer
Dim Offset As Integer
Dim SSFreqAft11, SSFreqFord14, SSVoltsAft71, SSVoltsAft74 As Single
Dim FaultOK As Integer
Dim FaultAlreadyEntered As Integer
Dim LenText As Integer
Dim FString As String
Dim TText As String
Dim Padding As Integer
Dim i1 As Integer


' ===25/3/98===
   For i = 1 To BiggestFaultNumber
           Fault(i) = False
   Next i
    '11/10/99 calc no of faults here
    
    NumberOfFaults = 0
    For i = 1 To 32
       If active_faults(i) > 0 Then
          Fault(active_faults(i)) = True
          NumberOfFaults = NumberOfFaults + 1
       End If
   Next i
   'Debug.Print "active_faults(1:3) =  " & active_faults(1) & "  " & active_faults(2) & "  " & active_faults(3)
'==============

   '11/10/99 only scan if no of faults has changed to save time
   If NumberOfFaults <> OldNumberOffaults Or FaultUpdateRequired Then
   
    FaultUpdateRequired = False
    '11/10/99 NumberOfFaults = 0
    For i = 1 To BiggestFaultNumber
    
   '11/10/99 reduce scan to real faults only
        If FaultXRef(i) > 0 Then
   
    
    
    
    
        If Fault(i) Then                                                                                     'if a fault is really is active......
                '11/10/99 NumberOfFaults = NumberOfFaults + 1
                'TimeStr = Time
                TimeStr = Format(Time, "h:m:s")
                FaultInListOK = False
                For j = 0 To form_name.ActiveFaultList.ListCount - 1              'check if it is in the Active Fault list
                        If Val(form_name.ActiveFaultList.List(j)) = i Then
                            FaultInListOK = True
                        End If
                 Next
                If FaultInListOK = False Then                                                        'if not the add it to the list using the Xref array
                            
                            FString = Faultext(FaultXRef(i))
                            LenText = Len(FString)
                            Padding = 70 - LenText
                            If Padding > 0 Then
                                FString = FString + Space(Padding) & fault_time(i)
                            Else
                                FString = FString & fault_time(i)
                            End If
                            
                            form_name.ActiveFaultList.AddItem FString

                           
                            Call Logging("INSTUCTOR FAULT ENTERED: " & Faultext(FaultXRef(i)), LogFaults)
              
                 End If
                        Else 'take it out if fault has been cleared
                         FaultInListOK = True
                         For j = 0 To form_name.ActiveFaultList.ListCount - 1
                                If Val(form_name.ActiveFaultList.List(j)) = i Then              'then there is  fault in the list which should not be there
                                FaultInListOK = False
                        End If
                 Next
                 If FaultInListOK = False Then  'clear list  so it can be updated
                        ActiveFaultList.clear
                        FaultUpdateRequired = True
                End If
         End If
         End If
         
   Next i
   OldNumberOffaults = NumberOfFaults '11/10/99
 End If
 
  
   
 
End Sub
Sub FillDescription(FirstFault As Integer, LastFault As Integer)
Dim i As Integer
Dim FaultOK As Integer
Dim ThisFaultnumber As Integer
Dim FaultAlreadyEntered As Integer



For i = FirstFault To LastFault                                     'check for range of fault numbers defined by group

     ThisFaultnumber = Val(Faultext(i))              'get fault number
     If ThisFaultnumber > 0 Then
                Description.AddItem " " & Faultext(i)
     End If
 
   
Next i

End Sub

Sub FillFaultType()

'NOT USED ANY MORE - Description included with FaultDescription.txt

'Fault_Types(1) = "Main Engine  "
'Fault_Types(2) = "Fuel Oil (FO)              "
'Fault_Types(3) = "High Temperature Cooling Water (HTCW) "
'Fault_Types(4) = "Low Temperature Cooling Water (LTCW) "
'Fault_Types(5) = "Sea Water Cooling (SW)"
'Fault_Types(6) = "Main Engine Lubricating Oil (LO) "
'Fault_Types(7) = "Gearbox Lubricating Oil "
'Fault_Types(8) = "Controllable Pitch Propeller (CPP)"
'Fault_Types(9) = "Compressed Air"
'Fault_Types(10) = " Miscellaneous                           "
'Fault_Types(11) = "                                                   "
'Fault_Types(12) = "                                                   "
'Fault_Types(13) = "                                                   "
'Fault_Types(14) = "                                                   "
'Fault_Types(15) = "                                                   "
'Fault_Types(16) = "                                                   "

End Sub
Sub fill_faultext()

Dim i As Integer





' 8/9/97  Copied from the original
'' ENGINE....
Faultext(1) = " 1     CYL #1 - WORN/BROKEN COMPRESSION RING"
Faultext(2) = " 2     CYL #2 - WORN/BROKEN COMPRESSION RING"
Faultext(3) = " 3     CYL #3 - WORN/BROKEN COMPRESSION RING"
Faultext(4) = " 4     CYL #4 - WORN/BROKEN COMPRESSION RING"
Faultext(5) = " 5     CYL #1 - BURNED PISTON CROWN"
Faultext(6) = " 6     CYL #2 - BURNED PISTON CROWN"
Faultext(7) = " 7     CYL #3 - BURNED PISTON CROWN"
Faultext(8) = " 8     CYL #4 - BURNED PISTON CROWN"
Faultext(9) = " 9     ME AUTO AIR START VALVE SEIZED SHUT"
Faultext(10) = "10   CYL #1 - BURNED EXHAUST VALVE SEAT"
Faultext(11) = "11   CYL #2 - BURNED EXHAUST VALVE SEAT"
Faultext(12) = "12   CYL #3 - BURNED EXHAUST VALVE SEAT"
Faultext(13) = "13   CYL #4 - BURNED EXHAUST VALVE SEAT"
Faultext(14) = "14   TURBOCHARGER AIR FILTER DIRTY"
Faultext(15) = "15   TURBOCHARGER TURBINE BLADES FOULED"
Faultext(16) = "16   MAIN ENGINE OVERLOAD"
       
' 2 FUEL OIL
Faultext(21) = "51    HP FUEL PUMP #1 - LOW FUEL DISCHARGE"
Faultext(22) = "52    HP FUEL PUMP #2 - LOW FUEL DISCHARGE"
Faultext(23) = "53    HP FUEL PUMP #3 - LOW FUEL DISCHARGE"
Faultext(24) = "54    HP FUEL PUMP #4 - LOW FUEL DISCHARGE"
Faultext(25) = "55    HP FUEL PUMP #1 - HIGH FUEL DISCHARGE"
Faultext(26) = "56    HP FUEL PUMP #2 - HIGH FUEL DISCHARGE"
Faultext(27) = "57    HP FUEL PUMP #3 - HIGH FUEL DISCHARGE"
Faultext(28) = "58    HP FUEL PUMP #4 - HIGH FUEL DISCHARGE"
Faultext(29) = "59    FUEL INJECTOR #1 - CHOKED NOZZLE"
Faultext(30) = "60    FUEL INJECTOR #2 - CHOKED NOZZLE"
Faultext(31) = "61    FUEL INJECTOR #3 - CHOKED NOZZLE"
Faultext(32) = "62    FUEL INJECTOR #4 - CHOKED NOZZLE"
Faultext(33) = "63    FUEL INJECTOR #1 - LEAKING NOZZLE"
Faultext(34) = "64    FUEL INJECTOR #2 - LEAKING NOZZLE"
Faultext(35) = "65    FUEL INJECTOR #3 - LEAKING NOZZLE"
Faultext(36) = "66    FUEL INJECTOR #4 - LEAKING NOZZLE"
Faultext(37) = "67    FO BOOSTER PUMP #1 - LOW DISCHARGE PRESSURE"
Faultext(38) = "68    FO BOOSTER PUMP #2 - LOW DISCHARGE PRESSURE"
Faultext(39) = "69    FO BOOSTER PUMP #1 - MOTOR FAILURE"
Faultext(41) = "71    FO BOOSTER PUMP #2 - MOTOR FAILURE"
Faultext(42) = ""
Faultext(43) = ""
Faultext(44) = "74    EXCESS WATER IN FUEL OIL SERVICE TANK"
Faultext(45) = "75    FO SEPARATOR SEAL FAILURE"
Faultext(46) = ""
Faultext(47) = "77    FO DISCHARGE FILTER #1 - DIFF PRESSURE HIGH"
Faultext(48) = "78    FO DISCHARGE FILTER #2 - DIFF PRESSURE HIGH"

' 3 HTCW
'Faultext(51) = "101   CYL #1 - JACKET LINER SEAL LEAKING"
'Faultext(52) = "102   CYL #2 - JACKET LINER SEAL LEAKING"
'Faultext(53) = "103   CYL #3 - JACKET LINER SEAL LEAKING"
'Faultext(54) = "104   CYL #4 - JACKET LINER SEAL LEAKING"
Faultext(55) = "105   HTCW PUMP # 1 - MOTOR FAILURE"
Faultext(56) = "106   HTCW PUMP # 2 - MOTOR FAILURE"
Faultext(57) = "107   HTCW PUMP # 1 - WORN IMPELLER"
Faultext(58) = "108   HTCW PUMP # 2 - WORN IMPELLER"
Faultext(59) = "109   HTCW PUMP # 1 - GLAND LEAKING"
Faultext(60) = "110   HTCW PUMP # 2 - GLAND LEAKING"
Faultext(61) = "111   HTCW COOLER THERMOSTAT FAULTY - FULL BYPASS"
Faultext(62) = "112   HTCW COOLER THERMOSTAT FAULTY - NO BYPASS"
Faultext(63) = "113   HTCW COOLER #1 - FOULED ON HTCW SIDE"
Faultext(64) = "114   HTCW COOLER #2 - FOULED ON HTCW SIDE"
Faultext(65) = "115   HTCW HEATER FAILURE"
'
''  4 LTCW
Faultext(71) = "201   LTCW PUMP #1 - MOTOR FAILURE"
Faultext(72) = "202   LTCW PUMP #2 - MOTOR FAILURE"
Faultext(73) = "203   LTCW PUMP #1 - WORN IMPELLER"
Faultext(74) = "204   LTCW PUMP #2 - WORN IMPELLER"
Faultext(75) = ""
Faultext(76) = ""
Faultext(77) = "207   LTCW COOLER #1 - FOULING LTCW SIDE"
Faultext(78) = "208   LTCW COOLER #2 - FOULING LTCW SIDE"
'Faultext(79) = "209   GEARBOX LO COOLER - FOULING LTCW SIDE"
Faultext(80) = "210   LO COOLER #1 - FOULING LTCW SIDE"
Faultext(81) = "211   LO COOLER #2 - FOULING LTCW SIDE"
Faultext(82) = "212   HTCW COOLER #1 - FOULING LTCW SIDE"
Faultext(83) = "213   HTCW COOLER #2 - FOULING LTCW SIDE"
Faultext(84) = "214   CHARGE AIR COOLER - FOULING LTCW SIDE"
Faultext(85) = "215   LTCW COOLER THERMOSTAT FAULTY - FULL BYPASS"
Faultext(86) = "216   LTCW COOLER THERMOSTAT FAULTY - NO BYPASS"
Faultext(87) = "217   FAULTY CHARGE AIR THERMOSTAT - FULL BYPASS"
Faultext(88) = "218   FAULTY CHARGE AIR THERMOSTAT - NO BYPASS"
'
''  5 SW
Faultext(91) = "251   SW PUMP #1 - MOTOR FAILURE"
Faultext(92) = "252   SW PUMP #2 - MOTOR FAILURE"
Faultext(93) = "253   SW PUMP #1 - WORN IMPELLER"
Faultext(94) = "254   SW PUMP #2 - WORN IMPELLER"
Faultext(95) = "255   SW SUCTION STRAINER #1 - DP HIGH"
Faultext(96) = "256   SW SUCTION STRAINER #2 - DP HIGH"
Faultext(97) = "257   CENTRAL COOLER #1 FOULED ON SW SIDE"
Faultext(98) = "258   CENTRAL COOLER #2 FOULED ON SW SIDE"
'
' 6  MELO
Faultext(101) = "151   LO PUMP #1 - MOTOR FAILURE"
Faultext(102) = "152   LO PUMP #2 - MOTOR FAILURE"
Faultext(103) = "153   LO PUMP #1 - LOW DISCHARGE PRESSURE"
Faultext(104) = "154   LO PUMP #2 - LOW DISCHARGE PRESSURE"
Faultext(105) = ""
Faultext(106) = ""
Faultext(107) = ""
Faultext(108) = ""
Faultext(109) = "159   LO FILTER #1 - DIFF PRESSURE HIGH"
Faultext(110) = "160   LO FILTER #2 - DIFF PRESSURE HIGH"
Faultext(111) = "161   LO COOLER THERMOSTAT FAULTY - NO BYPASS"
Faultext(112) = "162   LO COOLER THERMOSTAT FAULTY - FULL BYPASS"
'Faultext(113) = "163   LO SEPARATOR SEAL FAILURE"
Faultext(114) = "164   LO COOLER #1 FOULED - LO SIDE"
Faultext(115) = "165   LO COOLER #2 FOULED - LO SIDE"
Faultext(116) = "166   MAIN ENGINE SUMP LEVEL LOW"

'7 GLO
Faultext(121) = "301   GEARBOX LO PUMP #1 - MOTOR FAILURE   "
Faultext(122) = "302   GEARBOX LO PUMP #2 - MOTOR FAILURE   "
Faultext(123) = ""
Faultext(124) = ""
Faultext(125) = ""
Faultext(126) = ""
Faultext(127) = ""

'8 CPP
Faultext(131) = "351   CPP PUMP #1 - MOTOR FAILURE"
Faultext(132) = "352   CPP PUMP #2 - MOTOR FAILURE"
Faultext(133) = ""
Faultext(134) = ""
Faultext(135) = ""

'9 AIR
Faultext(141) = ""
Faultext(142) = "402   AIR COMPRESSOR #1 - MOTOR FAILURE"
Faultext(143) = "403   AIR COMPRESSOR #2 - MOTOR FAILURE"
Faultext(144) = ""
Faultext(145) = "405   MAIN AIR LEAKAGE (BOTTLE) No. 1  "
Faultext(146) = "406   MAIN AIR LEAKAGE (BOTTLE) No. 2"
Faultext(147) = ""
Faultext(148) = ""

'MISCELLONOUS
Faultext(151) = ""
Faultext(152) = ""
Faultext(153) = ""
Faultext(154) = "454   TURNING GEAR INTERLOCK SEIZED 'ON'"
Faultext(155) = ""
Faultext(156) = "456   BILGE SUCTION STRUM BOX BLOCKED"
Faultext(157) = "457   BILGE LEVEL HIGH"
Faultext(158) = ""
Faultext(159) = "459   ME CLUTCH FRICTION PLATES WORN"


    For i = 1 To 160
        ThisFaultnumber = Val(Faultext(i))
        If ThisFaultnumber > 0 Then
            FaultXRef(ThisFaultnumber) = i
        Else
            FaultXRef(ThisFaultnumber) = -1
        End If
    Next

End Sub

Sub FillXref()
Dim i As Integer


'Makes a fault cross referenmce array.
'       Element number = actual fault number (1-1000)
'       Value = Text array reference(1-260)


For i = 1 To 260
ThisFaultnumber = Val(Faultext(i))
    If ThisFaultnumber > 0 Then
        FaultXRef(ThisFaultnumber) = i
    Else
        FaultXRef(ThisFaultnumber) = -1
    End If
Next

End Sub


Private Sub ActiveFaultList_Click()
Dim AllText As String
Dim MyPos As Long

'strip off time data from string and put in Selected Fault box
'AllText = ActiveFaultList.List(ActiveFaultList.ListIndex)
'MyPos = InStr(4, AllText, "Entered", 1)
'SelectedFault.Text = Mid(ActiveFaultList.List(ActiveFaultList.ListIndex), 1, MyPos - 1)
SelectedFault.text = Mid(ActiveFaultList.List(ActiveFaultList.ListIndex), 1, 50)
SelectedFault.Enabled = True

'enable the clear command because the fault can only be cleared
FaultEnterCommand.Enabled = False
FaultClearCommand.Enabled = True
ClearAll.Enabled = True
End Sub

Private Sub ActiveFaultList_GotFocus()
'FocusOnActiveFaultList = True

End Sub


Private Sub ActiveFaultList_LostFocus()
'FocusOnActiveFaultList = False

End Sub


Private Sub Bearing_check_Click(Index As Integer)

End Sub

Private Sub Bearing_severity_lbl_Click(Index As Integer)

If Bearing_severity_lbl(Index).BorderStyle = 1 Then
  Bearing_severity_lbl(Index).BorderStyle = 0
Else
  Bearing_severity_lbl(Index).BorderStyle = 1
End If

End Sub

Private Sub Cancel_Click()
Unload Me

End Sub

Private Sub ClearAll_Click()
Dim i As Integer


' 26/3/97  Clear a ALL Faults...
'If Option1(0) Then 'we are looking at active faults


' 25/3/98  The FAULT array is in model.dll
'    For i = 1 To 500
'       Fault(i) = False
'    Next
    
    Call setFAULT(3, 1)

    '4/11/97 call logging not AddToLog
     ' If (IncludeFaultsInLog = True) Then
               ' Call AddToLog("All Faults Cleared by Instructor " & Time)
        ' End If
     Call Logging("All Faults Cleared by Instructor ", LogFaults)

SelectedFault.text = ClearText
'Else 'we are looking at Scheduled Faults, so clear scheduled fault list

    FaultScheduleIndex = 0
'End If

'deselect any item from list box so it can be updated
Call Timer1_Timer
End Sub

Private Sub Command1_Click()
Dim i As Integer

'find the FaultIndex and clear it from the selected fault list
    If FaultScheduleIndex > 0 Then
        For i = 1 To FaultScheduleIndex
             If FaultSchedule(i, 1) = Val(SelectedFault.text) Then
             
 ' 24/2/98  Does not exist so remove for now
 '11/10/99 put back in
              Call frm_Host.ClearScheduledFault(i)
              SelectedFault.text = ClearText
              
            End If
        Next
    End If
    
End Sub

Private Sub Command2_Click()

    FaultScheduleIndex = 0
    
End Sub

Private Sub Description_Click()
Dim ValueRequiredString As String * 1
Dim length As Integer
Dim i As Integer
Dim SpecialFault As Byte

'---------------------------------------------------------------------------------------------------------------------
'FUNCTION:
'Puts selected fault from Description list into Selected fault Box
'If fault is already entered, show it in black and enable the clear buttons, disable the enter
'Show it in grey and enable Enter button and disable clear buttons
'If it is a special fault, show the spin button, value box and current value
'-------------------------------------------------------------------------------------------------------------------


'Check if selected fault is already entered, if yes then just offer clear functions and put Selected Text in bold

SelectedFault.text = Description.List(Description.ListIndex)
ThisFaultnumber = Val(SelectedFault.text)

' === 10/04/02 TBD ===
'If ThisFaultnumber = 1164 Or ThisFaultnumber = 1165 Then
'    EARTH_LEAK.ZOrder
'End If


'Debug.Print "SelectedFault.Text = " & SelectedFault.Text
'Debug.Print "ThisFaultnumber = " & ThisFaultnumber
If ThisFaultnumber > 0 And ThisFaultnumber < BiggestFaultNumber Then

  If Fault(ThisFaultnumber) Then        'ie Fault already entered, offer the cler button
    FaultEnterCommand.Enabled = False
    FaultClearCommand.Enabled = True
    'ClearAll.Enabled = True
    SelectedFault.Enabled = True
  Else
  'Enable Enter button and disable clear buttons
        FaultEnterCommand.Enabled = True    'enables ENTRY
        FaultClearCommand.Enabled = False  'disables CLEAR
        
        'If fault requires additional value input, show box but keep text grey until accepted
        
  End If
  
  
      ' == 14/7/00 =====================
    
      If FAULT_severity_needed(ThisFaultnumber) Then
    
        ' Find the CURRENT severity value (from the model)...
        Call get_FAULT_severity(ThisFaultnumber, severity)
           
        ' Set the scroll value to the CURRENT severity value...
        Severity_scroll.Value = severity
      
        Severity_label.Visible = True
      
        Severity_scroll.Visible = True
        Severity_scroll.Enabled = True
        Severity_frame.Visible = True
   
    ' 24/7/00  added 321
        If ThisFaultnumber = 321 Or ThisFaultnumber = 405 Or ThisFaultnumber = 406 _
          Or ThisFaultnumber = 609 Or ThisFaultnumber = 610 _
          Or ThisFaultnumber = 619 Or ThisFaultnumber = 809 Then
          
            Sticky_valve = True
          Severity_frame.Caption = "Sticking Valve"
          Severity_label.Caption = "Valve Sticking point  = " & Format(severity, "##0") & "%"
        
        Else
        
            Sticky_valve = False
          Severity_frame.Caption = "Fault Severity"
          Severity_label.Caption = "Fault Severity = " & Format(severity, "##0") & "%"
        End If
        
        If ThisFaultnumber = 811 Then    '  BEARING FAULT ( There are 25 bearings )
        
          Bearing_frame.Visible = True
          Call get_BEARING_severity(BEARING_severity(1))   ' From the MODEL
          For i = 1 To 25
            Bearing_severity_lbl(i).Caption = Format(BEARING_severity(i), "##0") & "%"
          Next i

        Else
           Bearing_frame.Visible = False
        End If
        
      Else
    
        Severity_label.Visible = False
      
        Severity_scroll.Visible = False
        Severity_scroll.Enabled = False
        
        Severity_frame.Visible = False
        
        Bearing_frame.Visible = False
    
      End If
    
    '================================

  
End If

End Sub

Sub Description_DblClick()
Dim DummyText As String

'does the same as Fault Entry except if it is a special fault (requiring a value input).
' In this case Fault Entry must be used

FaultEntryflag = True

'if fault not already entered, call fault entry routine
DummyText = Description.List(Description.ListIndex)
ThisFaultnumber = Val(DummyText)

If Not Fault(ThisFaultnumber) Then
    Call FaultEnterCommand_Click
End If


End Sub


Private Sub FaultEnter_Click()

End Sub





Private Sub FaultClearCommand_Click()
Dim Offset As Integer
Dim i As Integer
 'Clear selected fault text  to show fault is cleared
 
'SelectedFault.Enabled = False
'If Option1(0) Then 'we are looking at active faults
    ThisFaultnumber = Val(SelectedFault.text)
    'Debug.Print "in clear SelectedFault.Text = " & SelectedFault.Text & "; Val = " & ThisFaultnumber
    If ThisFaultnumber > 0 And ThisFaultnumber < BiggestFaultNumber Then
    
    
' 25/3/98  The FAULT array is in model.dll
'''          Fault(ThisFaultnumber) = False
          
             Call setFAULT(0, ThisFaultnumber)

        
        '4/11/97 log using call logging not call addtolog
            ' If (frmMainInterface.mnuLogfaults.Checked) Then
                'Call AddToLog(SelectedFault.Text & " Cleared by Instructor " & Time)
            'End If
            Call Logging(SelectedFault.text & " Cleared by Instructor ", LogFaults)
            SelectedFault.text = ClearText
    End If
'Else
'find the FaultIndex and clear it from the selected fault list
'    If FaultScheduleIndex > 0 Then
'        For i = 1 To FaultScheduleIndex
'             If FaultSchedule(i, 1) = Val(SelectedFault.Text) Then
'                    Call frmMainInterface.ClearScheduledFault(i)
'            End If
'        Next
'    End If
'End If
 
 'update Active Fault List immediately
 Call Timer1_Timer


End Sub

Sub FaultEnterCommand_Click()
'---------------------------------------------------------------------------------------------------------------------
'FUNCTION:
'This routine entered EITHER after enter button clicked OR by double clicking a fault in the description list
'I
'-------------------------------------------------------------------------------------------------------------------



'does the same as DblClick and enters faults

Dim ValueRequiredString As String * 1
Dim length As Integer
Dim i As Integer
Dim msg As String
Dim Title As String
Dim Response As Integer
Dim DummyText As String
Dim SpecialFault As Byte
Dim Freq As String
Dim Volts As String
Dim Offset As Integer




'if entered from double click, see if there is a ? in the text string, if yes then fault will require additional value to be set
If FaultEntryflag Then
    FaultEntryflag = False
    DummyText = Description.List(Description.ListIndex)
    length = Len(DummyText)
 
     
End If




'Get the fault number
SelectedFault.Enabled = True
ThisFaultnumber = Val(SelectedFault.text)

'24/11/97 If (ThisFaultnumber > 0 And ThisFaultnumber < 460) Then
    '24/22/97 Fault(ThisFaultnumber) = True                   'Set the fault
    
If Val(FaultScheduleTime.text) = 0 Then 'enter fault immediately
    If (ThisFaultnumber > 0 And ThisFaultnumber < BiggestFaultNumber) Then
    
   
'  25/3/98  The FAULT array is now in model.dll
'''''    Fault(ThisFaultnumber) = True                   'Set the fault
        
          ' 25/3/98   Send the action to the model
        Call setFAULT(1, ThisFaultnumber)

   
'4/11/97 Catch the time
        fault_time(ThisFaultnumber) = Time
    
    '4/11/97 call logging not addtolog
    'If (frmMainInterface.mnuLogfaults.Checked) Then
                'Call AddToLog(SelectedFault.Text & " Entered by Instructor " & Time)
   ' End If
        'Call Logging(SelectedFault.Text & " Entered ", LogFaults)
    End If
'24/11/97
Else ' put it in a FaultSchedule and wait for timer on main form to activate it
    FaultScheduleIndex = FaultScheduleIndex + 1
    'debug.print " FaultScheduleIndex set to : " & FaultScheduleIndex; ""
    If FaultScheduleIndex <= 100 Then
        FaultSchedule(FaultScheduleIndex, 1) = ThisFaultnumber
        'Debug.Print " FaultSchedule(" & FaultScheduleIndex & "),1  set to : " & ThisFaultnumber

        FaultSchedule(FaultScheduleIndex, 2) = Val(FaultScheduleTime.text) ' set  delay in mins
        FaultSchedule(FaultScheduleIndex, 3) = Val(FaultScheduleTime.text) * 60  'set  delay in seconds  left
        'Debug.Print " FaultSchedule(" & FaultScheduleIndex & "),1  set to : " & Val(FaultScheduleTime.Text) * 60
        FaultScheduleTime.text = str(0)
        
    Else
        'display a message  cause > 100 scheduled faults ?
    End If
End If

   

    
'End If

' === 17/7/00 =================
Severity_label.Visible = False
Severity_scroll.Visible = False
Severity_scroll.Enabled = False
Severity_frame.Visible = False
' === 17/7/00 end of edit ======



'clear the Selected Fault text and make value box and button invisible
SelectedFault.text = ClearText
Description.ListIndex = -1                                                  'deselect any item from list box

'Fault is entered so diable the Enter command and enable the clear
FaultEnterCommand.Enabled = False
FaultClearCommand.Enabled = True
 'update Active Fault List immediately
Call Timer1_Timer
End Sub

Private Sub FaultTypes_Click()
Dim i As Integer



' ==== 14/7/00 ====
Severity_label.Visible = False
      
Severity_scroll.Visible = False
Severity_scroll.Enabled = False
Severity_frame.Visible = False
' =================

Description.clear
'Select Case FaultTypes.ListIndex
i = FaultTypes.ListIndex

'debug.print " in faultTypesClick, Index = " & i & " FaultGroup(i, 1) = " & FaultGroup(i, 1) & "; FaultGroup(i, 2) = " & FaultGroup(i, 2); ""
'Case 0                      'Main Engine
    Call FillDescription(FaultGroup(i, 1), FaultGroup(i, 2))
 
    
'Case 1                      'FO
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))

'Case 2                      'HTCW
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))
    
'Case 3                      'LTCW
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))

'Case 4                      'SW
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))
  

'Case 5                      'MELO
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))
   
'Case 6                      'GBX
   'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))


'Case 7                      'CPP
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))

   
'Case 8                      'Comp Air
   ' Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))

'Case 9                      'Misc
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))

    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))
    'Call FillDescription(FaultGroup(0, 1), FaultGroup(0, 2))

'End Select




End Sub


Private Sub FaultValueScroll_Change()

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Long
Dim ii As Integer
Dim Filepathname As String
Dim LastLine As Byte
Dim FG1 As Integer
Dim FG2 As Integer

Dim counter As Integer
Dim ErrorFlag As Byte
Dim NumberofTypes As Integer
Dim MatchPosition As Integer
Dim FTypeStr As String
Dim MyChar As String * 1



' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


' === 4/12/00 RWJ ===
'  Active Fault List gets wiped clean if you exit then re-enter this form so
FaultUpdateRequired = True


' ==== 14/7/00 ====

Severity_label.Visible = False
Severity_scroll.Visible = False
Severity_scroll.Enabled = False

Severity_frame.Visible = False

Bearing_frame.Visible = False

' === 13/12/02 19.141 ===
FAULT_severity_needed(107) = True
FAULT_severity_needed(108) = True

      
FAULT_severity_needed(201) = True
FAULT_severity_needed(202) = True

FAULT_severity_needed(205) = True
FAULT_severity_needed(206) = True

' 24/7/00
FAULT_severity_needed(321) = True

FAULT_severity_needed(405) = True
FAULT_severity_needed(406) = True

' 24/7/00
FAULT_severity_needed(501) = True
FAULT_severity_needed(502) = True
FAULT_severity_needed(503) = True
FAULT_severity_needed(504) = True

FAULT_severity_needed(507) = True
FAULT_severity_needed(508) = True

FAULT_severity_needed(603) = True
FAULT_severity_needed(604) = True

FAULT_severity_needed(608) = True

FAULT_severity_needed(609) = True
FAULT_severity_needed(610) = True

FAULT_severity_needed(619) = True

' === 30/09/02 14.126 ===
FAULT_severity_needed(702) = True

FAULT_severity_needed(803) = True
FAULT_severity_needed(804) = True

FAULT_severity_needed(807) = True
FAULT_severity_needed(808) = True

FAULT_severity_needed(809) = True

FAULT_severity_needed(811) = True

FAULT_severity_needed(903) = True
FAULT_severity_needed(904) = True

FAULT_severity_needed(1115) = True
FAULT_severity_needed(1116) = True


' =================


ErrorFlag = False
 ClearText = "                                                                              "
 
'read in strings for fault page menu text until the end of file
    Call GetLanguageStrings("Faults.txt", ErrorFlag)



frm_faults.Caption = Language_Str(1)
Label4.Caption = Language_Str(2)
Label3.Caption = Language_Str(3)
Label7.Caption = Language_Str(4)

FaultClearCommand.Caption = Language_Str(5)
Command1.Caption = Language_Str(5)

ClearAll.Caption = Language_Str(6)
Command2.Caption = Language_Str(6)

Cancel.Caption = Language_Str(7)
Label1.Caption = Language_Str(8)
Label2.Caption = Language_Str(9)
FaultEnterCommand.Caption = Language_Str(10)

'25/11/97 added:
Label9.Caption = Language_Str(11)
Label8.Caption = Language_Str(12)

'read in strings for actual steam fault text until the end of file

' === 01/03/04 Note Fault 1163,1164,1165 are HARDWARE faults
' so take them OUT if this is PTT...
If PTT Then
    Call GetLanguageStrings("SteamFaultText_PTT.txt", ErrorFlag)
Else
    Call GetLanguageStrings("SteamFaultText.txt", ErrorFlag)
End If


If Not ErrorFlag Then
    NumberofTypes = 0
    counter = 1
    FG1 = 0
    FG2 = -1
    LastLine = False
    i = 0
    Do
        Language_Str(i) = Trim(Language_Str(i)) 'get rid of leading & trailing spaces
        FTypeStr = ""
        
        LastLine = Language_Str(i) Like "*END OF TEXT*"    ' Returns True if a match
        If LastLine = False Then
            MatchPosition = InStr(1, Language_Str(i), "{", 1)
            'If Language_Str(i) Like "*{*}*" Then 'it is a group title, so put it in type list and save next address
            If MatchPosition <> 0 Then 'it is a group title, so put it in type list and save next address
                
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change curly bracket for a space
                MatchPosition = InStr(1, Language_Str(i), "}", 1)
                Mid(Language_Str(i), MatchPosition, 1) = Space(1) 'change other curly bracket for a space

                
                
                
                
                
                
                'For j = 1 To Len(Language_Str(i))
                   ' Debug.Print "found group title: " & Language_Str(i)
                    'MyChar = Mid(Language_Str(i), j, 1)
                    'If MyChar <> "{" And MyChar <> "}" Then
                    '    FTypeStr = FTypeStr + MyChar
                    'End If
                'Next
                'FTypeStr = LTrim(Language_Str(i))
                
                
                Fault_Types(FG1) = Language_Str(i)
                'Fault_Types(FG1) = " " & FTypeStr

                'FaultGroup(FG1, 1) = i
                FaultGroup(FG1, 1) = counter

                'debug.print "set group for FG1 = " & FG1 & ",  " & Language_Str(i)
                FG1 = FG1 + 1
                'FG2 = FG2 + 1
                NumberofTypes = NumberofTypes + 1
            Else
                If Language_Str(i) Like "###*" Then           'it is a fault entry
                    'Debug.Print "found fault and  FG2 = " & FG2
                  If FG1 > 0 Then
                    'Debug.Print "FG2 = " & FG2 & "FOR :  " & Language_Str(i)
                    FaultGroup(FG1 - 1, 2) = i                  'it may be the last entry in this group. If not keep incrementing
                    FaultGroup(FG1 - 1, 2) = counter                  'it may be the last entry in this group. If not keep incrementing

                    Faultext(counter) = Language_Str(i)        'store the fault text sequentially in an array
                    ThisFaultnumber = Val(Language_Str(i))     'get the actual fault number from the text
                    FaultXRef(ThisFaultnumber) = counter       'fill in the x-ref for retrieval later
                    '27/2/98
                    If ThisFaultnumber > BiggestFaultNumber Then
                        BiggestFaultNumber = ThisFaultnumber
                    End If
                    counter = counter + 1
                    'added 27/2/98
                  
                  End If
                End If
            End If
        Else
            'debug.print "Found Last Line TRUE when i = " & i
            Exit Do
        End If
        i = i + 1
        If i > 2000 Then ' in case END OF TEXT is not found
            Exit Do
        End If
    Loop Until LastLine = True
End If

For i = 0 To 10
    'debug.print "i = " & i
    'debug.print "   FaultGroup(i, 1) = " & FaultGroup(i, 1) & " FaultGroup(i, 2) = " & FaultGroup(i, 2)
Next

For i = 78 To 85
    'debug.print "Language_Str(" & i & ") = " & Language_Str(i)
Next
For i = 1 To 30
    'Debug.Print "Faultext(" & i & ") = " & Faultext(i)
Next
For i = 120 To 150
    'Debug.Print "FaultXRef(" & i & ") = " & FaultXRef(i)
Next


'========================================
'24/11/97
FaultScheduleTime.text = 0

FaultTypes.clear
Description.clear

For i = 0 To NumberofTypes - 1
   FaultTypes.AddItem Fault_Types(i)
Next i

End Sub




Private Sub List1_Click()


End Sub




Private Sub Form_Unload(Cancel As Integer)
 Unload Me
 
End Sub

Private Sub Option1_Click(Index As Integer)

'If Index = 1 Then 'show scheduled faults
'    ActiveFaultList.Visible = False
'    ScheduledFaultList.Visible = True
'    Label3.Caption = "Scheduled Faults"
'Else
'    ActiveFaultList.Visible = True
'    ScheduledFaultList.Visible = False
'     Label3.Caption = "Active Faults"
'
'End If

End Sub

Private Sub Severity_scroll_Change()
Dim i As Integer


' 14/7/00
If Severity_scroll.Enabled Then

      severity = Severity_scroll.Value

      ' Tell the model about this...
      Call set_FAULT_severity(ThisFaultnumber, severity)

      
    If Sticky_valve Then
      Severity_frame.Caption = "Sticking Valve"
      Severity_label.Caption = "Valve Sticking point  = " & Format(severity, "##0") & "%"
    Else
      Severity_frame.Caption = "Fault Severity"
      Severity_label.Caption = "Fault Severity = " & Format(severity, "##0") & "%"
    End If
            
      
      
      
      ' ==== 20/7/00  BEARING FAULTS ================
      If ThisFaultnumber = 811 Then
      
        For i = 1 To 25
        
          If Bearing_severity_lbl(i).BorderStyle = 1 Then
            Bearing_severity_lbl(i).Caption = Format(severity, "##0") & "%"
            BEARING_severity(i) = severity
          End If
          
        Next i
        
              Call set_BEARING_severity(BEARING_severity(1))
              
      End If
      ' ==== 20/7/00  BEARING FAULTS end of edit =====
      
      

End If


End Sub

Private Sub Timer1_Timer()
Dim i As Integer

'calls Fill_active_fault_list to update list. Add new faults, remove old faults
'Enable buttons according to number of faults entered etc
Call Fill_Active_Fault_List(frm_faults)
Call Fill_Scheduled_Fault_List


' ==== 21/7/00  BEARING FAULTS ================
If ThisFaultnumber = 811 Then
      
        flash_bearing_severity = Not flash_bearing_severity
        
        For i = 1 To 25
        
          If Bearing_severity_lbl(i).BorderStyle = 1 Then
            If flash_bearing_severity Then
              Bearing_severity_lbl(i).ForeColor = &HFF&       ' RED
            Else
              Bearing_severity_lbl(i).ForeColor = &H80000008       ' BLACK
            End If
          Else
              Bearing_severity_lbl(i).ForeColor = &H80000008       ' BLACK
          End If
          
        Next i
                      
End If
' ==== 21/7/00  BEARING FAULTS end of edit =====



 'enable/disable clear all button as appropriate
     If NumberOfFaults > 0 Then
            ClearAll.Enabled = True
    Else
            ClearAll.Enabled = False
    End If
   'enable/disable clear  button as appropriate
   ' If ActiveFaultList.ListIndex > -1 Then
      '  FaultClearCommand.Enabled = True
   ' Else
    '    FaultClearCommand.Enabled = False
   ' End If
    
     'enable/disable enter button as appropriate
 '   If Description.ListIndex > -1 Then
    '    FaultEnterCommand.Enabled = True
   ' Else
     '   FaultEnterCommand.Enabled = False
 '   End If
Exit Sub

End Sub

Sub Fill_Scheduled_Fault_List()
Dim i As Integer
Dim NumberMins As Integer
Dim NumberSecs As Integer

'If FaultScheduleIndex > 0 And FaultScheduleIndex <> ScheduledFaultList.ListCount Then

ScheduledFaultList.clear
    If FaultScheduleIndex > 0 Then
        For i = 1 To FaultScheduleIndex
            NumberMins = Int(FaultSchedule(i, 3) / 60)
            NumberSecs = FaultSchedule(i, 3) Mod 60
            ScheduledFaultList.AddItem (FaultSchedule(i, 1)) & " Scheduled in " & NumberMins & " mins " & NumberSecs & " secs"
        Next
    Else
             ScheduledFaultList.AddItem "No Scheduled Faults"
    End If
    
End Sub

Private Sub UpDown1_DownClick()
Dim NewTime As Integer

NewTime = Val(FaultScheduleTime.text) - 1
If NewTime < 0 Then NewTime = 0



FaultScheduleTime.text = str(NewTime)
End Sub

Private Sub UpDown1_UpClick()
Dim NewTime As Integer

NewTime = Val(FaultScheduleTime.text) + 1
If NewTime > 300 Then NewTime = 300

FaultScheduleTime.text = str(NewTime)
End Sub
