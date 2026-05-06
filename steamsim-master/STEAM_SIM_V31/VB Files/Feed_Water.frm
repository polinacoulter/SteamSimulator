VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "HavenStretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Feed_Water 
   Caption         =   "Feed Water & Aux Steam"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   13680
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   9675
      Left            =   0
      Top             =   0
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   17066
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
         Left            =   4200
         TabIndex        =   247
         Text            =   "xx F"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox TLO_main_feed 
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
         Left            =   12300
         TabIndex        =   246
         Text            =   "xx F"
         Top             =   7620
         Width           =   795
      End
      Begin VB.TextBox TLO_main_feed 
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
         Left            =   11460
         TabIndex        =   245
         Text            =   "xx F"
         Top             =   7620
         Width           =   795
      End
      Begin VB.TextBox P_main_out 
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
         Left            =   2880
         TabIndex        =   244
         Text            =   "x.x Bar"
         Top             =   600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox PLO_total_feed_pump 
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
         Left            =   12300
         TabIndex        =   243
         Text            =   "x.x Bar"
         Top             =   7920
         Width           =   795
      End
      Begin VB.TextBox PLO_total_feed_pump 
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
         Left            =   11460
         TabIndex        =   242
         Text            =   "x.x Bar"
         Top             =   7920
         Width           =   795
      End
      Begin VB.CommandButton REMOTE_feed_control 
         BackColor       =   &H00C0C0C0&
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   9060
         Width           =   255
      End
      Begin VB.CommandButton REMOTE_feed_control 
         BackColor       =   &H0000FFFF&
         Caption         =   "R"
         Height          =   255
         Index           =   1
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   238
         Top             =   9060
         Width           =   255
      End
      Begin VB.TextBox P_main_out 
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
         Left            =   2880
         TabIndex        =   237
         Text            =   "x.x Bar"
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton AUTO_feed_pump 
         BackColor       =   &H0000FFFF&
         Caption         =   "A"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   5940
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton AUTO_feed_pump 
         BackColor       =   &H0000FFFF&
         Caption         =   "A"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   4980
         Visible         =   0   'False
         Width           =   255
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
         Left            =   10760
         TabIndex        =   228
         Top             =   120
         Width           =   375
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
         Left            =   7920
         TabIndex        =   227
         Top             =   80
         Width           =   375
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
         Left            =   2800
         TabIndex        =   226
         Top             =   2880
         Width           =   375
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
         Left            =   3480
         TabIndex        =   225
         Top             =   2460
         Width           =   375
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   12360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   8760
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton sbyPUMP 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   11760
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   8760
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
         Index           =   4
         Left            =   12360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   8520
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
         Index           =   3
         Left            =   11760
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   8520
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
         Index           =   4
         Left            =   12360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   8280
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
         Index           =   3
         Left            =   11760
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   8280
         UseMaskColor    =   -1  'True
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2940
         TabIndex        =   213
         Top             =   5940
         Width           =   315
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   135
         Left            =   5520
         Max             =   1000
         TabIndex        =   212
         Top             =   7680
         Width           =   1215
      End
      Begin VB.TextBox T_HP_htr 
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
         Left            =   3900
         TabIndex        =   211
         Text            =   "xx F"
         Top             =   7200
         Width           =   795
      End
      Begin VB.TextBox T_DFT 
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
         Left            =   12300
         TabIndex        =   210
         Text            =   "xx F"
         Top             =   4080
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
         Left            =   4080
         TabIndex        =   207
         Text            =   "x.x Bar"
         Top             =   4620
         Width           =   795
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
         Left            =   9840
         TabIndex        =   200
         Text            =   "x.x Bar"
         Top             =   1380
         Width           =   795
      End
      Begin VB.VScrollBar VScroll6 
         Height          =   1335
         Left            =   6600
         Max             =   1000
         TabIndex        =   199
         Top             =   1080
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
         Left            =   7020
         TabIndex        =   198
         Text            =   "x.x Bar"
         Top             =   1080
         Width           =   795
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   1335
         Left            =   6360
         Max             =   1000
         TabIndex        =   197
         Top             =   1200
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
         Left            =   5220
         TabIndex        =   196
         Text            =   "x.x Bar"
         Top             =   1680
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
         Left            =   1140
         TabIndex        =   195
         Text            =   "x.x Bar"
         Top             =   3180
         Width           =   795
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
         Left            =   4200
         TabIndex        =   194
         Text            =   "x.x Bar"
         Top             =   5700
         Width           =   795
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   1335
         Left            =   780
         Max             =   1000
         TabIndex        =   193
         Top             =   3360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   1335
         Left            =   7740
         Max             =   1000
         TabIndex        =   192
         Top             =   2220
         Width           =   135
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   135
         Left            =   2220
         Max             =   1000
         TabIndex        =   191
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Z_DFT 
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
         Left            =   12300
         TabIndex        =   190
         Text            =   "20 %"
         Top             =   3780
         Width           =   795
      End
      Begin VB.TextBox F_aux_feed 
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
         Left            =   2400
         TabIndex        =   187
         Text            =   "xx.x kg/s"
         Top             =   8640
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_aux_feed 
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
         Left            =   1680
         TabIndex        =   186
         Text            =   "xx.x kg/s"
         Top             =   7920
         Visible         =   0   'False
         Width           =   1035
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
         Index           =   2
         Left            =   2340
         TabIndex        =   185
         Text            =   "xx.x kg/s"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   1620
         TabIndex        =   184
         Text            =   "xx.x kg/s"
         Top             =   6480
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_aux 
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
         Left            =   7260
         TabIndex        =   183
         Text            =   "xx.x kg/s"
         Top             =   7980
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_aux 
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
         Left            =   5700
         TabIndex        =   182
         Text            =   "xx.x kg/s"
         Top             =   7980
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_recirc 
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
         Left            =   11580
         TabIndex        =   181
         Text            =   "xx.x kg/s"
         Top             =   5100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF_recirc 
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
         Left            =   12120
         TabIndex        =   180
         Text            =   "xx.x kg/s"
         Top             =   5460
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF 
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
         Left            =   4980
         TabIndex        =   179
         Text            =   "xx.x kg/s"
         Top             =   6540
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_MF 
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
         Left            =   4920
         TabIndex        =   178
         Text            =   "xx.x kg/s"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_main 
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
         Left            =   9900
         TabIndex        =   177
         Text            =   "xx.x kg/s"
         Top             =   6600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox F_main 
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
         Left            =   9900
         TabIndex        =   176
         Text            =   "xx.x kg/s"
         Top             =   7260
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox P_AF 
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
         Left            =   3900
         TabIndex        =   175
         Text            =   "x.x Bar"
         Top             =   7980
         Width           =   795
      End
      Begin VB.TextBox P_MF 
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
         Left            =   3900
         TabIndex        =   174
         Text            =   "x.x Bar"
         Top             =   6780
         Width           =   795
      End
      Begin VB.TextBox P_main 
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
         Left            =   6300
         TabIndex        =   173
         Text            =   "x.x Bar"
         Top             =   7260
         Width           =   795
      End
      Begin VB.TextBox P_main 
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
         Left            =   6300
         TabIndex        =   172
         Text            =   "x.x Bar"
         Top             =   6660
         Width           =   795
      End
      Begin VB.CommandButton CLOSE_VV 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   2
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   7020
         Width           =   675
      End
      Begin VB.CommandButton CLOSE_VV 
         BackColor       =   &H008080FF&
         Caption         =   "CLOSE"
         Height          =   255
         Index           =   1
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   6360
         Width           =   675
      End
      Begin VB.CommandButton OPEN_VV 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OPEN"
         Height          =   255
         Index           =   2
         Left            =   11220
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   7020
         Width           =   615
      End
      Begin VB.CommandButton OPEN_VV 
         BackColor       =   &H0000FF00&
         Caption         =   "OPEN"
         Height          =   255
         Index           =   1
         Left            =   11220
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Port Boiler"
         Height          =   315
         Index           =   1
         Left            =   900
         TabIndex        =   160
         Top             =   8460
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stbd Boiler"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   159
         Top             =   7920
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Port Boiler"
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   158
         Top             =   7260
         Width           =   1095
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
         Left            =   7140
         TabIndex        =   156
         Text            =   "P"
         Top             =   2130
         Width           =   195
      End
      Begin VB.CommandButton AUTO_VV 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
         Height          =   255
         Index           =   2
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   7020
         Width           =   615
      End
      Begin VB.CommandButton AUTO_VV 
         BackColor       =   &H0000FFFF&
         Caption         =   "AUTO"
         Height          =   255
         Index           =   1
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   6360
         Width           =   615
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   135
         Left            =   7200
         Max             =   1000
         TabIndex        =   129
         Top             =   5640
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   135
         Left            =   6480
         Max             =   1000
         TabIndex        =   128
         Top             =   4560
         Width           =   1215
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
         Left            =   6840
         TabIndex        =   126
         Text            =   " xx psi"
         Top             =   120
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
         Left            =   4200
         TabIndex        =   125
         Text            =   "xx psi"
         Top             =   120
         Width           =   675
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   6
         Left            =   4920
         TabIndex        =   6
         Top             =   2155
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
         Index           =   22
         Left            =   4920
         TabIndex        =   41
         Top             =   2925
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
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   780
         Top             =   2700
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
         Height          =   255
         Left            =   5100
         TabIndex        =   101
         Text            =   "P"
         Top             =   3300
         Width           =   195
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   135
         Left            =   5880
         Max             =   1000
         TabIndex        =   100
         Top             =   7080
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   135
         Left            =   6360
         Max             =   1000
         TabIndex        =   99
         Top             =   8640
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         Left            =   11280
         Max             =   1000
         TabIndex        =   96
         Top             =   1620
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   3300
         Max             =   1000
         TabIndex        =   95
         Top             =   4320
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1155
         Left            =   6000
         Max             =   1000
         TabIndex        =   94
         Top             =   2220
         Width           =   135
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   8160
         Max             =   1000
         TabIndex        =   93
         Top             =   1680
         Width           =   1215
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
         Height          =   255
         Left            =   3660
         TabIndex        =   92
         Text            =   "P"
         Top             =   4500
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
         Height          =   240
         Left            =   10740
         TabIndex        =   90
         Text            =   "P"
         Top             =   765
         Width           =   195
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
         Height          =   240
         Left            =   9720
         TabIndex        =   87
         Text            =   "P"
         Top             =   765
         Width           =   195
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
         Height          =   255
         Left            =   4680
         TabIndex        =   86
         Text            =   "P"
         Top             =   3420
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stbd Boiler"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   6720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Condensate System"
         Height          =   675
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2100
         Width           =   1095
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
         Left            =   9720
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   8880
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
         Left            =   9300
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   8880
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
         Left            =   8820
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   8880
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox P_aux_feed_out 
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
         Left            =   7860
         TabIndex        =   13
         Text            =   "75 Bar"
         Top             =   8520
         Width           =   735
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   0
         Left            =   8880
         TabIndex        =   0
         Top             =   8040
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
         Index           =   1
         Left            =   9540
         TabIndex        =   1
         Top             =   6780
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
         Index           =   2
         Left            =   9780
         TabIndex        =   3
         Top             =   8280
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
         Index           =   5
         Left            =   5880
         TabIndex        =   4
         Top             =   7140
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         BottomColour    =   32768
         OrientVertical  =   0   'False
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   13
         Left            =   6060
         TabIndex        =   7
         Top             =   6060
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
         Index           =   14
         Left            =   6060
         TabIndex        =   8
         Top             =   5160
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
         Index           =   15
         Left            =   9300
         TabIndex        =   9
         Top             =   6540
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   9240
         TabIndex        =   18
         Top             =   8325
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   8940
         TabIndex        =   19
         Top             =   6840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   4
         Left            =   7560
         TabIndex        =   21
         Top             =   480
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
         Height          =   315
         Index           =   8
         Left            =   2640
         TabIndex        =   22
         Top             =   3240
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
         Index           =   9
         Left            =   5640
         TabIndex        =   23
         Top             =   720
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
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   10
         Left            =   11520
         TabIndex        =   24
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   2
         Left            =   8100
         TabIndex        =   29
         Top             =   7440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         OrientLeft      =   0   'False
         LineCol         =   16711680
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   17
         Left            =   8940
         TabIndex        =   30
         Top             =   7380
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
         Index           =   3
         Left            =   5880
         TabIndex        =   2
         Top             =   6540
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
         Height          =   255
         Index           =   16
         Left            =   8460
         TabIndex        =   32
         Top             =   7080
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   18
         Left            =   4060
         TabIndex        =   38
         Top             =   3370
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
         Index           =   19
         Left            =   3300
         TabIndex        =   39
         Top             =   3370
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
         Left            =   3660
         TabIndex        =   40
         Top             =   3840
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
         Height          =   375
         Index           =   23
         Left            =   5580
         TabIndex        =   42
         Top             =   2580
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
         Height          =   375
         Index           =   24
         Left            =   4920
         TabIndex        =   43
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
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   25
         Left            =   9240
         TabIndex        =   44
         Top             =   3540
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
         Index           =   26
         Left            =   8400
         TabIndex        =   45
         Top             =   3540
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
         Height          =   435
         Index           =   27
         Left            =   7560
         TabIndex        =   46
         Top             =   5160
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
         Index           =   29
         Left            =   7560
         TabIndex        =   47
         Top             =   6060
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
         Index           =   30
         Left            =   6420
         TabIndex        =   48
         Top             =   6060
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
         Index           =   31
         Left            =   6420
         TabIndex        =   49
         Top             =   5160
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
         Index           =   32
         Left            =   6720
         TabIndex        =   50
         Top             =   4620
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
         Index           =   34
         Left            =   5460
         TabIndex        =   52
         Top             =   3660
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
         Index           =   35
         Left            =   5460
         TabIndex        =   53
         Top             =   4200
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
         Index           =   36
         Left            =   3720
         TabIndex        =   54
         Top             =   5760
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
         Height          =   255
         Index           =   40
         Left            =   9180
         TabIndex        =   58
         Top             =   600
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
         Index           =   41
         Left            =   8820
         TabIndex        =   59
         Top             =   720
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
         Index           =   42
         Left            =   8160
         TabIndex        =   60
         Top             =   720
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
         Left            =   8580
         TabIndex        =   61
         Top             =   1200
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
         Height          =   1335
         Index           =   44
         Left            =   12060
         TabIndex        =   62
         Top             =   3000
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   2355
         TopColour       =   65280
         BottomColour    =   65280
         LineWidth       =   6
         OrientVertical  =   0   'False
         ColourClosed    =   16711680
         BackColour      =   12632256
         ValveType       =   9
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   28
         Left            =   12300
         TabIndex        =   64
         Top             =   420
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
         Index           =   37
         Left            =   11100
         TabIndex        =   65
         Top             =   720
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
         Index           =   38
         Left            =   11760
         TabIndex        =   66
         Top             =   720
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
         Left            =   11520
         TabIndex        =   67
         Top             =   1140
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
         Index           =   45
         Left            =   3090
         TabIndex        =   69
         Top             =   4580
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
         Index           =   46
         Left            =   2340
         TabIndex        =   70
         Top             =   4580
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
         Left            =   2700
         TabIndex        =   71
         Top             =   5040
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
         Index           =   48
         Left            =   7980
         TabIndex        =   73
         Top             =   4200
         Visible         =   0   'False
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
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   315
         Index           =   12
         Left            =   7140
         TabIndex        =   97
         Top             =   6960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   51
         Left            =   6780
         TabIndex        =   98
         Top             =   7860
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         TopColour       =   32768
         BottomColour    =   32768
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   7
         Left            =   6720
         TabIndex        =   114
         Top             =   2220
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
         Index           =   11
         Left            =   6720
         TabIndex        =   115
         Top             =   2880
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
         Left            =   7320
         TabIndex        =   116
         Top             =   2640
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
         Left            =   6360
         TabIndex        =   127
         Top             =   3060
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
         Index           =   52
         Left            =   4140
         TabIndex        =   134
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   767
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
         Height          =   255
         Index           =   53
         Left            =   2400
         TabIndex        =   135
         Top             =   4140
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
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   3
         Left            =   9360
         TabIndex        =   214
         Top             =   5280
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   65535
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   255
         Index           =   4
         Left            =   8640
         TabIndex        =   215
         Top             =   6120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         OrientLeft      =   0   'False
         LineCol         =   65535
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   54
         Left            =   8760
         TabIndex        =   229
         Top             =   6540
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   255
         Index           =   55
         Left            =   7980
         TabIndex        =   230
         Top             =   7080
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         BottomColour    =   255
         EnabledTopLeft  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   56
         Left            =   4440
         TabIndex        =   231
         Top             =   5000
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
         Height          =   315
         Index           =   49
         Left            =   6840
         TabIndex        =   74
         Top             =   6420
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         TopColour       =   32768
         BottomColour    =   32768
         OrientVertical  =   0   'False
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   32768
         BackColour      =   12632256
         ValveType       =   10
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   33
         Left            =   6780
         TabIndex        =   51
         Top             =   5580
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
         Index           =   85
         Left            =   11340
         TabIndex        =   248
         Top             =   5760
         Visible         =   0   'False
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
         Index           =   86
         Left            =   11280
         TabIndex        =   249
         Top             =   7260
         Visible         =   0   'False
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
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   8
         X1              =   5040
         X2              =   4320
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   3720
         Picture         =   "Feed_Water.frx":0000
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   255
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   0
         X1              =   3840
         X2              =   3420
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Label Label30 
         Caption         =   "  35 psi  Regulated"
         Height          =   195
         Left            =   2880
         TabIndex        =   83
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "Auxilliary Exhaust Steam"
         Height          =   195
         Left            =   1080
         TabIndex        =   258
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   4
         X1              =   5760
         X2              =   8100
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Label BYPASS_feed_2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7140
         TabIndex        =   257
         Top             =   5820
         Width           =   555
      End
      Begin VB.Label BYPASS_feed_1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7080
         TabIndex        =   256
         Top             =   4860
         Width           =   555
      End
      Begin VB.Label BYPASS_35_dump 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7200
         TabIndex        =   255
         Top             =   3060
         Width           =   555
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   2
         X1              =   3600
         X2              =   3000
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   27
         X1              =   3600
         X2              =   3600
         Y1              =   5340
         Y2              =   4860
      End
      Begin VB.Label BYPASS_atom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3060
         TabIndex        =   254
         Top             =   5100
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   9
         X1              =   4560
         X2              =   4560
         Y1              =   4140
         Y2              =   3660
      End
      Begin VB.Label BYPASS_875_150 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4020
         TabIndex        =   253
         Top             =   3900
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   22
         X1              =   12420
         X2              =   12420
         Y1              =   1440
         Y2              =   1020
      End
      Begin VB.Label BYPASS_875_35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11880
         TabIndex        =   251
         Top             =   1140
         Width           =   555
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   11
         X1              =   9420
         X2              =   9420
         Y1              =   1500
         Y2              =   1020
      End
      Begin VB.Label BYPASS_IP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8880
         TabIndex        =   250
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Tripped_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Tripped"
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   241
         Top             =   4740
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Left            =   9060
         TabIndex        =   240
         Top             =   5100
         Width           =   255
      End
      Begin VB.Image Image25 
         Height          =   255
         Left            =   4080
         Picture         =   "Feed_Water.frx":030A
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   255
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Hotel Services"
         Height          =   195
         Left            =   3720
         TabIndex        =   232
         Top             =   5400
         Width           =   1155
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   0
         X1              =   4200
         X2              =   5040
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label AUX_2 
         Caption         =   "# 2"
         Height          =   195
         Left            =   12840
         TabIndex        =   224
         Top             =   8280
         Width           =   255
      End
      Begin VB.Label AUX_1 
         Caption         =   "# 1"
         Height          =   195
         Left            =   11460
         TabIndex        =   223
         Top             =   8280
         Width           =   255
      End
      Begin VB.Label AUX_LO_label 
         Alignment       =   2  'Center
         Caption         =   "Feed Pump Aux LO"
         Height          =   195
         Left            =   11520
         TabIndex        =   222
         Top             =   9360
         Width           =   1575
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "HP Regulator"
         Height          =   195
         Left            =   3780
         TabIndex        =   206
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label VV_pos_875_150_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3660
         TabIndex        =   205
         Top             =   3060
         Width           =   555
      End
      Begin VB.Label VV_pos_HP_reg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   204
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label VV_pos_dump_reg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   203
         Top             =   2640
         Width           =   555
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   9540
         X2              =   8700
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label VV_pos_IP_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8460
         TabIndex        =   202
         Top             =   240
         Width           =   555
      End
      Begin VB.Label VV_pos_875_35_reg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11340
         TabIndex        =   201
         Top             =   360
         Width           =   555
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   4620
         X2              =   3840
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label VV_feed_aux_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   7260
         TabIndex        =   189
         Top             =   7500
         Width           =   555
      End
      Begin VB.Label VV_feed_aux_pos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "63 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   6540
         TabIndex        =   188
         Top             =   8400
         Width           =   555
      End
      Begin VB.Label RECIRC_label 
         Alignment       =   2  'Center
         Caption         =   "Reciculation Valve # 2"
         Height          =   195
         Index           =   2
         Left            =   11280
         TabIndex        =   171
         Top             =   6780
         Width           =   1755
      End
      Begin VB.Label RECIRC_label 
         Alignment       =   2  'Center
         Caption         =   "Reciculation Valve # 1"
         Height          =   195
         Index           =   1
         Left            =   11220
         TabIndex        =   169
         Top             =   6120
         Width           =   1755
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "DFT Water level"
         Height          =   615
         Left            =   12300
         TabIndex        =   161
         Top             =   3120
         Width           =   555
      End
      Begin VB.Image Image24 
         Height          =   255
         Left            =   2040
         Picture         =   "Feed_Water.frx":0614
         Stretch         =   -1  'True
         Top             =   8460
         Width           =   255
      End
      Begin VB.Image Image23 
         Height          =   255
         Left            =   1380
         Picture         =   "Feed_Water.frx":091E
         Stretch         =   -1  'True
         Top             =   7920
         Width           =   255
      End
      Begin VB.Image Image22 
         Height          =   255
         Left            =   2040
         Picture         =   "Feed_Water.frx":0C28
         Stretch         =   -1  'True
         Top             =   7260
         Width           =   255
      End
      Begin VB.Image Image20 
         Height          =   255
         Left            =   1380
         Picture         =   "Feed_Water.frx":0F32
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   31
         X1              =   2040
         X2              =   2760
         Y1              =   8580
         Y2              =   8580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   30
         X1              =   2040
         X2              =   2760
         Y1              =   7380
         Y2              =   7380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   29
         X1              =   1380
         X2              =   2760
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   19
         X1              =   1380
         X2              =   2760
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Image Image21 
         Height          =   255
         Left            =   10140
         Picture         =   "Feed_Water.frx":123C
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   7
         X1              =   2760
         X2              =   2760
         Y1              =   8580
         Y2              =   8040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   2760
         X2              =   2760
         Y1              =   7380
         Y2              =   6840
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "HP Heater"
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
         Height          =   255
         Left            =   3960
         TabIndex        =   157
         Top             =   7560
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   18
         X1              =   6840
         X2              =   6840
         Y1              =   2100
         Y2              =   2340
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   17
         X1              =   7260
         X2              =   7020
         Y1              =   2460
         Y2              =   2760
      End
      Begin VB.Label Label17 
         Caption         =   "No 1"
         Height          =   255
         Left            =   9420
         TabIndex        =   11
         Top             =   7260
         Width           =   495
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   255
         Left            =   8280
         TabIndex        =   148
         Top             =   6060
         Width           =   255
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
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
         Left            =   8220
         TabIndex        =   142
         Top             =   60
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   5
         X1              =   10680
         X2              =   6840
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label41 
         Caption         =   "To Make Up Feed Regulator"
         Height          =   675
         Left            =   10080
         TabIndex        =   124
         Top             =   4560
         Width           =   795
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   5
         Left            =   10260
         Picture         =   "Feed_Water.frx":1546
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   255
      End
      Begin VB.Line Line17 
         BorderStyle     =   3  'Dot
         X1              =   10380
         X2              =   10380
         Y1              =   3540
         Y2              =   4500
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   10440
         TabIndex        =   123
         Top             =   3540
         Width           =   195
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LT"
         Height          =   255
         Left            =   10440
         TabIndex        =   122
         Top             =   2940
         Width           =   195
      End
      Begin VB.Shape Shape32 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   10380
         Top             =   2880
         Width           =   315
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   17
         X1              =   9420
         X2              =   8940
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   9720
         X2              =   9720
         Y1              =   6480
         Y2              =   6300
      End
      Begin VB.Line Line7 
         X1              =   9600
         X2              =   9780
         Y1              =   6360
         Y2              =   6240
      End
      Begin VB.Line Line6 
         X1              =   9660
         X2              =   9840
         Y1              =   5640
         Y2              =   5520
      End
      Begin VB.Line Line5 
         X1              =   10980
         X2              =   11280
         Y1              =   5820
         Y2              =   5700
      End
      Begin VB.Line Line2 
         X1              =   10980
         X2              =   11280
         Y1              =   5580
         Y2              =   5460
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   3
         Left            =   7920
         Picture         =   "Feed_Water.frx":1850
         Stretch         =   -1  'True
         Top             =   2220
         Width           =   255
      End
      Begin VB.Image Image18 
         Height          =   255
         Left            =   6720
         Picture         =   "Feed_Water.frx":1B5A
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   38
         X1              =   7380
         X2              =   7380
         Y1              =   2100
         Y2              =   3420
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   39
         X1              =   6840
         X2              =   7380
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "To Main Condenser"
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
         Left            =   6900
         TabIndex        =   121
         Top             =   3480
         Width           =   795
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1395
         Index           =   1
         Left            =   9180
         Top             =   5460
         Width           =   75
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   4
         Left            =   7530
         Picture         =   "Feed_Water.frx":1E64
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "To Evap."
         Height          =   435
         Left            =   7920
         TabIndex        =   113
         Top             =   2460
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   255
         Index           =   0
         Left            =   9240
         Picture         =   "Feed_Water.frx":216E
         Stretch         =   -1  'True
         Top             =   3180
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   3
         X1              =   8040
         X2              =   8040
         Y1              =   2100
         Y2              =   2520
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   1
         Left            =   4910
         Picture         =   "Feed_Water.frx":2478
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   2175
         Index           =   1
         Left            =   4995
         Top             =   300
         Width           =   75
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   16
         X1              =   2880
         X2              =   2880
         Y1              =   4440
         Y2              =   4500
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   3180
         X2              =   2880
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   8160
         Picture         =   "Feed_Water.frx":2782
         Stretch         =   -1  'True
         Top             =   5340
         Width           =   255
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   1
         X1              =   7080
         X2              =   9060
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   6
         X1              =   9360
         X2              =   9360
         Y1              =   2100
         Y2              =   5100
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   10140
         Picture         =   "Feed_Water.frx":2A8C
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   17
         X1              =   8040
         X2              =   9720
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   15
         X1              =   5460
         X2              =   5460
         Y1              =   2760
         Y2              =   3180
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   9480
         TabIndex        =   103
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   8700
         TabIndex        =   102
         Top             =   3300
         Width           =   195
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   14
         X1              =   5460
         X2              =   5280
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   13
         X1              =   5460
         X2              =   5460
         Y1              =   3120
         Y2              =   2760
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   12
         X1              =   5460
         X2              =   5280
         Y1              =   3180
         Y2              =   3420
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   27
         X1              =   7200
         X2              =   7200
         Y1              =   8340
         Y2              =   7560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   26
         X1              =   6840
         X2              =   6840
         Y1              =   8340
         Y2              =   7440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   25
         X1              =   7200
         X2              =   7200
         Y1              =   7320
         Y2              =   6840
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   11
         X1              =   2880
         X2              =   2880
         Y1              =   4440
         Y2              =   4500
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   10
         X1              =   3720
         X2              =   3120
         Y1              =   4620
         Y2              =   4440
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   2
         X1              =   7080
         X2              =   8220
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label Label7 
         Caption         =   "Turbo Feed Pump Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   91
         Top             =   1980
         Width           =   2415
      End
      Begin VB.Line Turbo_feed_line 
         BorderColor     =   &H00606060&
         BorderWidth     =   3
         Index           =   0
         X1              =   180
         X2              =   840
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   9
         Left            =   5040
         Top             =   3540
         Width           =   675
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   8
         Left            =   5040
         Top             =   2040
         Width           =   675
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   11640
         X2              =   11160
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   10800
         X2              =   11100
         Y1              =   900
         Y2              =   660
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   10
         X1              =   9180
         X2              =   11160
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   9780
         X2              =   9540
         Y1              =   900
         Y2              =   480
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   8700
         X2              =   8700
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "35 #"
         Height          =   315
         Left            =   9780
         TabIndex        =   89
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "150 #"
         Height          =   315
         Left            =   4560
         TabIndex        =   88
         Top             =   3720
         Width           =   435
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   3900
         X2              =   3840
         Y1              =   3300
         Y2              =   3360
      End
      Begin VB.Line control_line 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   4800
         X2              =   4620
         Y1              =   3540
         Y2              =   3300
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   855
         Index           =   5
         Left            =   7620
         Top             =   120
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   435
         Index           =   4
         Left            =   7920
         Top             =   1020
         Width           =   75
      End
      Begin VB.Shape UN_reg 
         BackColor       =   &H00808080&
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   3
         Left            =   7920
         Top             =   1440
         Width           =   675
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   5040
         Top             =   960
         Width           =   3195
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   0
         Left            =   180
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label32 
         Caption         =   "Bled Steam"
         Height          =   195
         Left            =   1020
         TabIndex        =   85
         Top             =   2280
         Width           =   1635
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   16
         X1              =   10920
         X2              =   11640
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   12480
         Picture         =   "Feed_Water.frx":2D96
         Stretch         =   -1  'True
         Top             =   900
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   15
         X1              =   12120
         X2              =   12840
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "KEY"
         Height          =   195
         Left            =   1320
         TabIndex        =   84
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "150 psi  Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   82
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "875 psi  Regulated"
         Height          =   195
         Left            =   1020
         TabIndex        =   81
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   24
         X1              =   840
         X2              =   180
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   8
         X1              =   840
         X2              =   180
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   14
         X1              =   180
         X2              =   840
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Steam_line_150 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   1
         X1              =   4380
         X2              =   3360
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   12
         X1              =   5040
         X2              =   5040
         Y1              =   6900
         Y2              =   3300
      End
      Begin VB.Image Image14 
         Height          =   255
         Left            =   1500
         Picture         =   "Feed_Water.frx":30A0
         Stretch         =   -1  'True
         Top             =   3540
         Width           =   255
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   1560
         X2              =   3420
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   7920
         TabIndex        =   80
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5760
         TabIndex        =   79
         Top             =   540
         Width           =   195
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   5340
         TabIndex        =   78
         Top             =   1380
         Width           =   195
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   6180
         TabIndex        =   77
         Top             =   5880
         Width           =   195
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   2
         Left            =   1800
         Picture         =   "Feed_Water.frx":33AA
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "875 #"
         Height          =   255
         Left            =   960
         TabIndex        =   76
         Top             =   3540
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         Height          =   255
         Left            =   6180
         TabIndex        =   75
         Top             =   4920
         Width           =   195
      End
      Begin VB.Label Label20 
         Caption         =   "Atomising Steam to  burners"
         Height          =   675
         Left            =   1200
         TabIndex        =   72
         Top             =   4920
         Width           =   795
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   28
         X1              =   2820
         X2              =   2100
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   26
         X1              =   2100
         X2              =   2100
         Y1              =   5340
         Y2              =   4860
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   25
         X1              =   2400
         X2              =   1920
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Steam_line_875 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   0
         X1              =   1920
         X2              =   1920
         Y1              =   4860
         Y2              =   3660
      End
      Begin VB.Label Label19 
         Caption         =   "875 # / 35 #"
         Height          =   195
         Left            =   11160
         TabIndex        =   68
         Top             =   120
         Width           =   1035
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   23
         X1              =   12420
         X2              =   11820
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   21
         X1              =   10920
         X2              =   10920
         Y1              =   1440
         Y2              =   1020
      End
      Begin VB.Label Label15 
         Caption         =   "To ADT"
         Height          =   255
         Left            =   5760
         TabIndex        =   63
         Top             =   8760
         Width           =   675
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   5640
         Picture         =   "Feed_Water.frx":36B4
         Stretch         =   -1  'True
         Top             =   8580
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   24
         X1              =   5460
         X2              =   5940
         Y1              =   8700
         Y2              =   8700
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   5280
         Top             =   8580
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   23
         X1              =   5340
         X2              =   5340
         Y1              =   8760
         Y2              =   8460
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   22
         X1              =   5340
         X2              =   5340
         Y1              =   8280
         Y2              =   8160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   11
         X1              =   11400
         X2              =   12060
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   11640
         X2              =   12120
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Image Image17 
         Height          =   255
         Left            =   6060
         Picture         =   "Feed_Water.frx":39BE
         Stretch         =   -1  'True
         Top             =   4380
         Width           =   255
      End
      Begin VB.Image Image12 
         Height          =   255
         Left            =   6060
         Picture         =   "Feed_Water.frx":3CC8
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   7
         X1              =   5040
         X2              =   6420
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label13 
         Caption         =   "To Main Air Ejectors"
         Height          =   315
         Left            =   6480
         TabIndex        =   56
         Top             =   3900
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "To Gland Seal Regulator"
         Height          =   435
         Left            =   3240
         TabIndex        =   55
         Top             =   6240
         Width           =   1755
      End
      Begin VB.Image Image19 
         Height          =   255
         Left            =   3360
         Picture         =   "Feed_Water.frx":3FD2
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   255
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   11
         X1              =   3420
         X2              =   5040
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   8400
         Picture         =   "Feed_Water.frx":42DC
         Stretch         =   -1  'True
         Top             =   3180
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   16
         X1              =   8040
         X2              =   8040
         Y1              =   6360
         Y2              =   5880
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   6
         X1              =   5760
         X2              =   8040
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   14
         X1              =   8100
         X2              =   8100
         Y1              =   5460
         Y2              =   4860
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   12
         X1              =   5760
         X2              =   5760
         Y1              =   5460
         Y2              =   4860
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   10
         X1              =   4560
         X2              =   3120
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   4
         X1              =   3120
         X2              =   3120
         Y1              =   4140
         Y2              =   3660
      End
      Begin VB.Image Image13 
         Height          =   255
         Left            =   9600
         Picture         =   "Feed_Water.frx":45E6
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Feed pump reciculation to Deaerator"
         Height          =   675
         Left            =   9300
         TabIndex        =   37
         Top             =   5640
         Width           =   1035
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   2880
         Picture         =   "Feed_Water.frx":48F0
         Stretch         =   -1  'True
         Top             =   7020
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   6780
         X2              =   5760
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   9
         X1              =   2760
         X2              =   4920
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Label Label11 
         Caption         =   "Main Feed"
         Height          =   255
         Left            =   3060
         TabIndex        =   35
         Top             =   6840
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Aux  Feed"
         Height          =   195
         Left            =   3060
         TabIndex        =   34
         Top             =   8040
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4950
         Top             =   7260
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   8
         X1              =   5160
         X2              =   5760
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   50
         X1              =   4920
         X2              =   4920
         Y1              =   7140
         Y2              =   7800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   51
         X1              =   5160
         X2              =   5160
         Y1              =   7140
         Y2              =   7800
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "H.P. Bleed"
         Height          =   315
         Left            =   3900
         TabIndex        =   33
         Top             =   1380
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   6
         X1              =   5760
         X2              =   5760
         Y1              =   7440
         Y2              =   6840
      End
      Begin VB.Label Label1 
         Caption         =   "875 psi"
         Height          =   255
         Left            =   5160
         TabIndex        =   31
         Top             =   5640
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "De-Aerator Feed Tank"
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
         Height          =   435
         Left            =   10620
         TabIndex        =   28
         Top             =   3000
         Width           =   975
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1215
         Index           =   2
         Left            =   8340
         Top             =   6360
         Width           =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   10
         X1              =   8520
         X2              =   11100
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Image Image16 
         Height          =   255
         Left            =   11100
         Picture         =   "Feed_Water.frx":4BFA
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   5
         X1              =   9420
         X2              =   11100
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   11100
         X2              =   11100
         Y1              =   5520
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   4
         X1              =   10920
         X2              =   11700
         Y1              =   2295
         Y2              =   2280
      End
      Begin VB.Image Image11 
         Height          =   255
         Index           =   0
         Left            =   10560
         Picture         =   "Feed_Water.frx":4F04
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image Image15 
         Height          =   360
         Left            =   11100
         Picture         =   "Feed_Water.frx":520E
         Stretch         =   -1  'True
         Top             =   1755
         Width           =   285
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   2
         X1              =   10680
         X2              =   10680
         Y1              =   2520
         Y2              =   1020
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   7
         X1              =   8520
         X2              =   8520
         Y1              =   2100
         Y2              =   5940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   43
         X1              =   11100
         X2              =   12000
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Cascade"
         Height          =   195
         Left            =   5580
         TabIndex        =   27
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "875 # / 150 #"
         Height          =   315
         Left            =   3240
         TabIndex        =   26
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "IP Bleed"
         Height          =   195
         Left            =   6960
         TabIndex        =   25
         Top             =   600
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   61
         X1              =   11100
         X2              =   11100
         Y1              =   8580
         Y2              =   5760
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   2880
         Picture         =   "Feed_Water.frx":5518
         Stretch         =   -1  'True
         Top             =   8220
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "No 2"
         Height          =   255
         Left            =   8520
         TabIndex        =   12
         Top             =   7860
         Width           =   555
      End
      Begin VB.Label Label16 
         Caption         =   "Turbo Feed Pumps"
         Height          =   735
         Index           =   0
         Left            =   8640
         TabIndex        =   10
         Top             =   3900
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   5280
         Picture         =   "Feed_Water.frx":5822
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   255
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   0
         X1              =   5760
         X2              =   5760
         Y1              =   6360
         Y2              =   5460
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   21
         X1              =   9720
         X2              =   9720
         Y1              =   5640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   20
         X1              =   8820
         X2              =   8820
         Y1              =   6840
         Y2              =   6480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   18
         X1              =   5760
         X2              =   8160
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   16
         X1              =   10140
         X2              =   11100
         Y1              =   8580
         Y2              =   8580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   14
         X1              =   2760
         X2              =   8940
         Y1              =   8340
         Y2              =   8340
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   13
         X1              =   5760
         X2              =   9060
         Y1              =   6840
         Y2              =   6840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   12
         X1              =   8040
         X2              =   8040
         Y1              =   6720
         Y2              =   6480
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   1
         X1              =   5220
         X2              =   5760
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   9720
         X2              =   10860
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label2 
         Caption         =   "Electric Motor"
         Height          =   255
         Left            =   10320
         TabIndex        =   5
         Top             =   8880
         Width           =   1095
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00A0A0A0&
         BorderWidth     =   3
         Index           =   13
         X1              =   5760
         X2              =   6780
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Shape Shape29 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   315
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   7620
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   4800
         Top             =   6900
         Width           =   555
      End
      Begin VB.Line SS_line 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Index           =   9
         X1              =   5040
         X2              =   6420
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Label Label14 
         Caption         =   "To Distil Air Ejector"
         Height          =   255
         Left            =   6480
         TabIndex        =   57
         Top             =   4320
         Width           =   1575
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   615
         Left            =   8100
         OleObjectBlob   =   "Feed_Water.frx":5B2C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisFdPmp.vsd"
         TabIndex        =   105
         Top             =   5880
         Width           =   615
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   615
         Left            =   8880
         OleObjectBlob   =   "Feed_Water.frx":D544
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisFdPmp.vsd"
         TabIndex        =   104
         Top             =   4980
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   10380
         Top             =   3480
         Width           =   315
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   10560
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   5
         Left            =   10560
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   10740
         Shape           =   3  'Circle
         Top             =   2115
         Width           =   735
      End
      Begin VB.Line Steam_line 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Index           =   37
         X1              =   6840
         X2              =   6840
         Y1              =   3180
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   15
         X1              =   8040
         X2              =   8040
         Y1              =   6960
         Y2              =   7440
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":14F5C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   154
         Top             =   5100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":1D174
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   153
         Top             =   5100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7440
         OleObjectBlob   =   "Feed_Water.frx":25B8C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   162
         Top             =   7020
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   7440
         OleObjectBlob   =   "Feed_Water.frx":2E3A4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   163
         Top             =   7020
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   10500
         OleObjectBlob   =   "Feed_Water.frx":361BC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   164
         Top             =   5280
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE VV_feed_recirc_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   10500
         OleObjectBlob   =   "Feed_Water.frx":3E9D4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   165
         Top             =   5640
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label VV_pos_feed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   7035
         TabIndex        =   233
         Top             =   5100
         Width           =   555
      End
      Begin VB.Label VV_pos_feed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   7020
         TabIndex        =   234
         Top             =   6060
         Width           =   555
      End
      Begin VB.Label Tripped_lbl 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Tripped"
         Height          =   255
         Index           =   2
         Left            =   8520
         TabIndex        =   147
         Top             =   5620
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label BYPASS_HP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "100 %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5460
         TabIndex        =   252
         Top             =   2400
         Width           =   555
      End
      Begin VB.Shape UN_reg 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   6
         Left            =   5640
         Top             =   2040
         Width           =   75
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
         Left            =   11460
         OleObjectBlob   =   "Feed_Water.frx":467EC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   138
         Top             =   660
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
         Left            =   11460
         OleObjectBlob   =   "Feed_Water.frx":4EE04
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   141
         Top             =   660
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
         Left            =   11460
         OleObjectBlob   =   "Feed_Water.frx":5781C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   140
         Top             =   660
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
         Index           =   3
         Left            =   11460
         OleObjectBlob   =   "Feed_Water.frx":5FA34
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   139
         Top             =   660
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
         Left            =   8520
         OleObjectBlob   =   "Feed_Water.frx":6784C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   130
         Top             =   660
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
         Left            =   8520
         OleObjectBlob   =   "Feed_Water.frx":6FE64
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   133
         Top             =   660
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
         Left            =   8520
         OleObjectBlob   =   "Feed_Water.frx":7887C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   132
         Top             =   660
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
         Index           =   3
         Left            =   8520
         OleObjectBlob   =   "Feed_Water.frx":80A94
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   131
         Top             =   660
         Visible         =   0   'False
         Width           =   315
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
         Left            =   4905
         OleObjectBlob   =   "Feed_Water.frx":888AC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   109
         Top             =   2520
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
         Left            =   4920
         OleObjectBlob   =   "Feed_Water.frx":908C4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   108
         Top             =   2520
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
         Left            =   4920
         OleObjectBlob   =   "Feed_Water.frx":990DC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   107
         Top             =   2520
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
         Index           =   3
         Left            =   4920
         OleObjectBlob   =   "Feed_Water.frx":A10F4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   106
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
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
         Left            =   6720
         OleObjectBlob   =   "Feed_Water.frx":A8F0C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   117
         Top             =   2580
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
         Left            =   6720
         OleObjectBlob   =   "Feed_Water.frx":B1724
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   120
         Top             =   2580
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
         Left            =   6720
         OleObjectBlob   =   "Feed_Water.frx":B973C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   119
         Top             =   2580
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
         Index           =   3
         Left            =   6720
         OleObjectBlob   =   "Feed_Water.frx":C1754
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   118
         Top             =   2580
         Visible         =   0   'False
         Width           =   375
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
         Left            =   3660
         OleObjectBlob   =   "Feed_Water.frx":C956C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   209
         Top             =   3300
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
         Left            =   3660
         OleObjectBlob   =   "Feed_Water.frx":D1B84
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   145
         Top             =   3300
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
         Left            =   3660
         OleObjectBlob   =   "Feed_Water.frx":DA59C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   144
         Top             =   3300
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
         Index           =   3
         Left            =   3660
         OleObjectBlob   =   "Feed_Water.frx":E27B4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   143
         Top             =   3300
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
         Index           =   0
         Left            =   2700
         OleObjectBlob   =   "Feed_Water.frx":EA7CC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   208
         Top             =   4500
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
         Left            =   2700
         OleObjectBlob   =   "Feed_Water.frx":F2DE4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   110
         Top             =   4500
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
         Left            =   2700
         OleObjectBlob   =   "Feed_Water.frx":FB7FC
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   112
         Top             =   4500
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
         Index           =   2
         Left            =   2700
         OleObjectBlob   =   "Feed_Water.frx":103814
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   111
         Top             =   4500
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":10BA2C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   146
         Top             =   5100
         Width           =   315
      End
      Begin VB.OLE Reg_pump_1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":114044
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   155
         Top             =   5100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":11BE5C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_black_mushroom.vsd"
         TabIndex        =   149
         Top             =   6000
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":124474
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green_mushroom.vsd"
         TabIndex        =   152
         Top             =   6000
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":12C28C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green66_mushroom.vsd"
         TabIndex        =   151
         Top             =   6000
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OLE Reg_pump_2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   6780
         OleObjectBlob   =   "Feed_Water.frx":1344A4
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Hor_up_green33_mushroom.vsd"
         TabIndex        =   150
         Top             =   6000
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2100
      TabIndex        =   17
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Feed_Water"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer

Private Sub AUTO_feed_pump_Click(Index As Integer)



' 18/7/00   ADDED this:

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   

If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        
    Exit Sub
    
End If


' Try to put the feed pump into AUTO....

If Index = 1 Then
  Call setPCDI(23)
ElseIf Index = 2 Then
  Call setPCDI(120)
End If


End Sub

Private Sub AUTO_VV_Click(Index As Integer)


set_dummy_focus = True


' 17/10/99
Select Case Index

  Case 1
    Call setPCDI(11)
  Case 2
    Call setPCDI(99)
    
End Select


End Sub


Private Sub CLOSE_VV_Click(Index As Integer)

' === 13/11/02 ===
set_dummy_focus = True

' 2/11/99   CLOSE
Select Case Index

  Case 1
    Call setPCDI(12)
  Case 2
    Call setPCDI(100)
    
End Select


End Sub

Private Sub Command1_Click()

Condensate.Show

End Sub

Private Sub Command2_Click(Index As Integer)

Stbd_Boiler.Show

End Sub

Private Sub Command3_Click()

' 14/2/00
'Shaft_rpm.Caption = Format(LSS_reals(80), "##0") & " rpm"
If LSS_reals(80) > 25# Then
  Gland_Steam_Full_away.Show
Else
  Gland_Steam.Show
End If


End Sub

Private Sub Command4_Click(Index As Integer)

Port_Boiler.Show

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idFeed_water)


' === 18/02/03 TESTING NO CRASH...===
'    RECIRC_label(55).Visible = False
' === 18/02/03 TESTING method or data member not found....===
'Valve(44).fred = 1
' === 18/02/03 TESTiNG Worked ?...
'Valve(44).OpenBtmRight = "Hello"
' === 18/02/03 TESTiNG Type Mis-match...
'Command1.Picture = "Hello"

' === 18/02/03 TESTiNG Worked OK ( also tested Steam_SimV22_test.exe and was OK )...
'VScroll1.Value = 255 * 4


' === 18/11/02 ===
update_VSCROLL(23) = True
update_VSCROLL(25) = True
update_VSCROLL(28) = True




' 29/4/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    
    ' === 27/09/02 14.114 ===
    RECIRC_label(1).Visible = False
    RECIRC_label(2).Visible = False
    OPEN_VV(1).Visible = False
    OPEN_VV(2).Visible = False
    CLOSE_VV(1).Visible = False
    CLOSE_VV(2).Visible = False
    AUTO_VV(1).Visible = False
    AUTO_VV(2).Visible = False
    ' =======================
    
        ' === 10/04/02 S Green e-mail 30/3/02 point 2b ===
        REMOTE_feed_control(1).Visible = False
        REMOTE_feed_control(2).Visible = False
    
    
     ' 11/12/00
     BYPASS_IP.Visible = False
     BYPASS_875_35.Visible = False
     BYPASS_HP.Visible = False
     BYPASS_atom.Visible = False
     BYPASS_875_150.Visible = False
     BYPASS_35_dump.Visible = False
     BYPASS_feed_1.Visible = False
     BYPASS_feed_2.Visible = False
   
   
    ' 11/12/00
    HScroll1.Visible = False
    
    
    ' === 18/9/00  Point 8.126 ===
    
    ' === 01/03/01 Point 10.2b NOT the Elec Pump ===
    'ONpump(0).Visible = False
    'OFFpump(0).Visible = False
    'sbyPUMP(0).Visible = False

    ONpump(3).Visible = False
    OFFpump(3).Visible = False
    sbyPUMP(3).Visible = False
    
    ONpump(4).Visible = False
    OFFpump(4).Visible = False
    sbyPUMP(4).Visible = False
    
    AUX_LO_label.Visible = False
    AUX_1.Visible = False
    AUX_2.Visible = False
    
    ' 7/9/00
    TLO_main_feed(1).Visible = False
    TLO_main_feed(2).Visible = False
    
    ' 10/12/00
    PLO_total_feed_pump(1).Visible = False
    PLO_total_feed_pump(2).Visible = False
    
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
        
        VV_pos_feed(1).Visible = False
        VV_pos_feed(2).Visible = False
        
        VV_feed_aux_pos(1).Visible = False
        VV_feed_aux_pos(2).Visible = False
        
    
Else

    ' === 27/09/02 14.114 ===
    RECIRC_label(1).Visible = True
    RECIRC_label(2).Visible = True
    OPEN_VV(1).Visible = True
    OPEN_VV(2).Visible = True
    CLOSE_VV(1).Visible = True
    CLOSE_VV(2).Visible = True
    AUTO_VV(1).Visible = True
    AUTO_VV(2).Visible = True
    ' =======================

        ' === 10/04/02 S Green e-mail 30/3/02 point 2b ===
        REMOTE_feed_control(1).Visible = True
        REMOTE_feed_control(2).Visible = True

     ' 11/12/00
     BYPASS_IP.Visible = True
     BYPASS_875_35.Visible = True
     BYPASS_HP.Visible = True
     BYPASS_atom.Visible = True
     BYPASS_875_150.Visible = True
     BYPASS_35_dump.Visible = True
     BYPASS_feed_1.Visible = True
     BYPASS_feed_2.Visible = True



    ' 11/12/00
    HScroll1.Visible = True

    ' === 18/9/00  Point 8.126 ===
    ONpump(0).Visible = True
    OFFpump(0).Visible = True
    sbyPUMP(0).Visible = True


    ONpump(3).Visible = True
    OFFpump(3).Visible = True
    sbyPUMP(3).Visible = True
    
    ONpump(4).Visible = True
    OFFpump(4).Visible = True
    sbyPUMP(4).Visible = True
    
    AUX_LO_label.Visible = True
    AUX_1.Visible = True
    AUX_2.Visible = True
    
    ' 7/9/00
    TLO_main_feed(1).Visible = True
    TLO_main_feed(2).Visible = True

    ' 10/12/00
    PLO_total_feed_pump(1).Visible = True
    PLO_total_feed_pump(2).Visible = True

    
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
        
        VV_pos_feed(1).Visible = True
        VV_pos_feed(2).Visible = True
        
        VV_feed_aux_pos(1).Visible = True
        VV_feed_aux_pos(2).Visible = True


End If


' 1/12/99
VScroll1.Visible = PTT
HScroll2.Visible = PTT
HScroll3.Visible = PTT

HScroll5.Visible = PTT
HScroll4.Visible = PTT

HScroll9.Visible = PTT

HScroll7.Visible = PTT

HScroll6.Visible = PTT

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
    
    
    Call JYAScreenResize(Me, idFeed_water)
    
    
    HavenStretch1.Visible = True
    
    
    
   
End If

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

Private Sub HScroll4_Change()
Dim AIvalue As Integer


' 3/11/99  Aux feed from Main pump # 2

' 1/12/99
'Call set_VSCROLL(20, HScroll4.Value)
AIvalue = HScroll4.Value / 4
Call setAI(13, AIvalue)


End Sub

Private Sub HScroll5_Change()
Dim AIvalue As Integer

' 3/11/99  Aux feed from Main pump # 1

' 1/12/99
'Call set_VSCROLL(19, HScroll5.Value)
AIvalue = HScroll5.Value / 4
Call setAI(12, AIvalue)


End Sub

Private Sub HScroll6_Change()
Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)

' 6/11/99  Feed_pump_bypass   # 1

' 1/12/99
'Call set_VSCROLL(26, HScroll6.Value)

AIvalue = HScroll6.Value / 4
Call setAI(16, AIvalue)



End Sub

Private Sub HScroll7_Change()
Dim AIvalue As Integer

'c     21     Eight75_150_bypass (*1000)
'c     22     Eight75_35_bypass  (*1000)
'c     23     Dump_35_bypass (*1000)
'c     24     HP_bypass (*1000)
'c     25     IP_bypass (*1000)
'c     26     Feed_pump_bypass   # 1 (*1000)
'c     27     Feed_pump_bypass   # 2 (*1000)
'c     28     Atom_bypass (*1000)

' 6/11/99  Feed_pump_bypass   # 2

' 1/12/99
'Call set_VSCROLL(27, HScroll7.Value)
AIvalue = HScroll7.Value / 4
Call setAI(15, AIvalue)


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

Private Sub HScroll9_Change()
Dim AIvalue As Integer

' 1/12/99  VALVE is now variable ( used to be clickable )

AIvalue = HScroll9.Value / 4
Call setAI(14, AIvalue)


End Sub

Private Sub LO_feed_remote_Click(Index As Integer)

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(4=OFF,nPump)
If Index = 0 Then
'  11/10/99
'   Call setPCDI(PumpIO_LSS(4, 3))
    Call setPCDI(PumpIO_LSS(4, 9))
    
ElseIf Index = 3 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(4, 6))       ' Feed pump AUX LO # 1
        
ElseIf Index = 4 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(4, 18))      ' Feed pump AUX LO  # 2

End If

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)
If Index = 0 Then
    '  11/10/99
    '   Call setPCDI(PumpIO_LSS(2, 3))
        Call setPCDI(PumpIO_LSS(2, 9))      ' Feed Water Elec pump
    
ElseIf Index = 3 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(2, 6))       ' Feed pump AUX LO # 1
        
ElseIf Index = 4 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(2, 18))      ' Feed pump AUX LO  # 2

End If

End Sub

Private Sub OPEN_VV_Click(Index As Integer)

set_dummy_focus = True


' 2/11/99   OPEN
Select Case Index

  Case 1
    Call setPCDI(10)
  Case 2
    Call setPCDI(98)
    
End Select


End Sub

Private Sub P_MF_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_LO = 75
Alarm_set_point.Show

End Sub

Private Sub PLO_total_feed_pump_Click(Index As Integer)

' 27/8/00
Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_LO = 78
Else
    nAlarm_LO = 79
End If

Alarm_set_point.Show

End Sub

Private Sub REMOTE_feed_control_Click(Index As Integer)

set_dummy_focus = True

' 17/8/00  If on the REAL sim then this is a HARD switch in the ER

' 17/8/00 If PTT then do the following...
If PTT Then
    If Index = 1 Then
        Call setPCDI(356)
    Else
        Call setPCDI(357)
    End If
End If

End Sub

Private Sub sbyPUMP_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO(5=Stand-bye,nPump)
If Index = 0 Then
'  11/10/99
'   Call setPCDI(PumpIO_LSS(5, 3))
    Call setPCDI(PumpIO_LSS(5, 9))
    
ElseIf Index = 3 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(5, 6))       ' Feed pump AUX LO # 1
        
ElseIf Index = 4 Then

        ' 2/5/00
        Call setPCDI(PumpIO_LSS(5, 18))      ' Feed pump AUX LO  # 2

End If

End Sub


Private Sub SP_875_150_Click()

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 257

Alarm_set_point.Show

End Sub

Private Sub SP_875_35_Click()

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 259

Alarm_set_point.Show

End Sub

Private Sub SP_HP_Click()


'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 256

Alarm_set_point.Show

End Sub

Private Sub SP_IP_Click()

'  4/5/00
Unload Alarm_set_point   ' Do this first

nAlarm_LO = 258

Alarm_set_point.Show

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
     Call update_the_slider(VScroll1, VB_AI(6) * 4)
     
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
If update_VBAI(12) Then
     update_VBAI(12) = False
     
     'HScroll5.Value = VB_AI(12) * 4
     Call update_the_slider(HScroll5, VB_AI(12) * 4)
     
End If
If update_VBAI(13) Then
     update_VBAI(13) = False
     
     'HScroll4.Value = VB_AI(13) * 4
     Call update_the_slider(HScroll4, VB_AI(13) * 4)
     
End If
If update_VBAI(14) Then
     update_VBAI(14) = False
     
     'HScroll9.Value = VB_AI(14) * 4
     Call update_the_slider(HScroll9, VB_AI(14) * 4)
     
End If
If update_VBAI(15) Then
     update_VBAI(15) = False
     
     'HScroll7.Value = VB_AI(15) * 4
     Call update_the_slider(HScroll7, VB_AI(15) * 4)
     
End If
If update_VBAI(16) Then
     update_VBAI(16) = False
     
     ' === 19/02/03 ===
     'HScroll6.Value = VB_AI(16) * 4
     Call update_the_slider(HScroll6, VB_AI(16) * 4)
     ' TESTING...
     'i = HScroll6.Value
  
End If
' === 13/11/02 END ===


' 17/8/00
If VB_steam_logicals(123) Then
    REMOTE_feed_control(1).BackColor = &HFFFF&          ' Yellow
Else
    REMOTE_feed_control(1).BackColor = &HC0C0C0         ' Grey
End If
If VB_steam_logicals(124) Then
    REMOTE_feed_control(2).BackColor = &HFFFF&          ' Yellow
Else
    REMOTE_feed_control(2).BackColor = &HC0C0C0         ' Grey
End If


' 18/9?00 Point 8.126
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    Tripped_lbl(1).Visible = False
    Tripped_lbl(2).Visible = False
Else

  If VB_steam_logicals(125) Then
    Tripped_lbl(1).Visible = True
  Else
    Tripped_lbl(1).Visible = False
  End If
  If VB_steam_logicals(126) Then
    Tripped_lbl(2).Visible = True
  Else
    Tripped_lbl(2).Visible = False
  End If

End If



' 15/8/00
P_main_out(1).text = Format(LSS_reals(344) * Press_mult, "##0.0") & Press_chars

' 5/9/00
P_main_out(2).text = Format(LSS_reals(345) * Press_mult, "##0.0") & Press_chars


' 19/7/00....

If LSS_LOGICALS(61) Then
    AUTO_feed_pump(1).BackColor = &HFFFF&         '  YELLOW
Else
    AUTO_feed_pump(1).BackColor = &HC0C0C0     'GREY
End If

If LSS_LOGICALS(62) Then
    AUTO_feed_pump(2).BackColor = &HFFFF&         '  YELLOW
Else
    AUTO_feed_pump(2).BackColor = &HC0C0C0     'GREY
End If






'   2/5/00
' LO pump for Feed pump # 1
Call update_pump_LSS(Me, 3, 6)      ' Feed pump AUX LO  # 1
' LO pump for Feed pump # 1
Call update_pump_LSS(Me, 4, 18)     ' Feed pump AUX LO  # 2


' 27/8/00
Call paint_alarm(Me.PLO_total_feed_pump(1), 78, 0, 0, 0)
Call paint_alarm(Me.PLO_total_feed_pump(2), 79, 0, 0, 0)


' 7/1/00
Call paint_alarm(Me.P_MF, 75, 0, 0, 0)
Call paint_alarm(Me.Z_DFT, 76, 81, 0, 0)

' 10/1/00
'    PC_REALS(398) = F_feed_pump(1)
'    PC_REALS(399) = F_feed_pump(2)
If (LSS_reals(398) > 0.02) Then
  HavenPump1(1).PumpOn = True
Else
  HavenPump1(1).PumpOn = False
End If
If (LSS_reals(399) > 0.02) Then
  HavenPump1(2).PumpOn = True
Else
  HavenPump1(2).PumpOn = False
End If


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

T_DFT.text = Format(LSS_reals(309) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_HP_htr.text = Format(LSS_reals(310) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
    
    
'          PC_REALS(155) = TLO_main_feed(1)
'          PC_REALS(156) = TLO_main_feed(2)
' 7/9/00
TLO_main_feed(1).text = Format(LSS_reals(155) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Call paint_alarm(Me.TLO_main_feed(1), 83, 0, 0, 0)
TLO_main_feed(2).text = Format(LSS_reals(156) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
Call paint_alarm(Me.TLO_main_feed(2), 84, 0, 0, 0)
    
' 12/9/00
Tsteam11.text = Format(LSS_reals(71) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
    

    P_150.Visible = Advanced_mode
    P_875_desup_OUT.Visible = Advanced_mode
    P_HP_reg_IN.Visible = Advanced_mode
    P_IP_reg_IN.Visible = Advanced_mode
    P_35.Visible = Advanced_mode
    P_Atom.Visible = Advanced_mode
    
   ' 16/8/00
   
   ' 1/10/00
    P_main(1).Visible = Advanced_mode
    P_main(2).Visible = Advanced_mode
   ' P_main(1).Visible = False
   ' P_main(2).Visible = False
    
    ' 15/8/00
   ' F_main(1).Visible = Advanced_mode
   ' F_main(2).Visible = Advanced_mode
   ' F_MF(1).Visible = Advanced_mode
   ' F_MF(2).Visible = Advanced_mode
    F_main(1).Visible = False
    F_main(2).Visible = False
    F_MF(1).Visible = False
    F_MF(2).Visible = False
    
    F_MF_recirc(1).Visible = Advanced_mode
    F_MF_recirc(2).Visible = Advanced_mode
    
    F_MF_aux(1).Visible = Advanced_mode
    F_MF_aux(2).Visible = Advanced_mode
    
    F_main_feed(1).Visible = Advanced_mode
    F_main_feed(2).Visible = Advanced_mode
    
    F_aux_feed(1).Visible = Advanced_mode
    F_aux_feed(2).Visible = Advanced_mode
    

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


' 5/6/00
VV_pos_feed(1).Caption = Format(LSS_reals(374), "##0") & " %"
VV_pos_feed(2).Caption = Format(LSS_reals(376), "##0") & " %"


'  === 8/11/99 Regulator  Valves ===
Call update_regulator(364, Reg_875_150(0), Reg_875_150(1), Reg_875_150(2), Reg_875_150(3))
Call update_regulator(366, Reg_875_35(0), Reg_875_35(1), Reg_875_35(2), Reg_875_35(3))
Call update_regulator(368, Dump_35(0), Dump_35(1), Dump_35(2), Dump_35(3))
Call update_regulator(370, Reg_HP(0), Reg_HP(1), Reg_HP(2), Reg_HP(3))
Call update_regulator(372, Reg_IP(0), Reg_IP(1), Reg_IP(2), Reg_IP(3))
Call update_regulator(374, Reg_pump_1(0), Reg_pump_1(1), Reg_pump_1(2), Reg_pump_1(3))
Call update_regulator(376, Reg_pump_2(0), Reg_pump_2(1), Reg_pump_2(2), Reg_pump_2(3))
Call update_regulator(378, Reg_Atomising(0), Reg_Atomising(1), Reg_Atomising(2), Reg_Atomising(3))


' 26/8/00
'       PC_REALS(117) = VV_main_feed_pos(2) * 100#
Call update_green_VV(117, valve(5))


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
Call update_green_VV(375, valve(32))
' 11/12/00
BYPASS_feed_1.Caption = Format(LSS_reals(375), "##0") & " %"

'        PC_REALS(377) = Feed_pump_bypass(2) * 100#
Call update_green_VV(377, valve(33))
' 11/12/00
BYPASS_feed_2.Caption = Format(LSS_reals(377), "##0") & " %"



'        PC_REALS(379) = Atom_bypass * 100#
Call update_green_VV(379, valve(47))
' 11/12/00
BYPASS_atom.Caption = Format(LSS_reals(379), "##0") & " %"


' Main Feed #1 to AUX feed valve...
VV_feed_aux_pos(1).Caption = Format(LSS_reals(356), "##0") & " %"

Call update_green_VV(356, valve(12))

'If LSS_reals(356) > 99# Then
'    green_colour = 255
'ElseIf LSS_reals(356) < 1# Then
'    green_colour = 0
'Else
'    green_colour = LSS_reals(356) * 2.55
'End If
'Valve(12).ColourClosed = RGB(0, green_colour, 0)


' Main Feed #2 to AUX feed valve...
VV_feed_aux_pos(2).Caption = Format(LSS_reals(357), "##0") & " %"

Call update_green_VV(357, valve(51))

'If LSS_reals(357) > 99# Then
'    green_colour = 255
'ElseIf LSS_reals(357) < 1# Then
'    green_colour = 0
'Else
'    green_colour = LSS_reals(357) * 2.55
'End If
'Valve(51).ColourClosed = RGB(0, green_colour, 0)



If SI_units Then
    
    ' 17/8/00
    PLO_total_feed_pump(1).text = Format(LSS_reals(486) * Press_mult, "##0.0") & Press_chars
    PLO_total_feed_pump(2).text = Format(LSS_reals(487) * Press_mult, "##0.0") & Press_chars

    Z_DFT.text = Format(LSS_reals(451) * 25.4, "###0.0") & " mm"

    P_aux_feed_out.text = Format(LSS_reals(362) * Press_mult, "##0.0") & Press_chars
  
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0.0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0.0") & Press_chars
    
    P_MF.text = Format(LSS_reals(342) * Press_mult, "##0.0") & Press_chars
    P_AF.text = Format(LSS_reals(343) * Press_mult, "##0.0") & Press_chars

  If (Advanced_mode) Then
  
    P_150.text = Format(LSS_reals(380) * Press_mult, "##0.0") & Press_chars
    P_875_desup_OUT.text = Format(LSS_reals(381) * Press_mult, "##0.0") & Press_chars
    P_HP_reg_IN.text = Format(LSS_reals(382) * Press_mult, "##0.0") & Press_chars
    P_IP_reg_IN.text = Format(LSS_reals(383) * Press_mult, "##0.0") & Press_chars
    P_35.text = Format(LSS_reals(384) * Press_mult, "##0.0") & Press_chars
    
    P_Atom.text = Format(LSS_reals(385) * Press_mult, "##0.0") & Press_chars
    
    P_main(1).text = Format(LSS_reals(346) * Press_mult, "##0.0") & Press_chars
    P_main(2).text = Format(LSS_reals(347) * Press_mult, "##0.0") & Press_chars
    
    F_main(1).text = Format(LSS_reals(348) * Flow_mult, "###0.0") & Flow_chars
    F_main(2).text = Format(LSS_reals(349) * Flow_mult, "###0.0") & Flow_chars
    
    F_MF(1).text = Format(LSS_reals(350) * Flow_mult, "###0.0") & Flow_chars
    F_MF(2).text = Format(LSS_reals(351) * Flow_mult, "###0.0") & Flow_chars
    
    F_MF_recirc(1).text = Format(LSS_reals(352) * Flow_mult, "###0.0") & Flow_chars
    F_MF_recirc(2).text = Format(LSS_reals(353) * Flow_mult, "###0.0") & Flow_chars
    
    F_MF_aux(1).text = Format(LSS_reals(354) * Flow_mult, "###0.0") & Flow_chars
    F_MF_aux(2).text = Format(LSS_reals(355) * Flow_mult, "###0.0") & Flow_chars
    
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult, "###0.0") & Flow_chars
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult, "###0.0") & Flow_chars
    
    F_aux_feed(1).text = Format(LSS_reals(360) * Flow_mult, "###0.0") & Flow_chars
    F_aux_feed(2).text = Format(LSS_reals(361) * Flow_mult, "###0.0") & Flow_chars
  End If
  
Else

    ' 17/8/00
    PLO_total_feed_pump(1).text = Format(LSS_reals(486) * Press_mult, "##0") & Press_chars
    PLO_total_feed_pump(2).text = Format(LSS_reals(487) * Press_mult, "##0") & Press_chars


    Z_DFT.text = Format(LSS_reals(451), "###0.0") & " in"
    
    P_aux_feed_out.text = Format(LSS_reals(362) * Press_mult, "##0") & Press_chars
    
    Psteam11.text = Format(LSS_reals(61) * Press_mult, "##0") & Press_chars
    Psteam12.text = Format(LSS_reals(62) * Press_mult, "##0") & Press_chars
    
    P_MF.text = Format(LSS_reals(342) * Press_mult, "##0") & Press_chars
    P_AF.text = Format(LSS_reals(343) * Press_mult, "##0") & Press_chars
    
  If (Advanced_mode) Then
  
    P_150.text = Format(LSS_reals(380) * Press_mult, "##0") & Press_chars
    P_875_desup_OUT.text = Format(LSS_reals(381) * Press_mult, "##0") & Press_chars
    P_HP_reg_IN.text = Format(LSS_reals(382) * Press_mult, "##0") & Press_chars
    P_IP_reg_IN.text = Format(LSS_reals(383) * Press_mult, "##0") & Press_chars
    P_35.text = Format(LSS_reals(384) * Press_mult, "##0") & Press_chars
    P_Atom.text = Format(LSS_reals(385) * Press_mult, "##0") & Press_chars
    
    P_main(1).text = Format(LSS_reals(346) * Press_mult, "##0") & Press_chars
    P_main(2).text = Format(LSS_reals(347) * Press_mult, "##0") & Press_chars
    
    F_main(1).text = Format(LSS_reals(348) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_main(2).text = Format(LSS_reals(349) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MF(1).text = Format(LSS_reals(350) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_MF(2).text = Format(LSS_reals(351) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MF_recirc(1).text = Format(LSS_reals(352) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_MF_recirc(2).text = Format(LSS_reals(353) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_MF_aux(1).text = Format(LSS_reals(354) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_MF_aux(2).text = Format(LSS_reals(355) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_main_feed(1).text = Format(LSS_reals(358) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_main_feed(2).text = Format(LSS_reals(359) * Flow_mult / 3600#, "###0.0") & " lb/s"
    
    F_aux_feed(1).text = Format(LSS_reals(360) * Flow_mult / 3600#, "###0.0") & " lb/s"
    F_aux_feed(2).text = Format(LSS_reals(361) * Flow_mult / 3600#, "###0.0") & " lb/s"
  End If
  
End If

'  DFT Tank  ( -35 inches to +35 inches )
DFT_level = (LSS_reals(451) + 35#) * 100# / 70#
If DFT_level > 99 Then
  valve(44).LineWidth = 100
ElseIf DFT_level < 1 Then
  valve(44).LineWidth = 0
Else
  valve(44).LineWidth = DFT_level
End If


' === 21/11/00  Feed Pump re-circ valves===
   Call update_vvalve_LSS(Me, 85, 85, -1)
   Call update_vvalve_LSS(Me, 86, 86, -1)

' OPEN re-circ valves
' ==== 24/11/00 ====
'If (D_OUTPUT(145)) Then
If vv_status_LSS(85) = 3 Then


  OPEN_VV(1).BackColor = &HFF00&            ' GREEN
  
  ' 5/5/00
  'VV_feed_recirc_1(0).Visible = False
  'VV_feed_recirc_1(3).Visible = True
    valve(54).OpenBtmRight = True
    valve(54).OpenTopLeft = True
    
Else
  OPEN_VV(1).BackColor = &HC0C0C0           ' GREY
End If


' ==== 24/11/00 ====
'If (D_OUTPUT(493)) Then
If vv_status_LSS(86) = 3 Then

  OPEN_VV(2).BackColor = &HFF00&            ' GREEN
  
  ' 5/5/00
 ' VV_feed_recirc_2(0).Visible = False
 ' VV_feed_recirc_2(3).Visible = True
    valve(55).OpenBtmRight = True
    valve(55).OpenTopLeft = True
 
Else
  OPEN_VV(2).BackColor = &HC0C0C0           ' GREY
End If

' CLOSE re-circ valves
' ==== 24/11/00 ====
'If (D_OUTPUT(149)) Then
If vv_status_LSS(85) = 0 Then

  CLOSE_VV(1).BackColor = &H8080FF           ' PINK
  
  ' 5/5/00
 ' VV_feed_recirc_1(0).Visible = True
 ' VV_feed_recirc_1(3).Visible = False
    valve(54).OpenBtmRight = False
    valve(54).OpenTopLeft = False
 
Else
  CLOSE_VV(1).BackColor = &HC0C0C0           ' GREY
End If


' ==== 24/11/00 ====
'If (D_OUTPUT(497)) Then
If vv_status_LSS(86) = 0 Then

  CLOSE_VV(2).BackColor = &H8080FF           ' PINK
  
  ' 5/5/00
 ' VV_feed_recirc_2(0).Visible = True
 ' VV_feed_recirc_2(3).Visible = False
    valve(55).OpenBtmRight = False
    valve(55).OpenTopLeft = False

Else
  CLOSE_VV(2).BackColor = &HC0C0C0           ' GREY
End If


' Auto re-circ valves
If (D_OUTPUT(147)) Then
  AUTO_VV(1).BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_VV(1).BackColor = &HC0C0C0           ' GREY
End If
If (D_OUTPUT(495)) Then
  AUTO_VV(2).BackColor = &HFFFF&            ' YELLOW
Else
  AUTO_VV(2).BackColor = &HC0C0C0           ' GREY
End If


' Note: update_pump_LSS(syxxx,index,nPump)
Call update_pump_LSS(Me, 0, 9)

' update_vvalve_LSS(syxxx,nValve,index1,index2)


' 26/8/00  Valve(5) is a POT

'For i = 0 To 7
'   Call update_vvalve_LSS(Me, 70 + i, i, -1)
'Next i
For i = 0 To 4
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i
For i = 6 To 7
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i


For i = 9 To 11
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i
For i = 13 To 19
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i

   Call update_vvalve_LSS(Me, 70 + 22, 22, -1)
   
For i = 24 To 27
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i
For i = 29 To 31
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i
For i = 34 To 38
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i
For i = 41 To 42
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i

'For i = 45 To 47
For i = 45 To 46
   Call update_vvalve_LSS(Me, 70 + i, i, -1)
Next i
   
'  6/9/00   Point Hotel Services Valve
   Call update_vvalve_LSS(Me, 126, 56, -1)




End Sub

Private Sub TLO_main_feed_Click(Index As Integer)


' 7/9/00
Unload Alarm_set_point   ' Do this first

If Index = 1 Then
    nAlarm_HI = 83
Else
    nAlarm_HI = 84
End If

Alarm_set_point.Show

End Sub

Private Sub Tripped_lbl_Click(Index As Integer)
Dim nlog As Integer


' 17/8/00
If PTT Then
    nlog = 124 + Index
    status_flag = False
    
   ' === 22/11/02 ===
   'Call set_STEAM_logicals(nlog, status_flag)
    Call setSTEAMlogicals(nlog, status_flag)
    
End If

End Sub

Private Sub valve_Click(Index As Integer)

   Control_PC_check
   If This_Control_status = False Then Exit Sub
   
  
 ' 7/5/00
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
        
    '    IP            Cascade         HP
    
   ' === 10/06/02 Point 14.502 ===
   'If Index = 4 Or Index = 9 Or Index = 24 Then
   
   ' === 23/09/02 Point 14.502 MUST BE CLICKABLE FROM ALL STATIONS ===
   ' If Index = 4 Or Index = 9 Or Index = 24 Or Index = 27 Or Index = 29 Then
    If Index = 4 Or Index = 9 Or Index = 24 Then
    
        Exit Sub
        
    End If
    
End If
  
  
' 1/12/99
If Not PTT Then    ' NON-Clickable Valves...


   ' === 21/03/02 Point 14.502 ===
   
   ' === 10/06/02 PUT BACK ORIGINAL LINE... ===
    If Index = 13 Or Index = 14 Or Index = 30 Or Index = 31 Then
   'If Index = 13 Or Index = 14 Or Index = 30 Or Index = 31 Or Index = 27 Or Index = 29 Then
    
        Exit Sub
    End If
End If


' 26/8/00  Valve is a pot
If Index = 5 Then
        Exit Sub
End If

    Call valve_pcdi_LSS(Index + 70)


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

Private Sub VScroll2_Change()


End Sub

Private Sub VScroll3_Change()

' 6/11/99  P_875_desup_OUT
Call set_VSCROLL(30, VScroll3.Value)

End Sub

Private Sub VScroll4_Change()

' 6/11/99  P_HP_reg_IN
Call set_VSCROLL(31, VScroll4.Value)


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

Private Sub VScroll6_Change()

' 8/11/99  P_iP_reg_IN
Call set_VSCROLL(32, VScroll6.Value)

End Sub

Private Sub Z_DFT_Click()

' 7/1/00
Unload Alarm_set_point   ' Do this first
nAlarm_HI = 81
nAlarm_LO = 76
Alarm_set_point.Show

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
