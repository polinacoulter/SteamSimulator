VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.1#0"; "havenstretch.ocx"
Object = "{792D1B7B-9CB4-11D1-85C9-444553540000}#2.1#0"; "HavenPump v1_42.ocx"
Object = "{ECE244BE-884F-11D1-85C9-444553540000}#1.0#0"; "HavenValve V1_2.ocx"
Begin VB.Form Forced_Draft 
   Caption         =   "Combustion Air"
   ClientHeight    =   12570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   Icon            =   "Forced_Draft.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Forced_Draft.frx":030A
   ScaleHeight     =   12570
   ScaleWidth      =   13785
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   12255
      Left            =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   21616
      Begin VB.TextBox T_Flue_gas 
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
         Left            =   9720
         TabIndex        =   212
         Text            =   "77 F"
         Top             =   360
         Width           =   675
      End
      Begin VB.TextBox T_Flue_gas 
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
         Left            =   2160
         TabIndex        =   211
         Text            =   "77 F"
         Top             =   360
         Width           =   675
      End
      Begin VB.TextBox T_Flue_gas_ECON_out 
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
         Left            =   2400
         TabIndex        =   210
         Text            =   "77 F"
         Top             =   5400
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox T_Flue_gas_ECON_in 
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
         Left            =   2160
         TabIndex        =   209
         Text            =   "77 F"
         Top             =   7440
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F"
         Height          =   315
         Left            =   11580
         TabIndex        =   208
         Top             =   4200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "T"
         Height          =   315
         Left            =   11580
         TabIndex        =   207
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F"
         Height          =   315
         Left            =   360
         TabIndex        =   206
         Top             =   4140
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "T"
         Height          =   315
         Left            =   360
         TabIndex        =   205
         Top             =   3780
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Oxygen 
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
         Height          =   255
         Index           =   2
         Left            =   9360
         TabIndex        =   200
         Text            =   "xx %"
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox Oxygen 
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
         Height          =   255
         Index           =   1
         Left            =   2100
         TabIndex        =   199
         Text            =   "xx %"
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox Q_burn 
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
         Height          =   255
         Index           =   4
         Left            =   10440
         TabIndex        =   198
         Text            =   "xx kW"
         Top             =   7860
         Width           =   1035
      End
      Begin VB.TextBox Q_burn 
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
         Height          =   255
         Index           =   3
         Left            =   10500
         TabIndex        =   197
         Text            =   "xx kW"
         Top             =   10380
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FO Service"
         Height          =   495
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   190
         Top             =   4020
         Width           =   1095
      End
      Begin VB.CommandButton Burner2_flame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 2"
         Height          =   315
         Index           =   2
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   9240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Burner1_flame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1"
         Height          =   315
         Index           =   2
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   10680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox Q_burn 
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
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   187
         Text            =   "xx kW"
         Top             =   7800
         Width           =   1035
      End
      Begin VB.TextBox Q_burn 
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
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   186
         Text            =   "xx kW"
         Top             =   10440
         Width           =   1035
      End
      Begin VB.TextBox B_Smoke 
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
         Height          =   255
         Index           =   4
         Left            =   11400
         TabIndex        =   181
         Text            =   "xx %"
         Top             =   8280
         Width           =   1035
      End
      Begin VB.TextBox B_Smoke 
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
         Height          =   255
         Index           =   3
         Left            =   11340
         TabIndex        =   180
         Text            =   "xx %"
         Top             =   9780
         Width           =   1035
      End
      Begin VB.TextBox B_Smoke 
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   179
         Text            =   "xx %"
         Top             =   8220
         Width           =   1035
      End
      Begin VB.TextBox B_Smoke 
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
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   178
         Text            =   "xx %"
         Top             =   9720
         Width           =   1035
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 2"
         Height          =   315
         Index           =   4
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   11520
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1"
         Height          =   315
         Index           =   3
         Left            =   8940
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   11520
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Burner2_flame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 2"
         Height          =   315
         Index           =   1
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   9240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Burner_ON_OFF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1"
         Height          =   315
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   2040
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton Burner1_flame 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Burner 1"
         Height          =   315
         Index           =   1
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   10680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox FFO_burner 
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
         Height          =   255
         Index           =   4
         Left            =   8700
         TabIndex        =   172
         Text            =   "2.1 kg/s"
         Top             =   8280
         Width           =   1035
      End
      Begin VB.TextBox FFO_burner 
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
         Height          =   255
         Index           =   3
         Left            =   8700
         TabIndex        =   171
         Text            =   "2.1 kg/s"
         Top             =   9840
         Width           =   1035
      End
      Begin VB.TextBox FFO_burner 
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
         Height          =   255
         Index           =   2
         Left            =   2820
         TabIndex        =   170
         Text            =   "2.1 kg/s"
         Top             =   8340
         Width           =   1035
      End
      Begin VB.TextBox FFO_burner 
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
         Index           =   0
         Left            =   7440
         TabIndex        =   169
         Text            =   "2.1 kg/s"
         Top             =   11580
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox FFO_burner 
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
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   168
         Text            =   "2.1 kg/s"
         Top             =   9780
         Width           =   1035
      End
      Begin VB.TextBox F_air_burner 
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
         Index           =   4
         Left            =   9360
         TabIndex        =   167
         Text            =   "xx kg/s"
         Top             =   8640
         Width           =   915
      End
      Begin VB.TextBox F_air_burner 
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
         Index           =   3
         Left            =   9360
         TabIndex        =   166
         Text            =   "xx kg/s"
         Top             =   10140
         Width           =   915
      End
      Begin VB.TextBox F_air_burner 
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
         TabIndex        =   165
         Text            =   "xx kg/s"
         Top             =   8640
         Width           =   915
      End
      Begin VB.TextBox F_air_burner 
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
         Left            =   2340
         TabIndex        =   164
         Text            =   "xx kg/s"
         Top             =   10140
         Width           =   915
      End
      Begin VB.TextBox PFD_uptakes 
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
         Height          =   375
         Index           =   2
         Left            =   9480
         TabIndex        =   163
         Text            =   "8.5 ""wat"
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox PFD_uptakes 
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
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   162
         Text            =   "8.5 ""wat"
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox PFD_econ_in 
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
         Height          =   375
         Index           =   2
         Left            =   9540
         TabIndex        =   161
         Text            =   "8.5 ""wat"
         Top             =   7080
         Width           =   855
      End
      Begin VB.TextBox PFD_econ_in 
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
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   160
         Text            =   "8.5 ""wat"
         Top             =   7020
         Width           =   855
      End
      Begin VB.TextBox PFD_furnace 
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
         Height          =   375
         Index           =   1
         Left            =   1020
         TabIndex        =   159
         Text            =   "8.5 ""wat"
         Top             =   9120
         Width           =   855
      End
      Begin VB.TextBox PFD_furnace 
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
         Height          =   375
         Index           =   2
         Left            =   10560
         TabIndex        =   158
         Text            =   "8.5 ""wat"
         Top             =   9120
         Width           =   855
      End
      Begin VB.TextBox TFD_windbox 
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
         Left            =   8520
         TabIndex        =   157
         Text            =   "77 F"
         Top             =   6300
         Width           =   675
      End
      Begin VB.TextBox PFD_windbox 
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
         Height          =   375
         Index           =   2
         Left            =   8220
         TabIndex        =   156
         Text            =   "8.5 ""H2o"
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox FFD_fan 
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
         Left            =   4020
         TabIndex        =   155
         Text            =   "xx kg/s"
         Top             =   1980
         Width           =   915
      End
      Begin VB.TextBox FFD_fan 
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
         Left            =   7380
         TabIndex        =   154
         Text            =   "xx kg/s"
         Top             =   1980
         Width           =   915
      End
      Begin VB.TextBox PFD_fan 
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
         Left            =   7380
         TabIndex        =   153
         Text            =   "8.5 ""wat"
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox PFD_fan 
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
         Left            =   4020
         TabIndex        =   152
         Text            =   "8.5 ""wat"
         Top             =   1680
         Width           =   915
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
         Left            =   7380
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   4080
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
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
         Index           =   2
         Left            =   4740
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   4080
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOP"
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
         Left            =   7560
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   3840
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOP"
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
         Index           =   2
         Left            =   4440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   3840
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RUN"
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
         Left            =   7560
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RUN"
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
         Index           =   2
         Left            =   4440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   3600
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
         Index           =   1
         Left            =   6540
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   1140
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
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
         Left            =   5340
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1140
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOP"
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
         Left            =   7200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton OFFpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOP"
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
         Left            =   4680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FAST"
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
         Left            =   4680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton ONpump 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FAST"
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
         Left            =   7200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1275
         Index           =   1
         Left            =   1320
         Max             =   1000
         TabIndex        =   131
         Top             =   240
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1275
         Index           =   2
         Left            =   11220
         Max             =   1000
         TabIndex        =   130
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   3
         Left            =   7080
         Max             =   1000
         TabIndex        =   129
         Top             =   9060
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   2
         Left            =   7080
         Max             =   1000
         TabIndex        =   128
         Top             =   10440
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   1
         Left            =   4140
         Max             =   1000
         TabIndex        =   127
         Top             =   9060
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Index           =   0
         Left            =   4140
         Max             =   1000
         TabIndex        =   126
         Top             =   10440
         Width           =   1335
      End
      Begin VB.CommandButton Xover_damper 
         BackColor       =   &H0000FF00&
         Caption         =   "PORT"
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
         Index           =   2
         Left            =   6540
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   3540
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton Xover_damper 
         BackColor       =   &H0000FF00&
         Caption         =   "STBD"
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
         Left            =   5580
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   3540
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton Xover_damper 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BOTH"
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
         Left            =   6060
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   3540
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.TextBox Fuel_pos_stbd 
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
         Left            =   4800
         TabIndex        =   121
         Text            =   "xx %"
         Top             =   6840
         Width           =   855
      End
      Begin VB.TextBox Fuel_pos_port 
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
         Left            =   7260
         TabIndex        =   120
         Text            =   "xx %"
         Top             =   6840
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   300
         Top             =   6180
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   7380
         TabIndex        =   8
         Top             =   4380
         Width           =   435
      End
      Begin VB.CommandButton SLOW_FD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SLOW"
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
         Left            =   4680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   900
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton SLOW_FD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SLOW"
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
         Left            =   7200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   900
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton BYPASS_Air_heater 
         BackColor       =   &H0000FFFF&
         Caption         =   "By-Passed"
         Height          =   375
         Index           =   1
         Left            =   180
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Yellow = By-passed"
         Top             =   4620
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton BYPASS_Air_heater 
         BackColor       =   &H00C0C0C0&
         Caption         =   "By-Passed"
         Height          =   375
         Index           =   2
         Left            =   11460
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Yellow = By-passed"
         Top             =   4680
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.TextBox FFD_sys 
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
         Left            =   3900
         TabIndex        =   3
         Text            =   "xx kg/s"
         Top             =   2880
         Width           =   915
      End
      Begin VB.TextBox FFD_sys 
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
         Left            =   7860
         TabIndex        =   2
         Text            =   "xx kg/s"
         Top             =   2880
         Width           =   915
      End
      Begin VB.TextBox PFD_windbox 
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
         Height          =   375
         Index           =   1
         Left            =   3420
         TabIndex        =   1
         Text            =   "8.5 ""wat"
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox TFD_windbox 
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
         Left            =   3420
         TabIndex        =   0
         Text            =   "77 F"
         Top             =   6240
         Width           =   675
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   0
         Left            =   6180
         TabIndex        =   99
         Top             =   4800
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   1
         Left            =   6180
         TabIndex        =   100
         Top             =   5820
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   2
         Left            =   6840
         TabIndex        =   101
         Top             =   5340
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   103
         Top             =   7560
         Width           =   675
         _ExtentX        =   1191
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
         Index           =   4
         Left            =   6660
         TabIndex        =   104
         Top             =   7560
         Width           =   675
         _ExtentX        =   1191
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
         Index           =   11
         Left            =   7020
         TabIndex        =   105
         Top             =   8100
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   5
         Left            =   8100
         TabIndex        =   106
         Top             =   8100
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
         Index           =   6
         Left            =   7020
         TabIndex        =   107
         Top             =   9660
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   7
         Left            =   8100
         TabIndex        =   108
         Top             =   9660
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
         Index           =   8
         Left            =   5280
         TabIndex        =   109
         Top             =   8100
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   9
         Left            =   5280
         TabIndex        =   110
         Top             =   9600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   767
         TopColour       =   65280
         BottomColour    =   65280
         OrientVertical  =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   435
         Index           =   10
         Left            =   4110
         TabIndex        =   111
         Top             =   8100
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
         Index           =   12
         Left            =   4110
         TabIndex        =   98
         Top             =   9600
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
         Height          =   375
         Index           =   13
         Left            =   5640
         TabIndex        =   112
         Top             =   10560
         Width           =   675
         _ExtentX        =   1191
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
         Index           =   14
         Left            =   6660
         TabIndex        =   113
         Top             =   10620
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         TopColour       =   255
         BottomColour    =   255
         OpenTopLeft     =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
         ValveType       =   1
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   0
         Left            =   5460
         TabIndex        =   148
         Top             =   780
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   32768
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   1
         Left            =   6300
         TabIndex        =   149
         Top             =   840
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   32768
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   2
         Left            =   5040
         TabIndex        =   150
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   32768
         CentrLineWidth  =   2
      End
      Begin PumpPrj.Pump HavenPump1 
         Height          =   495
         Index           =   3
         Left            =   7020
         TabIndex        =   151
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         CentrifugalPump =   0   'False
         OrientLeft      =   0   'False
         LineCol         =   32768
         CentrLineWidth  =   2
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   15
         Left            =   5100
         TabIndex        =   191
         Top             =   6240
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin HavenValve1_2Prj.HavenValveV1_2 Valve 
         Height          =   375
         Index           =   16
         Left            =   7200
         TabIndex        =   192
         Top             =   6300
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   661
         TopColour       =   255
         OpenTopLeft     =   0   'False
         EnabledBtmRight =   0   'False
         OpenBtmRight    =   0   'False
         ColourClosed    =   255
         BackColour      =   12632256
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   11850
         Top             =   6240
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   11850
         Top             =   5970
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   11850
         Top             =   6540
         Width           =   195
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Removed"
         Height          =   225
         Index           =   0
         Left            =   12180
         TabIndex        =   216
         Top             =   5940
         Width           =   795
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "DO"
         Height          =   225
         Index           =   2
         Left            =   12180
         TabIndex        =   215
         Top             =   6240
         Width           =   585
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "HO"
         Height          =   225
         Index           =   3
         Left            =   12180
         TabIndex        =   214
         Top             =   6510
         Width           =   585
      End
      Begin VB.Label Vent 
         BackStyle       =   0  'Transparent
         Caption         =   "Burner Tip"
         Height          =   225
         Index           =   4
         Left            =   11850
         TabIndex        =   213
         Top             =   5700
         Width           =   1335
      End
      Begin VB.Shape Shape22 
         Height          =   1155
         Left            =   11730
         Top             =   5670
         Width           =   1455
      End
      Begin VB.Line Line87 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6840
         X2              =   6840
         Y1              =   6600
         Y2              =   6780
      End
      Begin VB.Line Line45 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   5820
         X2              =   5820
         Y1              =   6600
         Y2              =   6780
      End
      Begin VB.Label FD_position_label 
         Alignment       =   1  'Right Justify
         Caption         =   "100 %"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   204
         Top             =   180
         Width           =   495
      End
      Begin VB.Label FD_position_label 
         Caption         =   "100 %"
         Height          =   255
         Index           =   2
         Left            =   9300
         TabIndex        =   203
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "Oxygen"
         Height          =   225
         Left            =   9480
         TabIndex        =   202
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "Oxygen"
         Height          =   225
         Left            =   2280
         TabIndex        =   201
         Top             =   660
         Width           =   735
      End
      Begin VB.Shape Burner2_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   9780
         Shape           =   3  'Circle
         Top             =   7920
         Width           =   135
      End
      Begin VB.Shape Burner2_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   7980
         Width           =   135
      End
      Begin VB.Shape Burner1_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   9780
         Shape           =   3  'Circle
         Top             =   9420
         Width           =   135
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   7080
         TabIndex        =   196
         Top             =   10980
         Width           =   255
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   6000
         TabIndex        =   195
         Top             =   10920
         Width           =   255
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   7080
         TabIndex        =   194
         Top             =   7320
         Width           =   255
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   6060
         TabIndex        =   193
         Top             =   7320
         Width           =   255
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   9780
         Top             =   8340
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   9780
         Top             =   9840
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   2580
         Top             =   8340
         Width           =   195
      End
      Begin VB.Shape Burner_tip 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   2580
         Top             =   9840
         Width           =   195
      End
      Begin VB.Shape Shape19 
         BorderWidth     =   3
         Height          =   255
         Left            =   6180
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   3
         Height          =   255
         Left            =   10140
         Shape           =   3  'Circle
         Top             =   4620
         Width           =   195
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   3
         Height          =   255
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   4620
         Width           =   195
      End
      Begin VB.Label Label23 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6300
         TabIndex        =   102
         Top             =   5400
         Width           =   105
      End
      Begin VB.Line Line85 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   7020
         X2              =   6360
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Label Smoke_label 
         Alignment       =   2  'Center
         Caption         =   "Smoke"
         Height          =   225
         Index           =   4
         Left            =   11400
         TabIndex        =   185
         Top             =   7980
         Width           =   915
      End
      Begin VB.Label Smoke_label 
         Alignment       =   2  'Center
         Caption         =   "Smoke"
         Height          =   225
         Index           =   3
         Left            =   11340
         TabIndex        =   184
         Top             =   9540
         Width           =   915
      End
      Begin VB.Label Smoke_label 
         Alignment       =   2  'Center
         Caption         =   "Smoke"
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   183
         Top             =   7980
         Width           =   915
      End
      Begin VB.Label Smoke_label 
         Alignment       =   2  'Center
         Caption         =   "Smoke"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   182
         Top             =   9480
         Width           =   915
      End
      Begin VB.Line Line24 
         BorderStyle     =   3  'Dot
         X1              =   5280
         X2              =   6120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.OLE Rotary_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Index           =   0
         Left            =   2340
         OleObjectBlob   =   "Forced_Draft.frx":1004
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Rotary_Heater_18_A.vsd"
         TabIndex        =   132
         Top             =   3420
         Width           =   1935
      End
      Begin VB.OLE Rotary_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Index           =   1
         Left            =   2340
         OleObjectBlob   =   "Forced_Draft.frx":4E1C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Rotary_Heater_18_B.vsd"
         TabIndex        =   133
         Top             =   3420
         Width           =   1935
      End
      Begin VB.OLE Rotary_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Index           =   0
         Left            =   8220
         OleObjectBlob   =   "Forced_Draft.frx":9234
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Rotary_Heater_18_A.vsd"
         TabIndex        =   134
         Top             =   3420
         Width           =   1995
      End
      Begin VB.OLE Rotary_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   855
         Index           =   1
         Left            =   8220
         OleObjectBlob   =   "Forced_Draft.frx":D04C
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\STEAM_SIM\Images\Rotary_Heater_18_B.vsd"
         TabIndex        =   135
         Top             =   3420
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rotary Air Heater"
         Height          =   255
         Left            =   1740
         TabIndex        =   91
         Top             =   3180
         Width           =   1305
      End
      Begin VB.Line Line9 
         BorderWidth     =   3
         X1              =   2220
         X2              =   2220
         Y1              =   4380
         Y2              =   3450
      End
      Begin VB.Line Heater_bypass_stbd 
         BorderWidth     =   4
         Index           =   0
         X1              =   2220
         X2              =   1320
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   465
         Left            =   8955
         Shape           =   4  'Rounded Rectangle
         Top             =   2820
         Width           =   525
      End
      Begin VB.Label Label25 
         Caption         =   "No 1 Burner"
         Height          =   225
         Left            =   10440
         TabIndex        =   115
         Top             =   9780
         Width           =   1035
      End
      Begin VB.Label Label24 
         Caption         =   "No 2 Burner"
         Height          =   225
         Left            =   10500
         TabIndex        =   114
         Top             =   8340
         Width           =   1035
      End
      Begin VB.Line Line34 
         BorderWidth     =   3
         X1              =   8220
         X2              =   8220
         Y1              =   60
         Y2              =   1530
      End
      Begin VB.Line Line75 
         BorderWidth     =   3
         X1              =   12510
         X2              =   8520
         Y1              =   11340
         Y2              =   11340
      End
      Begin VB.Line Line32 
         BorderWidth     =   3
         X1              =   4050
         X2              =   0
         Y1              =   11340
         Y2              =   11340
      End
      Begin VB.Line Line21 
         BorderWidth     =   3
         X1              =   4080
         X2              =   4080
         Y1              =   6660
         Y2              =   11340
      End
      Begin VB.Line Line86 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   7020
         X2              =   6360
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line78 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   7020
         X2              =   7020
         Y1              =   6360
         Y2              =   4620
      End
      Begin VB.Line Line82 
         BorderWidth     =   3
         X1              =   3300
         X2              =   3300
         Y1              =   9780
         Y2              =   10080
      End
      Begin VB.Line Line83 
         BorderWidth     =   3
         X1              =   9240
         X2              =   9240
         Y1              =   8280
         Y2              =   8580
      End
      Begin VB.Shape Shape15 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   8460
         Top             =   8340
         Width           =   1545
      End
      Begin VB.Line Line84 
         BorderWidth     =   3
         X1              =   9240
         X2              =   9240
         Y1              =   9780
         Y2              =   10080
      End
      Begin VB.Line Line79 
         BorderWidth     =   3
         X1              =   3300
         X2              =   3300
         Y1              =   8280
         Y2              =   8580
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":11464
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   97
         Top             =   180
         Width           =   915
      End
      Begin VB.Line Line28 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2610
         Y1              =   8610
         Y2              =   8460
      End
      Begin VB.Label Label5 
         Caption         =   "Air to boiler front and to Burner registers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3390
         TabIndex        =   94
         Top             =   4380
         Width           =   885
      End
      Begin VB.Line Line51 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1770
         X2              =   3480
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   1620
         X2              =   120
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3690
         Shape           =   4  'Rounded Rectangle
         Top             =   750
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   4320
         X2              =   8190
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   6090
         Shape           =   4  'Rounded Rectangle
         Top             =   1980
         Width           =   375
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   3
         X1              =   6300
         X2              =   6300
         Y1              =   2550
         Y2              =   3270
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   3330
         X2              =   5460
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   6270
         X2              =   6270
         Y1              =   1950
         Y2              =   1560
      End
      Begin VB.Line Line8 
         BorderWidth     =   3
         X1              =   4350
         X2              =   8130
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line11 
         BorderWidth     =   3
         X1              =   1920
         X2              =   1920
         Y1              =   60
         Y2              =   1710
      End
      Begin VB.Line Line12 
         BorderWidth     =   3
         X1              =   1200
         X2              =   1890
         Y1              =   2520
         Y2              =   1710
      End
      Begin VB.Line Line14 
         BorderWidth     =   3
         X1              =   1200
         X2              =   1620
         Y1              =   4830
         Y2              =   5670
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   2130
         Shape           =   4  'Rounded Rectangle
         Top             =   4410
         Width           =   255
      End
      Begin VB.Line Line17 
         BorderWidth     =   4
         X1              =   3300
         X2              =   3300
         Y1              =   3300
         Y2              =   4290
      End
      Begin VB.Line Line18 
         BorderStyle     =   2  'Dash
         X1              =   3810
         X2              =   4320
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Port Boiler Forced Draft Fan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6660
         TabIndex        =   93
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Motor for Cross Over Damper"
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
         Left            =   6570
         TabIndex        =   92
         Top             =   1560
         Width           =   825
      End
      Begin VB.Line Line19 
         BorderWidth     =   3
         X1              =   1620
         X2              =   1620
         Y1              =   5670
         Y2              =   7470
      End
      Begin VB.Label Label12 
         Caption         =   "Economiser"
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
         Left            =   1830
         TabIndex        =   90
         Top             =   6690
         Width           =   1155
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2010
         Shape           =   3  'Circle
         Top             =   6330
         Width           =   135
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2730
         Shape           =   3  'Circle
         Top             =   6360
         Width           =   135
      End
      Begin VB.Line Line59 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3150
         X2              =   3150
         Y1              =   6660
         Y2              =   6930
      End
      Begin VB.Line Line57 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1740
         X2              =   3150
         Y1              =   6630
         Y2              =   6630
      End
      Begin VB.Line Line56 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1740
         X2              =   1740
         Y1              =   6300
         Y2              =   6600
      End
      Begin VB.Line Line55 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3180
         X2              =   3180
         Y1              =   6060
         Y2              =   6270
      End
      Begin VB.Line Line54 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1740
         X2              =   1740
         Y1              =   5760
         Y2              =   6030
      End
      Begin VB.Line Line52 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1740
         X2              =   3180
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line53 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1740
         X2              =   3180
         Y1              =   6270
         Y2              =   6270
      End
      Begin VB.Image Image8 
         Height          =   255
         Left            =   4950
         Picture         =   "Forced_Draft.frx":1367C
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   255
      End
      Begin VB.Image Image3 
         Height          =   255
         Left            =   3720
         Picture         =   "Forced_Draft.frx":13986
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   3660
         Picture         =   "Forced_Draft.frx":13C90
         Stretch         =   -1  'True
         Top             =   5940
         Width           =   255
      End
      Begin VB.Image Image11 
         Height          =   255
         Left            =   2160
         Picture         =   "Forced_Draft.frx":13F9A
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   2280
         Picture         =   "Forced_Draft.frx":142A4
         Stretch         =   -1  'True
         Top             =   5790
         Width           =   255
      End
      Begin VB.Line Line20 
         BorderWidth     =   3
         X1              =   4350
         X2              =   4080
         Y1              =   5910
         Y2              =   6690
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4410
         Top             =   840
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "Uptakes"
         Height          =   345
         Left            =   2250
         TabIndex        =   89
         Top             =   1320
         Width           =   735
      End
      Begin VB.Line Line25 
         BorderStyle     =   3  'Dot
         X1              =   6630
         X2              =   7470
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Line Heater_bypass_stbd 
         BorderWidth     =   4
         Index           =   4
         Visible         =   0   'False
         X1              =   3180
         X2              =   2280
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Line Line58 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1170
         X2              =   3150
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   4200
         TabIndex        =   88
         Top             =   7800
         Width           =   255
      End
      Begin VB.Line Line26 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   5790
         X2              =   4500
         Y1              =   8430
         Y2              =   8430
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2580
         Y1              =   8190
         Y2              =   8340
      End
      Begin VB.Label Label7 
         Caption         =   "No 2 Burner"
         Height          =   225
         Left            =   1260
         TabIndex        =   87
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   4200
         TabIndex        =   86
         Top             =   9300
         Width           =   255
      End
      Begin VB.Line Line29 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   5790
         X2              =   4380
         Y1              =   9900
         Y2              =   9900
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   2340
         X2              =   2580
         Y1              =   9690
         Y2              =   9840
      End
      Begin VB.Line Line31 
         BorderWidth     =   2
         X1              =   2310
         X2              =   2580
         Y1              =   10110
         Y2              =   9960
      End
      Begin VB.Label Label9 
         Caption         =   "No 1 Burner"
         Height          =   225
         Left            =   1290
         TabIndex        =   85
         Top             =   9750
         Width           =   1035
      End
      Begin VB.Shape Shape10 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   2550
         Top             =   9840
         Width           =   1545
      End
      Begin VB.Shape Shape9 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   2550
         Top             =   8340
         Width           =   1545
      End
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   345
         Left            =   3420
         OleObjectBlob   =   "Forced_Draft.frx":145AE
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisCFfan.vsd"
         TabIndex        =   84
         Top             =   1260
         Width           =   795
      End
      Begin VB.Line Line23 
         BorderWidth     =   3
         X1              =   3330
         X2              =   3330
         Y1              =   60
         Y2              =   2820
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   465
         Left            =   3045
         Shape           =   4  'Rounded Rectangle
         Top             =   2820
         Width           =   525
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   3810
         X2              =   3810
         Y1              =   1320
         Y2              =   1110
      End
      Begin VB.Line Line10 
         BorderWidth     =   3
         X1              =   4350
         X2              =   4350
         Y1              =   3420
         Y2              =   5910
      End
      Begin VB.Line Heater_bypass_stbd 
         BorderWidth     =   4
         Index           =   3
         Visible         =   0   'False
         X1              =   2220
         X2              =   3060
         Y1              =   4740
         Y2              =   5220
      End
      Begin VB.Line Heater_bypass_stbd 
         BorderWidth     =   4
         Index           =   2
         Visible         =   0   'False
         X1              =   2220
         X2              =   2220
         Y1              =   4740
         Y2              =   5580
      End
      Begin VB.Line Heater_bypass_stbd 
         BorderWidth     =   4
         Index           =   1
         Visible         =   0   'False
         X1              =   2220
         X2              =   1560
         Y1              =   4740
         Y2              =   5400
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   2
         Visible         =   0   'False
         X1              =   6300
         X2              =   6000
         Y1              =   2580
         Y2              =   3240
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   1
         Visible         =   0   'False
         X1              =   6300
         X2              =   5640
         Y1              =   2520
         Y2              =   2940
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   0
         Visible         =   0   'False
         X1              =   6300
         X2              =   5460
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   4
         Visible         =   0   'False
         X1              =   6300
         X2              =   6660
         Y1              =   2520
         Y2              =   3240
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   5
         Visible         =   0   'False
         X1              =   6300
         X2              =   7020
         Y1              =   2520
         Y2              =   2940
      End
      Begin VB.Line Cross_over 
         BorderWidth     =   4
         Index           =   6
         Visible         =   0   'False
         X1              =   6300
         X2              =   7200
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Shape Ign1_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   1800
         Shape           =   2  'Oval
         Top             =   10200
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   1800
         Shape           =   2  'Oval
         Top             =   8880
         Width           =   375
      End
      Begin VB.Shape Ign1_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   4500
         Shape           =   2  'Oval
         Top             =   10200
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   1
         Left            =   4500
         Shape           =   2  'Oval
         Top             =   8880
         Width           =   375
      End
      Begin VB.Shape Burner1_shape 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   9420
         Width           =   135
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   10530
         X2              =   10530
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Line Line15 
         BorderWidth     =   3
         X1              =   11340
         X2              =   11340
         Y1              =   2520
         Y2              =   4830
      End
      Begin VB.Line Line22 
         BorderWidth     =   3
         X1              =   10800
         X2              =   11340
         Y1              =   5670
         Y2              =   4830
      End
      Begin VB.Line Line33 
         BorderWidth     =   3
         X1              =   10530
         X2              =   11340
         Y1              =   1710
         Y2              =   2520
      End
      Begin VB.Line Line35 
         BorderWidth     =   3
         X1              =   9210
         X2              =   9210
         Y1              =   60
         Y2              =   3480
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   345
         Left            =   8280
         OleObjectBlob   =   "Forced_Draft.frx":1D7C6
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\VisCFfan.vsd"
         TabIndex        =   83
         Top             =   1260
         Width           =   795
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   8550
         Shape           =   4  'Rounded Rectangle
         Top             =   750
         Width           =   255
      End
      Begin VB.Line Line36 
         BorderWidth     =   2
         X1              =   8670
         X2              =   8670
         Y1              =   1290
         Y2              =   1080
      End
      Begin VB.Line Line37 
         BorderWidth     =   3
         X1              =   7080
         X2              =   9210
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Line Line38 
         BorderWidth     =   3
         X1              =   8130
         X2              =   8130
         Y1              =   3420
         Y2              =   5910
      End
      Begin VB.Line Line40 
         BorderWidth     =   3
         X1              =   12360
         X2              =   10860
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line41 
         BorderWidth     =   3
         X1              =   8520
         X2              =   8520
         Y1              =   6690
         Y2              =   11340
      End
      Begin VB.Line Line42 
         BorderWidth     =   3
         X1              =   8130
         X2              =   8490
         Y1              =   5910
         Y2              =   6690
      End
      Begin VB.Line Line43 
         BorderWidth     =   3
         X1              =   10290
         X2              =   10290
         Y1              =   4410
         Y2              =   3420
      End
      Begin VB.Shape Shape14 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   8460
         Top             =   9840
         Width           =   1545
      End
      Begin VB.Line Line46 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6840
         X2              =   8160
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Line Line47 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6840
         X2              =   8160
         Y1              =   9960
         Y2              =   9960
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   8160
         TabIndex        =   82
         Top             =   7860
         Width           =   255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         Height          =   255
         Left            =   8160
         TabIndex        =   81
         Top             =   9420
         Width           =   255
      End
      Begin VB.Line Heater_bypass_port 
         BorderWidth     =   4
         Index           =   4
         X1              =   10230
         X2              =   9330
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Line Heater_bypass_port 
         BorderWidth     =   4
         Index           =   3
         Visible         =   0   'False
         X1              =   10230
         X2              =   9570
         Y1              =   4740
         Y2              =   5400
      End
      Begin VB.Line Heater_bypass_port 
         BorderWidth     =   4
         Index           =   2
         Visible         =   0   'False
         X1              =   10230
         X2              =   10230
         Y1              =   4740
         Y2              =   5580
      End
      Begin VB.Line Heater_bypass_port 
         BorderWidth     =   4
         Index           =   1
         Visible         =   0   'False
         X1              =   10230
         X2              =   11070
         Y1              =   4740
         Y2              =   5220
      End
      Begin VB.Line Heater_bypass_port 
         BorderWidth     =   4
         Index           =   0
         Visible         =   0   'False
         X1              =   11160
         X2              =   10260
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Shape Shape16 
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   10140
         Shape           =   4  'Rounded Rectangle
         Top             =   4410
         Width           =   255
      End
      Begin VB.Line Line48 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9300
         X2              =   11280
         Y1              =   6990
         Y2              =   6990
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   9840
         Picture         =   "Forced_Draft.frx":269DE
         Stretch         =   -1  'True
         Top             =   5820
         Width           =   255
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9300
         X2              =   10740
         Y1              =   6300
         Y2              =   6300
      End
      Begin VB.Line Line50 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9300
         X2              =   10740
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line60 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   10740
         X2              =   10740
         Y1              =   5790
         Y2              =   6060
      End
      Begin VB.Line Line61 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9300
         X2              =   9300
         Y1              =   6090
         Y2              =   6300
      End
      Begin VB.Line Line62 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   10740
         X2              =   10740
         Y1              =   6330
         Y2              =   6630
      End
      Begin VB.Line Line63 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9330
         X2              =   10740
         Y1              =   6660
         Y2              =   6660
      End
      Begin VB.Line Line64 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9300
         X2              =   9300
         Y1              =   6690
         Y2              =   6960
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   10290
         Shape           =   3  'Circle
         Top             =   6390
         Width           =   135
      End
      Begin VB.Shape Shape18 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   9570
         Shape           =   3  'Circle
         Top             =   6360
         Width           =   135
      End
      Begin VB.Label Label15 
         Caption         =   "Economiser"
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
         Left            =   9570
         TabIndex        =   80
         Top             =   6720
         Width           =   1155
      End
      Begin VB.Line Line65 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   9060
         X2              =   10740
         Y1              =   5790
         Y2              =   5790
      End
      Begin VB.Line Line39 
         BorderWidth     =   3
         X1              =   10830
         X2              =   10830
         Y1              =   5670
         Y2              =   7470
      End
      Begin VB.Label Label16 
         Caption         =   "Rotary Air Heater"
         Height          =   255
         Left            =   9540
         TabIndex        =   79
         Top             =   3180
         Width           =   1305
      End
      Begin VB.Line Line66 
         BorderWidth     =   2
         X1              =   9990
         X2              =   10230
         Y1              =   8490
         Y2              =   8640
      End
      Begin VB.Line Line67 
         BorderWidth     =   2
         X1              =   9990
         X2              =   10260
         Y1              =   8340
         Y2              =   8190
      End
      Begin VB.Shape Ign2_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   10320
         Shape           =   2  'Oval
         Top             =   8880
         Width           =   375
      End
      Begin VB.Shape Ign2_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   7710
         Shape           =   2  'Oval
         Top             =   8880
         Width           =   375
      End
      Begin VB.Shape Ign1_shape_out 
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   7650
         Shape           =   2  'Oval
         Top             =   10200
         Width           =   375
      End
      Begin VB.Line Line70 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   5820
         X2              =   6840
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line72 
         BorderWidth     =   2
         X1              =   9990
         X2              =   10230
         Y1              =   9990
         Y2              =   10140
      End
      Begin VB.Line Line73 
         BorderWidth     =   2
         X1              =   9990
         X2              =   10260
         Y1              =   9840
         Y2              =   9690
      End
      Begin VB.Shape Ign1_shape_in 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   75
         Index           =   2
         Left            =   10320
         Shape           =   2  'Oval
         Top             =   10200
         Width           =   375
      End
      Begin VB.Shape Shape20 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   7890
         Top             =   840
         Width           =   165
      End
      Begin VB.Label Label17 
         Caption         =   "Stbd Boiler Forced Draft Fan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4500
         TabIndex        =   78
         Top             =   180
         Width           =   1365
      End
      Begin VB.Line Line74 
         BorderStyle     =   2  'Dash
         X1              =   8070
         X2              =   8580
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Port F.O. Recirculation"
         Height          =   615
         Left            =   6360
         TabIndex        =   77
         Top             =   11220
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Stbd F.O. Recirculation"
         Height          =   495
         Left            =   5340
         TabIndex        =   76
         Top             =   11220
         Width           =   1035
      End
      Begin VB.Label Label20 
         Caption         =   "Uptakes"
         Height          =   345
         Left            =   9480
         TabIndex        =   74
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   8520
         Picture         =   "Forced_Draft.frx":26CE8
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   9960
         Picture         =   "Forced_Draft.frx":26FF2
         Stretch         =   -1  'True
         Top             =   2460
         Width           =   255
      End
      Begin VB.Image Image9 
         Height          =   255
         Left            =   7500
         Picture         =   "Forced_Draft.frx":272FC
         Stretch         =   -1  'True
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Air to boiler front and to Burner registers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8280
         TabIndex        =   73
         Top             =   4380
         Width           =   885
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   8580
         Picture         =   "Forced_Draft.frx":27606
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "RED=NOT BY_PASSED"
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line Line81 
         BorderWidth     =   3
         X1              =   9240
         X2              =   9240
         Y1              =   10500
         Y2              =   11340
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         HelpContextID   =   60
         Index           =   1
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":27910
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   68
         Top             =   180
         Width           =   915
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":29928
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   67
         Top             =   180
         Width           =   915
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":2B940
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   64
         Top             =   8880
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":2D958
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   63
         Top             =   9300
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":2F970
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   62
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":31988
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   61
         Top             =   7800
         Width           =   675
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":339A0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   60
         Top             =   8820
         Width           =   675
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":359B8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   59
         Top             =   9240
         Width           =   675
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":379D0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   58
         Top             =   10380
         Width           =   675
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":399E8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   56
         Top             =   8820
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":3BA00
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   55
         Top             =   9300
         Width           =   675
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":3DA18
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   54
         Top             =   10380
         Width           =   675
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":3FA30
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   52
         Top             =   8820
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":41A48
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   51
         Top             =   9300
         Width           =   675
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":43A60
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   50
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":45A78
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   49
         Top             =   9300
         Width           =   675
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":47A90
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   48
         Top             =   8880
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":49AA8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   47
         Top             =   10380
         Width           =   675
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":4BAC0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   46
         Top             =   7800
         Width           =   675
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":4DAD8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   45
         Top             =   9300
         Width           =   675
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":4FAF0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   44
         Top             =   8880
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":51B08
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   43
         Top             =   10380
         Width           =   675
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":53B20
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   42
         Top             =   7800
         Width           =   675
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":55B38
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   41
         Top             =   9240
         Width           =   675
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":57B50
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   40
         Top             =   8880
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":59B68
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   39
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":5BB80
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   38
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":5DB98
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   37
         Top             =   9240
         Width           =   675
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":5FBB0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   36
         Top             =   8880
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":61BC8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   35
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":63BE0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   34
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":65BF8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   33
         Top             =   9240
         Width           =   675
      End
      Begin VB.Line Line13 
         BorderWidth     =   3
         X1              =   1200
         X2              =   1200
         Y1              =   2520
         Y2              =   4830
      End
      Begin VB.Line Line44 
         BorderWidth     =   3
         X1              =   9240
         X2              =   9240
         Y1              =   4260
         Y2              =   7920
      End
      Begin VB.OLE Register2_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":67C10
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   32
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":69C28
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   31
         Top             =   8880
         Width           =   675
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":6BC40
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   30
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":6DC58
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   28
         Top             =   9300
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":6FC70
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   27
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register1_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":71C88
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   26
         Top             =   9300
         Width           =   675
      End
      Begin VB.Line Line77 
         BorderWidth     =   3
         X1              =   3300
         X2              =   3300
         Y1              =   10500
         Y2              =   11340
      End
      Begin VB.Line Line80 
         BorderWidth     =   3
         X1              =   9240
         X2              =   9240
         Y1              =   8940
         Y2              =   9420
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   4290
         X2              =   4290
         Y1              =   60
         Y2              =   1530
      End
      Begin VB.OLE Register2_stbd_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":73CA0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   23
         Top             =   8880
         Width           =   675
      End
      Begin VB.Line Line16 
         BorderWidth     =   3
         X1              =   3300
         X2              =   3300
         Y1              =   4320
         Y2              =   7920
      End
      Begin VB.Line Line76 
         BorderWidth     =   3
         X1              =   3300
         X2              =   3300
         Y1              =   9000
         Y2              =   9480
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":75CB8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   22
         Top             =   8820
         Width           =   675
      End
      Begin VB.OLE Register2_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":77CD0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   21
         Top             =   8820
         Width           =   675
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   8580
         OleObjectBlob   =   "Forced_Draft.frx":79CE8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   20
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE Register1_port_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":7BD00
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   19
         Top             =   9240
         Width           =   675
      End
      Begin VB.OLE Register1_port_bot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":7DD18
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   18
         Top             =   10320
         Width           =   675
      End
      Begin VB.OLE FD_damper_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   8280
         OleObjectBlob   =   "Forced_Draft.frx":7FD30
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   17
         Top             =   240
         Width           =   915
      End
      Begin VB.OLE FD_damper_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   8220
         OleObjectBlob   =   "Forced_Draft.frx":81F48
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.OLE FD_damper_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   8220
         OleObjectBlob   =   "Forced_Draft.frx":83F60
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.OLE FD_damper_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   8280
         OleObjectBlob   =   "Forced_Draft.frx":85F78
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":87F90
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   13
         Top             =   180
         Width           =   915
      End
      Begin VB.OLE FD_damper_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   8280
         OleObjectBlob   =   "Forced_Draft.frx":89FA8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   12
         Top             =   240
         Width           =   915
      End
      Begin VB.OLE FD_damper_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   8280
         OleObjectBlob   =   "Forced_Draft.frx":8BFC0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   11
         Top             =   240
         Width           =   915
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":8DFD8
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   10
         Top             =   180
         Width           =   915
      End
      Begin VB.OLE FD_damper_stbd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":8FFF0
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   9
         Top             =   180
         Width           =   915
      End
      Begin VB.OLE Register2_port 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   9
         Left            =   8520
         OleObjectBlob   =   "Forced_Draft.frx":92008
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   66
         Top             =   8880
         Width           =   675
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   3
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   6180
         Shape           =   3  'Circle
         Top             =   5340
         Width           =   345
      End
      Begin VB.Line Line71 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6360
         X2              =   6360
         Y1              =   6600
         Y2              =   4500
      End
      Begin VB.Label Label4 
         Caption         =   "Air Heater By Pass Damper"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1770
         TabIndex        =   95
         Top             =   4860
         Width           =   765
      End
      Begin VB.Label Label14 
         Caption         =   "Air Heater By Pass Damper"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9780
         TabIndex        =   96
         Top             =   4860
         Width           =   765
      End
      Begin VB.Line Line69 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   6840
         X2              =   6840
         Y1              =   7200
         Y2              =   11100
      End
      Begin VB.Line Line68 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   5820
         X2              =   5820
         Y1              =   7200
         Y2              =   11130
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":94020
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register_Closed_Air.vsd"
         TabIndex        =   65
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":96038
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register30_air.vsd"
         TabIndex        =   53
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":98050
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register45_air.vsd"
         TabIndex        =   29
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":9A068
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register60_airvsd.vsd"
         TabIndex        =   25
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":9C080
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Registeropen_air.vsd"
         TabIndex        =   24
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE register2_stbd_top 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3360
         OleObjectBlob   =   "Forced_Draft.frx":9E098
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Steam Sim\Register15_air.vsd"
         TabIndex        =   57
         Top             =   7740
         Width           =   675
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   1
         Left            =   5700
         OleObjectBlob   =   "Forced_Draft.frx":A00B0
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   70
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   2
         Left            =   5700
         OleObjectBlob   =   "Forced_Draft.frx":A80C8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   69
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   1
         Left            =   6720
         OleObjectBlob   =   "Forced_Draft.frx":B00E0
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green33_mushroom.vsd"
         TabIndex        =   117
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   2
         Left            =   6720
         OleObjectBlob   =   "Forced_Draft.frx":B80F8
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green66_mushroom.vsd"
         TabIndex        =   118
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   6720
         OleObjectBlob   =   "Forced_Draft.frx":C0110
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   119
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   5700
         OleObjectBlob   =   "Forced_Draft.frx":C7F28
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_green_mushroom.vsd"
         TabIndex        =   71
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_stbd 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   0
         Left            =   5700
         OleObjectBlob   =   "Forced_Draft.frx":CFD40
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   75
         Top             =   6780
         Width           =   495
      End
      Begin VB.OLE Fuel_VV_port 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Index           =   0
         Left            =   6720
         OleObjectBlob   =   "Forced_Draft.frx":D8558
         OLETypeAllowed  =   1  'Embedded
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Vert_right_black_mushroom.vsd"
         TabIndex        =   116
         Top             =   6780
         Width           =   495
      End
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      TabIndex        =   122
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "Forced_Draft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private set_dummy_focus As Integer
Private flash_bypass As Boolean

Private Sub Burner_ON_OFF_Click(Index As Integer)

Burner_ON(Index) = Not Burner_ON(Index)
If Burner_ON(Index) Then
    Burner_ON_OFF(Index).BackColor = RGB(0, &HFF, 0)
Else
    Burner_ON_OFF(Index).BackColor = RGB(&HC0, &HC0, &HC0)
End If

' 7/7/99 SET in Steam_Sim.Dll
Call set_STEAM_logicals(Index, Burner_ON(Index))

End Sub

Private Sub Burner1_flame_Click(Index As Integer)

' 27/8/99
Burner1_ON(Index) = Not Burner1_ON(Index)

Call set_STEAM_logicals(Index, Burner1_ON(Index))

End Sub

Private Sub Burner2_flame_Click(Index As Integer)

' 27/8/99
Burner2_ON(Index) = Not Burner2_ON(Index)

'   2/9/99
'Call set_STEAM_logicals(index, Burner2_ON(index))
Call set_STEAM_logicals(Index + 2, Burner2_ON(Index))

End Sub

Private Sub BYPASS_Air_heater_Click(Index As Integer)

set_dummy_focus = True


' 24/9/99
'If BYPASS_Air_heater(Index).BackColor = &H8080FF Then      ' PINK ( NOT By-passed )
'    status_flag = True
'Else
'    status_flag = False
'End If
'
'Call set_STEAM_logicals(7 + Index, status_flag)

If Index = 1 Then

  ' === 11/11/02 ===
  'Call set_PCDI_LSS(7)    ' STBD
   Call setPCDI(7)    ' STBD
  
Else

  ' === 11/11/02 ===
  'Call set_PCDI_LSS(106)  ' PORT
   Call setPCDI(106)  ' PORT
  
End If


End Sub

Private Sub FAST_FD_Click(Index As Integer)

set_dummy_focus = True

End Sub

Private Sub Command1_Click()

' 6/89/99
FO_Service.Show

End Sub

Private Sub Command2_Click()


' === 26/09/01 ===
status_flag = True
Call setDI(7, status_flag)


End Sub

Private Sub Command3_Click()

' === 26/09/01 ===
status_flag = False
Call setDI(7, status_flag)

End Sub

Private Sub Command4_Click()

' === 26/09/01 ===
status_flag = True
Call setDI(106, status_flag)


End Sub

Private Sub Command5_Click()

' === 26/09/01 ===
status_flag = False
Call setDI(106, status_flag)

End Sub

Private Sub Form_Load()

On Error Resume Next
Call JYAScreenInit(Me, idForced_Draft)

set_dummy_focus = True


' 26/11/99
HScroll1(0).Visible = PTT
HScroll1(1).Visible = PTT
HScroll1(2).Visible = PTT
HScroll1(3).Visible = PTT


' ===08/03/01 ===
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so
    FD_position_label(1).Visible = False
    FD_position_label(2).Visible = False
    
    ' === 22/11/01 14.093 ===
    T_Flue_gas(1).Visible = False
    T_Flue_gas(2).Visible = False
    
    
    ' === 7/12/01 ===
     BYPASS_Air_heater(1).Visible = False
     BYPASS_Air_heater(2).Visible = False
   
Else
    FD_position_label(1).Visible = True
    FD_position_label(2).Visible = True
    
    ' === 22/11/01 ===
    T_Flue_gas(1).Visible = True
    T_Flue_gas(2).Visible = True
    
    ' === 7/12/01 14.093 ===
     BYPASS_Air_heater(1).Visible = True
     BYPASS_Air_heater(2).Visible = True
    
    
End If


' === 31/01/02 ===
HavenStretch1.Back_color = Bcolour
Call copy_havenstretch_colours(Me, Bcolour)


End Sub

Private Sub Form_Resize()

On Error Resume Next
If Not (Me.WindowState = 2 And frm_Host.WindowState = 1) Then
    HavenStretch1.Visible = False
    Call JYAScreenResize(Me, idForced_Draft)
    HavenStretch1.Visible = True
End If

End Sub

Private Sub HScroll1_Change(Index As Integer)

    Call set_VSCROLL(Index + 8, HScroll1(Index).Value)   ' Air register position

End Sub

Private Sub RUN_Air_heater_Click(Index As Integer)

End Sub

Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(4=OFF,nPump)

Select Case Index

    
    Case 0
        Call setPCDI(PumpIO_LSS(4, 5))    ' STBD FD FAN OFF
    Case 1
        Call setPCDI(PumpIO_LSS(4, 17))    ' PORT FD FAN OFF
    Case 2
        Call setPCDI(PumpIO_LSS(4, 31))    ' STBD Rotary Air Heater OFF
    Case 3
        Call setPCDI(PumpIO_LSS(4, 32))    ' PORT Rotary Air Heater OFF
         
End Select


End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)

Select Case Index

    ' STEAM SIM  ( See Breaker Xref    xls file )
    
    '   9/10/99
   ' Case 0
   '     Call setPCDI(PumpIO_LSS(2, 5))    ' STBD FD FAN ON (FAST)
   '     status_flag = False
   '     Call set_STEAM_logicals(12, status_flag)

   ' Case 1
   '     Call setPCDI(PumpIO_LSS(2, 17))    ' PORT FD FAN ON (FAST)
   '     status_flag = False
   '     Call set_STEAM_logicals(13, status_flag)
        
    Case 0
        Call setPCDI(PumpIO_LSS(5, 5))    ' STBD FD FAN ON (FAST)
    Case 1
        Call setPCDI(PumpIO_LSS(5, 17))    ' PORT FD FAN ON (FAST)
        
        
        
        
    Case 2
        Call setPCDI(PumpIO_LSS(2, 31))    ' STBD Rotary Air Heater ON
    Case 3
        Call setPCDI(PumpIO_LSS(2, 32))    ' PORT Rotary Air Heater PORT ON
        
End Select


End Sub

Private Sub Oxygen_Click(Index As Integer)


' 11/9/00
Unload Alarm_set_point   ' Do this first
If Index = 1 Then
  nAlarm_LO = 21
Else
  nAlarm_LO = 68
End If
  Alarm_set_point.Show

End Sub

Private Sub SLOW_FD_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
'   Note:  PumpIO_LSS(2=ON,nPump)

Select Case Index

    
    Case 0
        Call setPCDI(PumpIO_LSS(2, 5))    ' FD FAN STBD ON (SLOW)
        
        ' === 22/11/02 ===
        Call Sleep(200)

        status_flag = True
        ' === 22/11/02 ===
        'Call set_STEAM_logicals(12, status_flag)
         Call setSTEAMlogicals(12, status_flag)
        
    Case 1
        Call setPCDI(PumpIO_LSS(2, 17))    ' FD FAN PORT ON (SLOW)
        
        
        ' === 22/11/02 ===
        Call Sleep(200)
        
        status_flag = True
       ' === 22/11/02 ===
       'Call set_STEAM_logicals(13, status_flag)
        Call setSTEAMlogicals(13, status_flag)
           
End Select




End Sub

Private Sub STOP_FD_Click(Index As Integer)

set_dummy_focus = True

End Sub

Private Sub Timer1_Timer()
Dim single_char As String * 1
Dim i As Integer
Dim VV_pos As Integer


If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

' === 28/03/03 20.008 ===
For i = 1 To 4
    F_air_burner(i).Visible = Advanced_mode
    FFO_burner(i).Visible = Advanced_mode
    B_Smoke(i).Visible = Advanced_mode
    Smoke_label(i).Visible = Advanced_mode
    Q_burn(i).Visible = Advanced_mode
Next i


' === 7/12/01 ===
flash_bypass = Not flash_bypass


' 26/11/99   HO/DO Nozzles
'equivalence(STEAM_logicals(15),HO_nozzle(1))

For i = 0 To 3

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
For i = 1 To 2
    If VB_steam_logicals(i) Then
      ' Burner1_flame(i).BackColor = RGB(0, &HFF, 0)
        Burner1_shape(i).FillColor = RGB(&HFF, 0, 0)
    Else
        Burner1_shape(i).FillColor = RGB(&HC0, &HC0, &HC0)
    End If
    If VB_steam_logicals(i + 2) Then
        Burner2_shape(i).FillColor = RGB(&HFF, 0, 0)
    Else
        Burner2_shape(i).FillColor = RGB(&HC0, &HC0, &HC0)
    End If
Next i


'               72,73           DI( BIGN1_DI(1-2) )
'               74,75           DI( BIGN2_DI(1-2) )

' Igniter shapes
For i = 1 To 2

    If VB_steam_logicals(71 + i) Then
        Ign1_shape_in(i).Visible = True
        Ign1_shape_out(i).Visible = False
    Else
        Ign1_shape_in(i).Visible = False
        Ign1_shape_out(i).Visible = True
    End If
    
    If VB_steam_logicals(73 + i) Then
        Ign2_shape_in(i).Visible = True
        Ign2_shape_out(i).Visible = False
    Else
        Ign2_shape_in(i).Visible = False
        Ign2_shape_out(i).Visible = True
    End If
    
Next i

' === 11/9/00 ===
Oxygen(1).text = Format(LSS_reals(137), "###0.0") & "%"
Oxygen(2).text = Format(LSS_reals(138), "###0.0") & "%"
Call paint_alarm(Oxygen(1), 21, 0, 0, 0)
Call paint_alarm(Oxygen(2), 68, 0, 0, 0)



B_Smoke(1).text = Format(LSS_reals(184), "###0.0") & "%"
B_Smoke(2).text = Format(LSS_reals(185), "###0.0") & "%"
B_Smoke(3).text = Format(LSS_reals(186), "###0.0") & "%"
B_Smoke(4).text = Format(LSS_reals(187), "###0.0") & "%"

Q_burn(1).text = Format(LSS_reals(188), "####0") & " kW"
Q_burn(2).text = Format(LSS_reals(189), "####0") & " kW"

' 27/6/00
'         PC_REALS(141) = Q_burn(3)
'         PC_REALS(142) = Q_burn(4)
Q_burn(3).text = Format(LSS_reals(141), "####0") & " kW"
Q_burn(4).text = Format(LSS_reals(142), "####0") & " kW"


If SI_units Then
    F_air_burner(1).text = Format(LSS_reals(180) * Flow_mult, "###0.000") & Flow_chars
    F_air_burner(2).text = Format(LSS_reals(181) * Flow_mult, "###0.000") & Flow_chars
    F_air_burner(3).text = Format(LSS_reals(182) * Flow_mult, "###0.000") & Flow_chars
    F_air_burner(4).text = Format(LSS_reals(183) * Flow_mult, "###0.000") & Flow_chars
   
    FFO_burner(1).text = Format(LSS_reals(232) * Flow_mult, "###0.000") & Flow_chars
    FFO_burner(2).text = Format(LSS_reals(233) * Flow_mult, "###0.000") & Flow_chars
    FFO_burner(3).text = Format(LSS_reals(190) * Flow_mult, "###0.000") & Flow_chars
    FFO_burner(4).text = Format(LSS_reals(191) * Flow_mult, "###0.000") & Flow_chars
    
    FFD_fan(1).text = Format(LSS_reals(176), "###0.0") & " kg/s"
    FFD_fan(2).text = Format(LSS_reals(177), "###0.0") & " kg/s"
    
    ' === 08/03/01 ===
    FFD_sys(1).text = Format(LSS_reals(174), "###0.0") & " kg/s"
    FFD_sys(2).text = Format(LSS_reals(175), "###0.0") & " kg/s"


Else
    F_air_burner(1).text = Format(LSS_reals(180) * Flow_mult / 3600#, "###0.000") & " lb/s"
    F_air_burner(2).text = Format(LSS_reals(181) * Flow_mult / 3600#, "###0.000") & " lb/s"
    F_air_burner(3).text = Format(LSS_reals(182) * Flow_mult / 3600#, "###0.000") & " lb/s"
    F_air_burner(4).text = Format(LSS_reals(183) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
    FFO_burner(1).text = Format(LSS_reals(232) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_burner(2).text = Format(LSS_reals(233) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_burner(3).text = Format(LSS_reals(190) * Flow_mult / 3600#, "###0.000") & " lb/s"
    FFO_burner(4).text = Format(LSS_reals(191) * Flow_mult / 3600#, "###0.000") & " lb/s"
    
    ' === 08/03/01 ONLY to 2D ===
   ' FFD_fan(1).text = Format(LSS_reals(176) * Flow_mult / 3600#, "###0.000") & " lb/s"
   ' FFD_fan(2).text = Format(LSS_reals(177) * Flow_mult / 3600#, "###0.000") & " lb/s"
     FFD_fan(1).text = Format(LSS_reals(176) * Flow_mult / 3600#, "###0.00") & " lb/s"
     FFD_fan(2).text = Format(LSS_reals(177) * Flow_mult / 3600#, "###0.00") & " lb/s"
    
    ' === 08/03/01 ===
    FFD_sys(1).text = Format(LSS_reals(174) * Flow_mult / 3600#, "###0.00") & " lb/s"
    FFD_sys(2).text = Format(LSS_reals(175) * Flow_mult / 3600#, "###0.00") & " lb/s"
    
End If


' === 22/01/02 15.022  ===
If SI_units Then
    PFD_uptakes(1).text = Format(LSS_reals(162) * 25.4, "##0") & " mmH20"
    PFD_uptakes(2).text = Format(LSS_reals(163) * 25.4, "##0") & " mmH20"
    
    PFD_econ_in(1).text = Format(LSS_reals(166) * 25.4, "##0") & " mmH20"
    PFD_econ_in(2).text = Format(LSS_reals(167) * 25.4, "##0") & " mmH20"
    
    PFD_furnace(1).text = Format(LSS_reals(168) * 25.4, "##0") & " mmH20"
    PFD_furnace(2).text = Format(LSS_reals(169) * 25.4, "##0") & " mmH20"
    
    PFD_windbox(1).text = Format(LSS_reals(170) * 25.4, "##0") & " mmH20"
    PFD_windbox(2).text = Format(LSS_reals(171) * 25.4, "##0") & " mmH20"
    
    PFD_fan(1).text = Format(LSS_reals(172) * 25.4, "##0") & " mmH20"
    PFD_fan(2).text = Format(LSS_reals(173) * 25.4, "##0") & " mmH20"
       
    PFD_fan(1).text = Format(LSS_reals(172) * 25.4, "##0") & " mmH20"
    PFD_fan(2).text = Format(LSS_reals(173) * 25.4, "##0") & " mmH20"
Else
    PFD_uptakes(1).text = Format(LSS_reals(162), "##0.0") & " inH20"
    PFD_uptakes(2).text = Format(LSS_reals(163), "##0.0") & " inH20"
    
    PFD_econ_in(1).text = Format(LSS_reals(166), "##0.0") & " inH20"
    PFD_econ_in(2).text = Format(LSS_reals(167), "##0.0") & " inH20"
    
    PFD_furnace(1).text = Format(LSS_reals(168), "##0.0") & " inH20"
    PFD_furnace(2).text = Format(LSS_reals(169), "##0.0") & " inH20"
    
    PFD_windbox(1).text = Format(LSS_reals(170), "##0.0") & " inH20"
    PFD_windbox(2).text = Format(LSS_reals(171), "##0.0") & " inH20"
    
    PFD_fan(1).text = Format(LSS_reals(172), "##0.0") & " inH20"
    PFD_fan(2).text = Format(LSS_reals(173), "##0.0") & " inH20"
       
    PFD_fan(1).text = Format(LSS_reals(172), "##0.0") & " inH20"
    PFD_fan(2).text = Format(LSS_reals(173), "##0.0") & " inH20"
End If
    
    
   ' === 08/03/01 Moved up ===
   ' FFD_sys(1).text = Format(LSS_reals(174), "###0.0") & " kg/s"
   ' FFD_sys(2).text = Format(LSS_reals(175), "###0.0") & " kg/s"
    

TFD_windbox(1).text = Format(LSS_reals(178) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
TFD_windbox(2).text = Format(LSS_reals(179) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

'c         === 21/11/01 ===
'          PC_REALS(825) = T_flue_gas(1)
'          PC_REALS(826) = T_flue_gas(2)
'
'c       === 22/11/01 ===
'          PC_REALS(827) = T_flue_gas_ECON_in(1)
'          PC_REALS(828) = T_flue_gas_ECON_in(2)
'          PC_REALS(829) = T_flue_gas_ECON_out(1)
'          PC_REALS(830) = T_flue_gas_ECON_out(2)
T_Flue_gas_ECON_in(1).text = Format(LSS_reals(827) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_Flue_gas_ECON_out(1).text = Format(LSS_reals(829) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char
T_Flue_gas(1).text = Format(LSS_reals(825) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char

T_Flue_gas(2).text = Format(LSS_reals(826) * Temp_mult + Temp_const, "##0") & Chr$(176) & Temp_char


' Rotary Air Heater STBD
' 24/9/99
'If pp_status_LSS(36) = 3 Then
If pp_status_LSS(31) = 3 Then
    If Rotary_stbd(0).Visible Then
        Rotary_stbd(0).Visible = False
        Rotary_stbd(1).Visible = True
    Else
        Rotary_stbd(1).Visible = False
        Rotary_stbd(0).Visible = True
    End If
Else
        Rotary_stbd(1).Visible = False
        Rotary_stbd(0).Visible = True
End If

' Rotary Air Heater PORT
' 24/9/99
'If pp_status_LSS(37) = 3 Then
If pp_status_LSS(32) = 3 Then
    If Rotary_port(0).Visible Then
        Rotary_port(0).Visible = False
        Rotary_port(1).Visible = True
    Else
        Rotary_port(1).Visible = False
        Rotary_port(0).Visible = True
    End If
Else
        Rotary_port(1).Visible = False
        Rotary_port(0).Visible = True
End If


' FD Damper STBD
Call Air_VV_position(160, VV_pos)

' === 08/03/01 ===
FD_position_label(1).Caption = Format(LSS_reals(160), "##0") & "%"


'30/11/99
If VV_pos = 0 Then VV_pos = 1

For i = 0 To 5
    If i = VV_pos Then
        FD_damper_stbd(i).Visible = True
    Else
        FD_damper_stbd(i).Visible = False
    End If
Next i

' FD Damper PORT
Call Air_VV_position(161, VV_pos)

' === 08/03/01 ===
FD_position_label(2).Caption = Format(LSS_reals(161), "##0") & "%"


'30/11/99
If VV_pos = 0 Then VV_pos = 1

For i = 0 To 5
    If i = VV_pos Then
        FD_damper_port(i).Visible = True
    Else
        FD_damper_port(i).Visible = False
    End If
Next i


' Air Registers   ( VV_pos 0 to 5 )  ,  LSS_reals(   ) 0 to 100.0

' Register1_stbd
Call Air_VV_position(219, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        Register1_stbd_top(i).Visible = True
        Register1_stbd_bot(i).Visible = True
    Else
        Register1_stbd_top(i).Visible = False
        Register1_stbd_bot(i).Visible = False
    End If
Next i

' Register2_stbd
Call Air_VV_position(439, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        register2_stbd_top(i).Visible = True
        Register2_stbd_bot(i).Visible = True
    Else
        register2_stbd_top(i).Visible = False
        Register2_stbd_bot(i).Visible = False
    End If
Next i

' Register1_port
Call Air_VV_position(440, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        Register1_port_top(i).Visible = True
        Register1_port_bot(i).Visible = True
    Else
        Register1_port_top(i).Visible = False
        Register1_port_bot(i).Visible = False
    End If
Next i

' Register2_port
Call Air_VV_position(441, VV_pos)
For i = 0 To 5
    If i = VV_pos Then
        Register2_port_top(i).Visible = True
        Register2_port_bot(i).Visible = True
    Else
        Register2_port_top(i).Visible = False
        Register2_port_bot(i).Visible = False
    End If
Next i



' TESTING:
'Text1.Text = LSS_reals(219)


' Cross-over Damper
If LSS_reals(438) > 11# Then

    VV_pos = 6
    
    Xover_damper(0).BackColor = &HC0C0C0        ' GREY
    Xover_damper(1).BackColor = &HC0C0C0        ' GREY
    Xover_damper(2).BackColor = &HFF00&         ' GREEN  ( PORT )

ElseIf LSS_reals(438) > 9# Then
    VV_pos = 5
ElseIf LSS_reals(438) > 7# Then
    VV_pos = 4
ElseIf LSS_reals(438) > 5# Then
    VV_pos = 3
    
    Xover_damper(0).BackColor = &HFF00&         ' WHITE  ( BOTH )
    Xover_damper(1).BackColor = &HC0C0C0        ' GREY
    Xover_damper(2).BackColor = &HC0C0C0        ' GREY

ElseIf LSS_reals(438) > 3# Then
    VV_pos = 2
ElseIf LSS_reals(438) > 1# Then
    VV_pos = 1
Else
    VV_pos = 0
    
    Xover_damper(0).BackColor = &HC0C0C0        ' GREY
    Xover_damper(1).BackColor = &HFF00&         ' GREEN  ( STBD )
    Xover_damper(2).BackColor = &HC0C0C0        ' GREY

End If

For i = 0 To 6
    If i = VV_pos Then
        Cross_over(i).Visible = True
    Else
        Cross_over(i).Visible = False
    End If
Next i

' === 7/12/01 ===
If LSS_reals(436) > 9.5 Then
    BYPASS_Air_heater(1).BackColor = &HFFFF&     ' YELLOW ( By-passed )
ElseIf LSS_reals(436) > 0.5 Then
    If flash_bypass Then        ' === 7/12/01 ===
        BYPASS_Air_heater(1).BackColor = &HC0C0C0         ' GREY ( NOT By-passed )
    Else
        BYPASS_Air_heater(1).BackColor = &HFFFF&     ' YELLOW ( By-passed )
    End If
Else
    BYPASS_Air_heater(1).BackColor = &HC0C0C0         ' GREY ( NOT By-passed )
End If

' Air Heater By-Pass position  STBD
If LSS_reals(436) > 8# Then
    VV_pos = 4
    
  ' === 7/12/01 14.093 ===
  ' BYPASS_Air_heater(1).BackColor = &HFF00&     ' GREEN ( By-passed )
    
ElseIf LSS_reals(436) > 6# Then
    VV_pos = 3
ElseIf LSS_reals(436) > 4# Then
    VV_pos = 2
ElseIf LSS_reals(436) > 2# Then
    VV_pos = 1
Else
    VV_pos = 0
  ' === 7/12/01 14.093 ===
  ' BYPASS_Air_heater(1).BackColor = &H8080FF         ' PINK ( NOT By-passed )
    
End If

For i = 0 To 4
    If i = VV_pos Then
        Heater_bypass_stbd(i).Visible = True
    Else
        Heater_bypass_stbd(i).Visible = False
    End If
Next i

' Air Heater By-Pass position  PORT
' === 7/12/01 ===
If LSS_reals(437) > 9.5 Then
    BYPASS_Air_heater(2).BackColor = &HFFFF&     ' YELLOW ( By-passed )
ElseIf LSS_reals(437) > 0.5 Then
    If flash_bypass Then        ' === 7/12/01 ===
        BYPASS_Air_heater(2).BackColor = &HC0C0C0         ' GREY ( NOT By-passed )
    Else
        BYPASS_Air_heater(2).BackColor = &HFFFF&     ' YELLOW ( By-passed )
    End If
Else
    BYPASS_Air_heater(2).BackColor = &HC0C0C0         ' GREY ( NOT By-passed )
End If

If LSS_reals(437) > 8# Then
    VV_pos = 4
  ' === 7/12/01 14.093 ===
  ' BYPASS_Air_heater(2).BackColor = &HFF00&     ' GREEN ( By-passed )
ElseIf LSS_reals(437) > 6# Then
    VV_pos = 3
ElseIf LSS_reals(437) > 4# Then
    VV_pos = 2
ElseIf LSS_reals(437) > 2# Then
    VV_pos = 1
Else
    VV_pos = 0
  ' === 7/12/01 14.093 ===
  ' BYPASS_Air_heater(2).BackColor = &H8080FF         ' PINK ( NOT By-passed )
End If

For i = 0 To 4
    If i = VV_pos Then
        Heater_bypass_port(i).Visible = True
    Else
        Heater_bypass_port(i).Visible = False
    End If
Next i



' Fuel Valve position STBD
Fuel_pos_stbd.text = Format(LSS_reals(193), "###0") & " %"
If LSS_reals(193) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(193) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(193) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Fuel_VV_stbd(i).Visible = True
    Else
        Fuel_VV_stbd(i).Visible = False
    End If
Next i

' Fuel Valve position PORT
Fuel_pos_port.text = Format(LSS_reals(194), "###0") & " %"
If LSS_reals(194) > 99# Then
    VV_pos = 3
ElseIf LSS_reals(194) > 49# Then
    VV_pos = 2
ElseIf LSS_reals(194) > 0.5 Then
    VV_pos = 1
Else
    VV_pos = 0
End If
For i = 0 To 3
    If i = VV_pos Then
        Fuel_VV_port(i).Visible = True
    Else
        Fuel_VV_port(i).Visible = False
    End If
Next i

' NB  VB_steam_logicals(12) = STBD FD FAN SLOW
' NB  VB_steam_logicals(13) = PORT FD FAN SLOW


' ====  FD FANS ===

'Call update_pump_LSS(Forced_Draft, 0, 38)      ' FD Fan STBD
'Call update_pump_LSS(Forced_Draft, 1, 39)      ' FD Fan PORT
'single_char = Mid$(PC_INT_STRING, pump_no + 250, 1)

' 24/9/99  FD Fan STBD = pump  5
' 24/9/99  FD Fan PORT = pump 17
'single_char = Mid$(PC_INT_STRING, 288, 1)
single_char = Mid$(PC_INT_STRING, 255, 1)

'If single_char = "0" Then
'   pp_status_LSS(38) = 0
'ElseIf single_char = "1" Then
'   pp_status_LSS(38) = 1
'ElseIf single_char = "2" Then
'   pp_status_LSS(38) = 2
'Else
'   pp_status_LSS(38) = 3
'End If
If single_char = "0" Then
   pp_status_LSS(5) = 0
ElseIf single_char = "1" Then
   pp_status_LSS(5) = 1
ElseIf single_char = "2" Then
   pp_status_LSS(5) = 2
Else
   pp_status_LSS(5) = 3
End If

'single_char = Mid$(PC_INT_STRING, 289, 1)
single_char = Mid$(PC_INT_STRING, 267, 1)

'If single_char = "0" Then
'   pp_status_LSS(39) = 0
'ElseIf single_char = "1" Then
'   pp_status_LSS(39) = 1
'ElseIf single_char = "2" Then
'   pp_status_LSS(39) = 2
'Else
'   pp_status_LSS(39) = 3
'End If
If single_char = "0" Then
   pp_status_LSS(17) = 0
ElseIf single_char = "1" Then
   pp_status_LSS(17) = 1
ElseIf single_char = "2" Then
   pp_status_LSS(17) = 2
Else
   pp_status_LSS(17) = 3
End If

'  FD Fan STBD
If pp_status_LSS(5) = 0 Then  '  FD Fan STBD is OFF
            OFFpump(0).BackColor = &HFFFFFF           ' WHITE
            ONpump(0).BackColor = &HC0C0C0            ' GREY
            SLOW_FD(0).BackColor = &HC0C0C0           ' GREY
ElseIf pp_status_LSS(5) = 3 Then  '  FD Fan STBD is ON
    If (VB_steam_logicals(12)) Then
            OFFpump(0).BackColor = &HC0C0C0           ' GREY
            ONpump(0).BackColor = &HC0C0C0            ' GREY
            SLOW_FD(0).BackColor = &HC000&            ' DARK GREEN  ( SLOW )
    Else
            OFFpump(0).BackColor = &HC0C0C0           ' GREY
            ONpump(0).BackColor = &HFF00&             ' BRIGHT GREEN ( FAST )
            SLOW_FD(0).BackColor = &HC0C0C0           ' GREY
    
    End If
End If


'  FD Fan PORT
If pp_status_LSS(17) = 0 Then  '  FD Fan PORT is OFF
            OFFpump(1).BackColor = &HFFFFFF           ' WHITE
            ONpump(1).BackColor = &HC0C0C0            ' GREY
            SLOW_FD(1).BackColor = &HC0C0C0           ' GREY
ElseIf pp_status_LSS(17) = 3 Then  '  FD Fan PORT is ON
    If (VB_steam_logicals(13)) Then
            OFFpump(1).BackColor = &HC0C0C0           ' GREY
            ONpump(1).BackColor = &HC0C0C0            ' GREY
            SLOW_FD(1).BackColor = &HC000&            ' DARK GREEN  ( SLOW )
    Else
            OFFpump(1).BackColor = &HC0C0C0           ' GREY
            ONpump(1).BackColor = &HFF00&             ' BRIGHT GREEN ( FAST )
            SLOW_FD(1).BackColor = &HC0C0C0           ' GREY
    
    End If
End If



' === ROTARY AIR HEATERS ===

' Note: update_pump_LSS(frm As Form, index_no As Integer, pump_no As Integer)
'Call update_pump_LSS(Forced_Draft, 2, 36)      ' Rotary Air Heater STBD
'Call update_pump_LSS(Forced_Draft, 3, 37)      ' Rotary Air Heater PORT
Call update_pump_LSS(Forced_Draft, 2, 31)      ' Rotary Air Heater STBD
Call update_pump_LSS(Forced_Draft, 3, 32)      ' Rotary Air Heater PORT


' Note: update_vvalve_LSS(syxxx,nValve,index1,index2)

' 24/9/99 OLD LSS
'   Call update_vvalve_LSS(Forced_Draft, 206, 0, -1)
'   Call update_vvalve_LSS(Forced_Draft, 207, 1, -1)
'   Call update_vvalve_LSS(Forced_Draft, 208, 2, -1)
'   Call update_vvalve_LSS(Forced_Draft, 215, 3, -1)
'   Call update_vvalve_LSS(Forced_Draft, 209, 4, -1)
'   Call update_vvalve_LSS(Forced_Draft, 211, 5, -1)
'   Call update_vvalve_LSS(Forced_Draft, 212, 6, -1)
'   Call update_vvalve_LSS(Forced_Draft, 213, 7, -1)
'   Call update_vvalve_LSS(Forced_Draft, 216, 8, -1)
'   Call update_vvalve_LSS(Forced_Draft, 218, 9, -1)
'   Call update_vvalve_LSS(Forced_Draft, 217, 10, -1)
'   Call update_vvalve_LSS(Forced_Draft, 210, 11, -1)
'   Call update_vvalve_LSS(Forced_Draft, 219, 12, -1)
'   Call update_vvalve_LSS(Forced_Draft, 214, 13, -1)
'   Call update_vvalve_LSS(Forced_Draft, 220, 14, -1)

   Call update_vvalve_LSS(Forced_Draft, 200, 0, -1)
   Call update_vvalve_LSS(Forced_Draft, 201, 1, -1)
   Call update_vvalve_LSS(Forced_Draft, 207, 2, -1)
   Call update_vvalve_LSS(Forced_Draft, 217, 3, -1)
   Call update_vvalve_LSS(Forced_Draft, 216, 4, -1)
   Call update_vvalve_LSS(Forced_Draft, 202, 5, -1)
   Call update_vvalve_LSS(Forced_Draft, 205, 6, -1)
   Call update_vvalve_LSS(Forced_Draft, 203, 7, -1)
   Call update_vvalve_LSS(Forced_Draft, 211, 8, -1)
   Call update_vvalve_LSS(Forced_Draft, 212, 9, -1)
   Call update_vvalve_LSS(Forced_Draft, 209, 10, -1)
   Call update_vvalve_LSS(Forced_Draft, 204, 11, -1)
   Call update_vvalve_LSS(Forced_Draft, 210, 12, -1)
   Call update_vvalve_LSS(Forced_Draft, 215, 13, -1)
   Call update_vvalve_LSS(Forced_Draft, 208, 14, -1)


End Sub

Private Sub valve_Click(Index As Integer)

   
Control_PC_check
If This_Control_status = False Then Exit Sub


' === 09/02/2010 ===
If Not I_AM_INSTRUCTOR And Not PTT Then

    ' I MUST be an ACT Terminal (Full Mission) so

    ' HFO Root Valves
    If Index = 8 Or Index = 9 Or Index = 11 Or Index = 6 Then
        Exit Sub
    End If

End If



   
' 30/11/99  If Full Mission then do NOT allow clicks on MOTV and BOTV
If (Not PTT) Then
    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 7 Or Index = 10 Or Index = 12 Then
        Exit Sub
    End If
End If
   
   
Select Case Index
   
  Case 0
   Call valve_pcdi_LSS(200)
  Case 1
   Call valve_pcdi_LSS(201)
  Case 2
   Call valve_pcdi_LSS(207)
  Case 3
   Call valve_pcdi_LSS(217)
  Case 4
   Call valve_pcdi_LSS(216)
  Case 5
   Call valve_pcdi_LSS(202)
  Case 6
   Call valve_pcdi_LSS(205)
  Case 7
   Call valve_pcdi_LSS(203)
  Case 8
   Call valve_pcdi_LSS(211)
  Case 9
   Call valve_pcdi_LSS(212)
  Case 10
   Call valve_pcdi_LSS(209)
  Case 11
   Call valve_pcdi_LSS(204)
  Case 12
   Call valve_pcdi_LSS(210)
  Case 13
   Call valve_pcdi_LSS(215)
  Case 14
   Call valve_pcdi_LSS(208)
   
End Select

End Sub

Private Sub VScroll1_Change(Index As Integer)

    Call set_VSCROLL(Index + 11, VScroll1(Index).Value)   ' Air register position

End Sub

Private Sub Xover_damper_Click(Index As Integer)

set_dummy_focus = True

' OLD
'If Index = 0 Then
'    status_flag = False
'    Call set_STEAM_logicals(10, status_flag)  ' Xover to BOTH
'    Call set_STEAM_logicals(11, status_flag)
'ElseIf Index = 1 Then
'    status_flag = True
'    Call set_STEAM_logicals(10, status_flag)
'    status_flag = False
'    Call set_STEAM_logicals(11, status_flag)
'Else
'    status_flag = False
'    Call set_STEAM_logicals(10, status_flag)
'    status_flag = True
'    Call set_STEAM_logicals(11, status_flag)
'End If

' 26/9/99
If Index = 0 Then
  Call set_PCDI_LSS(108)  ' BOTH
ElseIf Index = 1 Then
  Call set_PCDI_LSS(107)  ' STBD
ElseIf Index = 2 Then
  Call set_PCDI_LSS(109)  ' PORT
End If



End Sub
Sub Air_VV_position(VV_number As Integer, VV_posn As Integer)

If LSS_reals(VV_number) > 95# Then
    VV_posn = 5
ElseIf LSS_reals(VV_number) > 75# Then
    VV_posn = 4
ElseIf LSS_reals(VV_number) > 50# Then
    VV_posn = 3
ElseIf LSS_reals(VV_number) > 25# Then
    VV_posn = 2
ElseIf LSS_reals(VV_number) > 5# Then
    VV_posn = 1
Else
    VV_posn = 0
End If

End Sub

