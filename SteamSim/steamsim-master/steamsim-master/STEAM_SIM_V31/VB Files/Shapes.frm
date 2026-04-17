VERSION 5.00
Object = "{030BE89F-5848-11D2-B1D2-00C04FA9737B}#1.0#0"; "HavenStretch.ocx"
Begin VB.Form Shapes 
   Caption         =   "Shapes"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin Stretch_Control.HavenStretch HavenStretch1 
      Height          =   6075
      Left            =   450
      Top             =   270
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   10716
      Begin VB.OLE OLE5 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   1245
         Left            =   3690
         OleObjectBlob   =   "Shapes.frx":0000
         TabIndex        =   4
         Top             =   3720
         Width           =   1455
      End
      Begin VB.OLE OLE4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   1035
         Left            =   960
         OleObjectBlob   =   "Shapes.frx":1C18
         SizeMode        =   1  'Stretch
         TabIndex        =   3
         Top             =   3750
         Width           =   1035
      End
      Begin VB.OLE OLE3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   885
         Left            =   5130
         OleObjectBlob   =   "Shapes.frx":3830
         SizeMode        =   1  'Stretch
         TabIndex        =   2
         Top             =   780
         Width           =   1155
      End
      Begin VB.OLE OLE2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Enabled         =   0   'False
         Height          =   945
         Left            =   3480
         OleObjectBlob   =   "Shapes.frx":5448
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\Turbineup.vsd"
         TabIndex        =   1
         Top             =   690
         Width           =   855
      End
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Visio.Drawing.5"
         Height          =   2295
         Left            =   780
         OleObjectBlob   =   "Shapes.frx":7060
         SizeMode        =   1  'Stretch
         SourceDoc       =   "D:\OIL_WATER_SEP.vsd"
         TabIndex        =   0
         Top             =   810
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Shapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
