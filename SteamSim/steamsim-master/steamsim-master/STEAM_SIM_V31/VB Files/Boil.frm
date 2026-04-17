VERSION 5.00
Begin VB.Form boil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OFB CONTROLS"
   ClientHeight    =   4245
   ClientLeft      =   3945
   ClientTop       =   4845
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "ov015"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   Visible         =   0   'False
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4020
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "boil.frx":0000
      Top             =   3240
      Width           =   1635
   End
   Begin VB.CommandButton OFFpump 
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
      Height          =   375
      Index           =   1
      Left            =   3120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3420
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton ONpump 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MANUAL"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2220
      UseMaskColor    =   -1  'True
      Width           =   795
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
      Height          =   375
      Index           =   0
      Left            =   4140
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton ONpump 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Index           =   0
      Left            =   4140
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1860
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2640
      Top             =   180
   End
   Begin VB.TextBox dummybox 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00800000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   855
      Left            =   2520
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H00800000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   1575
      Left            =   2520
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2220
      Width           =   375
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   375
   End
   Begin VB.Shape Shape14 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   1860
      Width           =   375
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   375
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00800000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   1875
      Left            =   300
      Top             =   2100
      Width           =   2055
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HIGH STEAM PRESS"
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
      TabIndex        =   9
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOW STEAM PRESS"
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
      TabIndex        =   8
      Top             =   3540
      Width           =   1350
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   3180
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   2580
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FLAME FAILURE"
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
      TabIndex        =   7
      Top             =   2940
      Width           =   1110
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HIGH WATER LEVEL"
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
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOW WATER LEVEL"
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
      TabIndex        =   5
      Top             =   2340
      Width           =   1290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      X1              =   304
      X2              =   304
      Y1              =   148
      Y2              =   172
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      X1              =   260
      X2              =   304
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PURGE"
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
      Left            =   660
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No.1 IGNITION"
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
      Left            =   660
      TabIndex        =   3
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No.1 FLAME"
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
      Left            =   660
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   300
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   600
      Width           =   375
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   900
      Width           =   375
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No.2 FLAME"
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
      Left            =   660
      TabIndex        =   1
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No.2 IGNITION"
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
      Left            =   660
      TabIndex        =   0
      Top             =   1260
      Width           =   930
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00800000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   1875
      Left            =   300
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "boil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim set_dummy_focus As Integer




Private Sub OFFpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
Select Case Index
    Case 0     ' Boiler Manual stop
        Call setPCDI(1844)
    Case 1     ' Boiler AUTO
        Call setPCDI(1845)
End Select

End Sub

Private Sub ONpump_Click(Index As Integer)

set_dummy_focus = True

Control_PC_check
If This_Control_status = False Then Exit Sub
   
Select Case Index
    Case 0     ' Boiler Manual start
        Call setPCDI(1838)
    Case 1     ' Boiler Manual
        Call setPCDI(1839)
End Select

End Sub

Private Sub Timer1_Timer()
'Dim single_char As String * 1

If set_dummy_focus Then
   set_dummy_focus = False
   dummybox.SetFocus
End If

'12/9/96:
flash_status = Not flash_status

' =======================================================
'start/stop the boiler

' 12/9/96 This is the Boiler ON lamp:
'Call get_ch(392, 2975)

' 12/9/96 We also need the Boiler ON FLASH:
'Call get_ch(309, 2974)

' 12/9/96 We also need the Boiler STOP FLASH:
'Call get_ch(310, 2946)

' 12/9/96 :  NEW Boiler Start/Stop logic:

'   If ch(2975) = True Then  ' Must be ON ( steady GREEN )
    If Mid$(PC_INT_STRING, 392, 1) = "1" Then
       Shape14.FillColor = &HFF00&     ' GREEN
       Shape12.FillColor = &HFFFFFF    ' WHITE
'   ElseIf ch(2974) = True Then                    ' Must FLASH GREEN START:
    ElseIf Mid$(PC_INT_STRING, 309, 1) = "1" Then  ' Must FLASH GREEN START:
       If flash_status Then
         Shape14.FillColor = &HFF00&     ' GREEN
       Else
         Shape14.FillColor = &HFFFFFF    ' WHITE
       End If
         Shape12.FillColor = &HFFFFFF   ' Steady WHITE
'   ElseIf ch(2946) = True Then                        ' Must FLASH STOP WHITE:
    ElseIf Mid$(PC_INT_STRING, 309, 1) = "1" Then      ' Must FLASH STOP WHITE:
     If flash_status Then
       Shape12.FillColor = &HFF00&     ' GREEN
     Else
       Shape12.FillColor = &HFFFFFF    ' WHITE
     End If
     Shape14.FillColor = &HFFFFFF      ' WHITE
   Else     ' Must be totally Stopped:
     Shape12.FillColor = &HFF00&     ' GREEN
     Shape14.FillColor = &HFFFFFF    ' WHITE
   End If




'get purge info
'    Call get_ch(366, 3031)
' purge
' j = 3031
'  If ch(j) = False Then
   If Mid$(PC_INT_STRING, 366, 1) = "0" Then
     Shape3.FillColor = &HFFFFFF
   Else
     Shape3.FillColor = &HFF00&
   End If
   

'no 1 ignition
'    Call get_ch(367, 3049)
' burn1 ign
' j = 3049
'  If ch(j) = False Then
   If Mid$(PC_INT_STRING, 367, 1) = "0" Then
     Shape4.FillColor = &HFFFFFF
   Else
     Shape4.FillColor = &HFF00&
   End If
   
   
'no 1 flame
'Call get_ch(368, 3051)
' burn1 flame
' j = 3051
'  If ch(j) = False Then
   If Mid$(PC_INT_STRING, 368, 1) = "0" Then
     Shape7.FillColor = &HFFFFFF
   Else
     Shape7.FillColor = &HFF00&
   End If
 
'no 2 ignition
'Call get_ch(369, 3047)
'1584 burn2 ign
' j = 3047
'   If ch(j) = False Then
    If Mid$(PC_INT_STRING, 369, 1) = "0" Then
     Shape8.FillColor = &HFFFFFF
   Else
     Shape8.FillColor = &HFF00&
   End If


'no 2 flame
'Call get_ch(370, 3033)
' burn2 flame
' j = 3033
   If Mid$(PC_INT_STRING, 370, 1) = "0" Then
     Shape9.FillColor = &HFFFFFF
   Else
     Shape9.FillColor = &HFF00&
   End If

 


'low water level
'Call get_ch(386, 3055)
'low water
 'j = 3055
    If Mid$(PC_INT_STRING, 386, 1) = "0" Then
     Shape1.FillColor = &HFFFFFF
   Else
     Shape1.FillColor = &HFF&
   End If


'hi water level
'Call get_ch(387, 3043)
'high water
' j = 3043
   If Mid$(PC_INT_STRING, 387, 1) = "0" Then
     Shape2.FillColor = &HFFFFFF
   Else
     Shape2.FillColor = &HFF&
   End If

'flame fail
'Call get_ch(388, 3037)
'flame fail
' j = 3037
   If Mid$(PC_INT_STRING, 388, 1) = "0" Then
     Shape5.FillColor = &HFFFFFF
   Else
     Shape5.FillColor = &HFF&
   End If

'hi steam pressure
'Call get_ch(389, 3025)
'high steam press
' j = 3025
   If Mid$(PC_INT_STRING, 389, 1) = "0" Then
     Shape6.FillColor = &HFFFFFF
   Else
     Shape6.FillColor = &HFF&
   End If

'lo steam pressure
'Call get_ch(390, 3045)
'low steam press
' j = 3045
   If Mid$(PC_INT_STRING, 390, 1) = "0" Then
     Shape10.FillColor = &HFFFFFF
   Else
     Shape10.FillColor = &HFF&
   End If

 


'status mam/auto
'Call get_ch(391, 2951)
'man/auto
' j = 2951
' 8/7/98  WRONG...
'  If Mid$(PC_INT_STRING, 390, 1) = "0" Then
   If Mid$(PC_INT_STRING, 391, 1) = "0" Then
     Shape16.FillColor = &HFFFFFF        ' MAN  WHITE
     Shape15.FillColor = &HFF00&         ' AUTO GREEN
   Else
     Shape16.FillColor = &HFF00&         ' MAN  GREEN
     Shape15.FillColor = &HFFFFFF        ' AUTO WHITE
   End If



End Sub

