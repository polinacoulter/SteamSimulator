VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_password_change 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2115
   ClientLeft      =   4665
   ClientTop       =   2595
   ClientWidth     =   3045
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
   Icon            =   "password_change.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2115
   ScaleWidth      =   3045
   Begin VB.TextBox text_new_password 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox text_old_password 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin Threed.SSCommand Command3D1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Cancel"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin Threed.SSCommand command_ok 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "OK"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "No Original Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter new password:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter old password:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_password_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim password_changed As Integer 'Boolean

Private Sub command_ok_Click()

'21/11/97
Label1.Enabled = True
text_old_password.Enabled = True

'Debug.Print "old=" & text_old_password.Text; "new=" & instructor_password
        
    If text_old_password.Text = instructor_password Then
    
'       Debug.Print "Write password"
       
        password_changed = True
        If instructor_password <> text_new_password Then
            instructor_password = text_new_password
            MsgBox "Password Changed"
        End If
        
'30/12/97
        PasswordSet = True
        
'21/11/97 added:
         If useEnglish Then
              Open App.Path & "\Epassword.txt" For Output As #1
         Else
              Open App.Path & "\Ipassword.txt" For Output As #1
         End If
                 Write #1, instructor_password
                 Close #1
        
        
        Unload Me
        
    Else
        MsgBox "Original password Incorrect"
        password_changed = False
        text_old_password.Text = ""
        text_old_password.SetFocus

    End If
    
End Sub

Private Sub Command3D1_Click()
'    new_password = ""
    password_changed = False
    Unload Me
End Sub

Private Sub Form_Load()

    If instructor_password <> "" Then
        text_old_password.Enabled = True
'21/11/97
        Label2.Enabled = False
    Else
        text_old_password.Enabled = False
        text_old_password.Visible = False
        Label1.Visible = False
        Label3.Visible = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'21/11/97 moved to OK  button...
'    If password_changed Then
'        If instructor_password <> text_new_password Then
'            instructor_password = text_new_password
''21/11/97
''            If Len(instructor_password) > 30 Then
''                MsgBox "Password Limited to first 30 characters, i.e. 'abcdefghijklmnopqrstuvwxyzabcdefgh' becomes 'abcdefghijklmnopqrstuvwxyzabcd'"
''                instructor_password = pwd_limit_to_30chars(instructor_password)
''            End If
'
'            MsgBox "Password Changed"
'        End If
'    End If
    
End Sub

Private Sub text_new_password_LostFocus()
    'make it all uppercase
    text_new_password.Text = UCase$(text_new_password.Text)

End Sub

Private Sub text_old_password_Change()

'21/11/97
'Debug.Print "old password = " & UCase(text_old_password.Text) & "password = " & instructor_password
If UCase(text_old_password.Text) = instructor_password Then
        Label2.Enabled = True
        Label1.Enabled = False
        text_old_password.Enabled = False
End If

End Sub

Private Sub text_old_password_LostFocus()
    'make it all uppercase
    text_old_password.Text = UCase$(text_old_password.Text)

End Sub

