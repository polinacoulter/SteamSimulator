VERSION 5.00
Begin VB.Form frm_config 
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Configuration"
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   6675
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   3720
         Top             =   5700
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   900
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   5100
         TabIndex        =   14
         Top             =   5640
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   435
         Left            =   1260
         TabIndex        =   13
         Top             =   5640
         Width           =   1275
      End
      Begin VB.Shape OnLineBox 
         BorderColor     =   &H000000FF&
         Height          =   675
         Left            =   1200
         Top             =   660
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   495
         Index           =   5
         Left            =   5820
         Top             =   4980
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   495
         Index           =   4
         Left            =   4920
         Top             =   4980
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   495
         Index           =   3
         Left            =   4020
         Top             =   4980
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   495
         Index           =   2
         Left            =   3120
         Top             =   4980
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   495
         Index           =   1
         Left            =   2220
         Top             =   4980
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   495
         Index           =   0
         Left            =   1320
         Top             =   4980
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   5
         Left            =   5820
         Top             =   4140
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   4
         Left            =   4920
         Top             =   4140
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   3
         Left            =   4020
         Top             =   4140
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   2
         Left            =   3120
         Top             =   4140
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   1
         Left            =   2220
         Top             =   4140
         Width           =   555
      End
      Begin VB.Image Image5 
         Height          =   495
         Index           =   0
         Left            =   1320
         Top             =   4140
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   5
         Left            =   5820
         Top             =   3300
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   4
         Left            =   4920
         Top             =   3300
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   3
         Left            =   4020
         Top             =   3300
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   2
         Left            =   3120
         Top             =   3300
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   1
         Left            =   2220
         Top             =   3300
         Width           =   555
      End
      Begin VB.Image Image4 
         Height          =   495
         Index           =   0
         Left            =   1320
         Top             =   3300
         Width           =   555
      End
      Begin VB.Label label2 
         Caption         =   "On Line"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   5
         Left            =   5820
         Top             =   2400
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   4
         Left            =   4920
         Top             =   2400
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   3
         Left            =   4020
         Top             =   2400
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   2
         Left            =   3120
         Top             =   2400
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   1
         Left            =   2220
         Top             =   2400
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   495
         Index           =   0
         Left            =   1320
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Group 6"
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   4920
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Group 5"
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   5
         Left            =   5820
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   4
         Left            =   4920
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   3
         Left            =   4020
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   2
         Left            =   3120
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   1
         Left            =   2220
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   0
         Left            =   1320
         Top             =   1560
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   5
         Left            =   5820
         Top             =   780
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   4
         Left            =   4920
         Top             =   780
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   3
         Left            =   4020
         Top             =   780
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   2
         Left            =   3120
         Top             =   780
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2220
         Picture         =   "Config.frx":0000
         Top             =   780
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1320
         Picture         =   "Config.frx":030A
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Group 4"
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Group 3"
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Group 2"
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Group 1"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label label2 
         Caption         =   "PC05"
         Height          =   315
         Index           =   6
         Left            =   5820
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label label2 
         Caption         =   "PC04"
         Height          =   315
         Index           =   5
         Left            =   4920
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label label2 
         Caption         =   "PC03"
         Height          =   315
         Index           =   4
         Left            =   4020
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label label2 
         Caption         =   "PC02"
         Height          =   315
         Index           =   3
         Left            =   3120
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label label2 
         Caption         =   "PC01"
         Height          =   315
         Index           =   2
         Left            =   2220
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label label2 
         Caption         =   "PC00"
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   6735
      Begin VB.Label Label1 
         Caption         =   "Configure Operation of each computer on the simulator network. "
         Height          =   675
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6675
      End
   End
End
Attribute VB_Name = "frm_config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DrawOnLineBox()
'draws a red box around all PCs in group 1 if they are on-line


Dim pc_count As Integer
Dim i As Integer

If Check1.Value = 1 Then
        OnLineBox.Left = 1200 ' box always starts here to left of PC00
                                'find out where the last Pc is in group 1
    
        pc_count = 0
        For i = 0 To Number_of_PCs - 1
            If PC_Group(i) = 1 Then
                pc_count = i
            End If
        Next
        'so if the last pc is number n,  Width = Coords of RHS last image - Coords of LHS First image
        OnLineBox.Width = (Image1(pc_count).Left + Image1(pc_count).Width - Image1(0).Left) + (Image1(pc_count).Width / 2)
        OnLineBox.Visible = True
    Else
        
        OnLineBox.Visible = False
    End If
End Sub
Sub DragDrop(Index As Integer, Source As Control, State As Integer)
    If Index = Source.Index Then

        Select Case State
            Case vbEnter
            ' Load icon.
                Source.DragIcon = LoadPicture("ICONS\ARROWS\POINT03.ICO")
         
            Case vbLeave
                Source.DragIcon = LoadPicture() ' Unload icon.
            End Select
        End If
End Sub




Private Sub Check1_Click()
Dim i As Integer
Dim pc_count As Integer

    
        
End Sub

Private Sub Image1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Index = Source.Index And Source.Top <> Image1(Index).Top Then
 
    Image1(Index).Picture = Source.Picture
    
    Source.Picture = LoadPicture() ' unload the icon from its last position
End If

End Sub

Private Sub Image1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If Index = Source.Index Then

    Select Case State
        Case vbEnter
            ' Load icon.
            Source.DragIcon = LoadPicture(App.Path & "\images\DownPointer.ico")
          
        Case vbLeave
            Source.DragIcon = LoadPicture() ' Unload icon.
    End Select
End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Response As Integer

'20/2/98 If Index = 0 Then
If Index = 0 And OffLineAllowed Then        'we cant put pc00 in another group 'cause its got the hw interface

      Response = MsgBox("You cannot move PC00 because it must be in group 1", vbOKOnly + vbExclamation, "Configuration")
End If






End Sub

Private Sub Image3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Index = Source.Index And Source.Top <> Image3(Index).Top Then

        Image3(Index).Picture = Source.Picture
        Source.Picture = LoadPicture() ' unload the icon from its last position
    End If
End Sub

Private Sub Image3_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)


If Index = Source.Index Then

    Select Case State
        Case vbEnter
            ' Load icon.
            Source.DragIcon = LoadPicture(App.Path & "\images\DownPointer.ico")
         
        Case vbLeave
            Source.DragIcon = LoadPicture() ' Unload icon.
    End Select
End If
End Sub






Private Sub Command2_Click()
  
    
    For i = 0 To 5
        Debug.Print " PC_Group(" & i & ") = " & PC_Group(i)

    Next
    
    For i = 0 To 5
        Debug.Print " my_host_id(" & i & ") = " & my_host_id(i)

    Next
 
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'TEST ONLY !!!!
    'Number_of_PCs = 6 'test only

    
    'set up icons for all PCs and put them all in group 1 by default
    'set up empty icons (images) for all possible groups
    'the number of possible groups will = the number of Pcs as they could all work stand alone
    
    'Show a label(column)  and group (row) for each PC up to maximum number allowed
    'Don't show any more groups if no off-line operation allowed
    
    For i = 1 To Number_of_PCs
        Label2(i).Visible = True
        If OffLineAllowed Then
            Label3(i - 1).Visible = True
        End If
    Next
    
    'if Off Line Operation is not allowed, then on line must be allowed so
    'tick OnLinecheck box and disable
    
    If OffLineAllowed = False Then
        Check1.Value = 1
        Check1.Enabled = False
    Else
        Check1.Value = OldOnLineStatus
    End If
    
    'If OnLine Operation not allowed, clear checkbox and make invisible
    If OnLineAllowed = False Then
        Check1.Value = 0
        Check1.Visible = False
        Label2(0).Visible = False
    Else
        Check1.Value = OldOnLineStatus
        Check1.Visible = True
        Label2(0).Visible = True
    End If

    
    
    For i = 0 To Number_of_PCs - 1
    
        'PC00 is an exception, this must be in group 1 - so don't allow drag+drop
        If i = 0 Then
            Image1(i).DragMode = 0
        Else
            If OffLineAllowed Then  'allow image to be dragged only if offline available
                Image1(i).DragMode = 1
            Else
                Image1(i).DragMode = 0
            End If
        End If
        
        Image1(i).Picture = LoadPicture(App.Path & "\images\computer.ico")
       
        If Number_of_PCs >= 2 Then 'allow 2 groups
            'Image2(i).DragMode = 1
            If OffLineAllowed Then  'allow image to be dragged only if offline available
                Image2(i).DragMode = 1
            Else
                Image2(i).DragMode = 0
            End If

            Image2(i).Picture = LoadPicture()
        Else
            Image2(i).Visible = False
        End If
        
        If Number_of_PCs >= 3 Then 'allow 3 groups
            'Image3(i).DragMode = 1
            If OffLineAllowed Then  'allow image to be dragged only if offline available
                Image3(i).DragMode = 1
            Else
                Image3(i).DragMode = 0
            End If

            Image3(i).Picture = LoadPicture()
        Else
            Image3(i).Visible = False
        End If
        
        If Number_of_PCs >= 4 Then 'allow 4 groups
            'Image4(i).DragMode = 1
            If OffLineAllowed Then  'allow image to be dragged only if offline available
                Image4(i).DragMode = 1
            Else
                Image4(i).DragMode = 0
            End If

            Image4(i).Picture = LoadPicture()
        Else
            Image4(i).Visible = False
        End If
       
        If Number_of_PCs >= 5 Then 'allow 5 groups
            'Image5(i).DragMode = 1
            If OffLineAllowed Then  'allow image to be dragged only if offline available
                Image5(i).DragMode = 1
            Else
                Image5(i).DragMode = 0
            End If

            Image5(i).Picture = LoadPicture()
        Else
            Image5(i).Visible = False
        End If
        
        If Number_of_PCs >= 6 Then 'allow 6 groups
            'Image6(i).DragMode = 1
            If OffLineAllowed Then  'allow image to be dragged only if offline available
                Image6(i).DragMode = 1
            Else
                Image6(i).DragMode = 0
            End If

            Image6(i).Picture = LoadPicture()
        Else
            Image6(i).Visible = False
        End If

    Next
    
    'set the online box invisible by default
        OnLineBox.Visible = False


    
    
End Sub


Private Sub Image2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

'If Index = Source.Index Then
If Index = Source.Index And Source.Top <> Image2(Index).Top Then

    Image2(Index).Picture = Source.Picture
    'If ImageBeenDragged Then
        Source.Picture = LoadPicture() ' unload the icon from its last position
    'End If
End If
'ImageBeenDragged = False
    
End Sub

Private Sub Image2_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    'Debug.Print "in image2_dragover, index = " & Index & " x = " & X & " y = " & Y
If Index = Source.Index Then

    Select Case State
        Case vbEnter
    ' Load icon.
            Source.DragIcon = LoadPicture(App.Path & "\images\DownPointer.ico")
        Case vbLeave
            Source.DragIcon = LoadPicture() ' Unload icon.
    End Select
End If
    
End Sub

Private Sub Image4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Index = Source.Index And Source.Top <> Image4(Index).Top Then

        Image4(Index).Picture = Source.Picture
        Source.Picture = LoadPicture() ' unload the icon from its last position
    End If

End Sub

Private Sub Image4_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If Index = Source.Index Then

    Select Case State
        Case vbEnter
            ' Load icon.
            Source.DragIcon = LoadPicture(App.Path & "\images\DownPointer.ico")
         
        Case vbLeave
            Source.DragIcon = LoadPicture() ' Unload icon.
    End Select
End If
End Sub

Private Sub Image5_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Index = Source.Index And Source.Top <> Image5(Index).Top Then

        Image5(Index).Picture = Source.Picture
        Source.Picture = LoadPicture() ' unload the icon from its last position
    End If

End Sub

Private Sub Image5_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If Index = Source.Index Then

    Select Case State
        Case vbEnter
            ' Load icon.
            Source.DragIcon = LoadPicture(App.Path & "\images\DownPointer.ico")
         
        Case vbLeave
            Source.DragIcon = LoadPicture() ' Unload icon.
    End Select
End If
End Sub

Private Sub Image6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Index = Source.Index And Source.Top <> Image6(Index).Top Then

        Image6(Index).Picture = Source.Picture
        Source.Picture = LoadPicture() ' unload the icon from its last position
    End If

End Sub

Private Sub Image6_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If Index = Source.Index Then

    Select Case State
        Case vbEnter
            ' Load icon.
            Source.DragIcon = LoadPicture(App.Path & "\images\DownPointer.ico")
         
        Case vbLeave
            Source.DragIcon = LoadPicture() ' Unload icon.
    End Select
End If
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim j As Integer
Dim HostAlreadySet(0 To 31) As Byte
Dim HostForGroup(0 To 31) As Integer

Dim pc_count As Integer
Dim NewOnLineCount As Integer

'Dim PreviousPCNumber As Integer
  

'scan through all PCs and set the group for each
'eg if PC # 3 is in group 2, then PC_Group(3) = 2

    For i = 0 To Number_of_PCs - 1
    
        If Image1(i).Picture <> LoadPicture() Then
            PC_Group(i) = 1
        
        ElseIf Image2(i).Picture <> LoadPicture() Then
            PC_Group(i) = 2
        
        ElseIf Image3(i).Picture <> LoadPicture() Then
            PC_Group(i) = 3
            
        ElseIf Image4(i).Picture <> LoadPicture() Then
            PC_Group(i) = 4
        
        ElseIf Image6(i).Picture <> LoadPicture() Then
            PC_Group(i) = 5
        
        ElseIf Image6(i).Picture <> LoadPicture() Then
            PC_Group(i) = 6
           

        End If
    Next
    

'now fill the my_host_id array by selecting the lowest numbered PC
'in each group to be the host.
'Note that PC00 is always a host and has the simulator interface
'eg
'    PC     Group   my_host_id
'     0         1       0
'     1         3       1
'     2         5       2
'     3         1       0
'     4         2       4
'     5         2       4

'Clear the flags

    For i = 0 To Number_of_PCs - 1
    
        HostAlreadySet(PC_Group(i)) = False
    Next

    
 'now set the groups
    For i = 0 To Number_of_PCs - 1
    
        
        If Not HostAlreadySet(PC_Group(i)) Then 'this is first pc in the group
            my_host_id(i) = i
            HostAlreadySet(PC_Group(i)) = True
            HostForGroup(PC_Group(i)) = i 'remember the host for this group
        Else 'this is a subsequent pc in a group
            my_host_id(i) = HostForGroup(PC_Group(i))
        End If
     
      
    Next

'if number of PCs in group 0 has changed this pass, redraw the on-line box
'first find  the number of Pcs in group 0
NewOnLineCount = 0
For i = 0 To Number_of_PCs - 1
    If my_host_id(i) = 0 Then
        NewOnLineCount = NewOnLineCount + 1
    End If
Next
If NewOnLineCount <> OldOnLineCount Then
    OldOnLineCount = NewOnLineCount
    Call DrawOnLineBox
End If


'show a ref box around group 1 if it is on-line to simulator
'dont show a box if off-line
NewOnLineStatus = Check1.Value
If NewOnLineStatus <> OldOnLineStatus Then
    OldOnLineStatus = NewOnLineStatus
    Call DrawOnLineBox
End If

End Sub
