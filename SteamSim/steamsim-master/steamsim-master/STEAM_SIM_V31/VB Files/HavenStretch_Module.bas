Attribute VB_Name = "HavenStretch_Module"
Sub copy_havenstretch_colours(iform As Form, B_colour As Long)


Dim my_obj As Object
Dim sObjName As String


' === 27/05/02 ===
Dim i As Integer

'For Each my_obj In Me
For Each my_obj In iform

        sObjName = TypeName(my_obj)
        
        Select Case sObjName
        
            Case Is = "Line"
            
            Case Is = "HavenValveV1_2"
                my_obj.BackColour = B_colour
                
            Case Is = "Haven3WayValve"
                my_obj.BackColour = B_colour
                
            Case Is = "Pump"
                my_obj.Back_color = B_colour
                
            Case Is = "Label"
               If my_obj.BackColor = &H8000000F Or my_obj.BackColor = &HC0C0C0 Or my_obj.BackColor = vbWhite Then
                   ' The label is GREY or grey or WHITE so switch colours....
                    my_obj.BackColor = B_colour
                End If
                
            Case Is = "Frame"
                my_obj.BackColor = B_colour
                                                 
        End Select
Next

' == Special cases fOR the Steam Simulator ===

' === 30/05/02 ===
If iform.Name = "Ballast_BULK" Then
    For i = 1 To 13
        iform.Valve(i).BackColour = &H808080             ' Dark Grey to match the Manifolds
    Next i
        iform.Valve(30).BackColour = &H808080             ' Dark Grey to match the Manifolds
End If


If iform.Name = "Bunkers" Then
        iform.Thick_line(0).BorderColor = B_colour
        iform.Thick_line(1).BorderColor = B_colour
End If
If iform.Name = "Stbd_Boiler" Then
    iform.Valve(443).BackColour = &HC0FFFF           ' Light YELLOW to match the Boiler
    iform.Valve(437).BackColour = &HC0FFFF           ' Light YELLOW to match the Boiler
    iform.Valve(433).BackColour = &HC0FFFF           ' Light YELLOW to match the Boiler
End If
If iform.Name = "Port_Boiler" Then
    iform.Valve(473).BackColour = &HC0FFFF           ' Light YELLOW to match the Boiler
    iform.Valve(467).BackColour = &HC0FFFF           ' Light YELLOW to match the Boiler
    iform.Valve(463).BackColour = &HC0FFFF           ' Light YELLOW to match the Boiler
End If
If iform.Name = "Condensate" Then
    iform.Loop_Seal_shape.FillColor = B_colour
End If
If iform.Name = "Gland_Steam_Full_away" Then
    iform.Loop_Seal_shape.FillColor = B_colour
End If
If iform.Name = "Gland_Steam" Then
    iform.Loop_Seal_shape.FillColor = B_colour
End If
If iform.Name = "Mollier" Then
    If B_colour = vbWhite Then
        iform.Graph1.Background = 15            ' WHITE
        iform.Graph1.OverlayColor = 0           ' BLACK
    Else
        iform.Graph1.Background = 7             ' GREY
        iform.Graph1.OverlayColor = 15          ' WHITE
    End If
End If

End Sub
