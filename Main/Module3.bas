Attribute VB_Name = "PictureEvents"
Public GridUpX As Integer
    'GridXYData(GridUpX) to first load
'Public GridDownX As Integer
    'GridXYData(GridDownX) to end load
Public XMovGrid As Integer
    'Width to pic2.line and pic1.line
Public XIndexGrid As Integer
    'Point + <-> X
Public GragX As Integer, tmpDragX As Integer
    'Point [] X to all Y
Public XPointerIndex As Long, DblXPointerIndex As Long
Public SubXPointerIndex As Integer
'____________________________________________________________________________________________________________________________
Public GridUpY As Integer
    'GridXYData(GridUpY) to first load
'Public GridDownY As Integer
    'GridXYData(GridDownY) to end load
Public YMovGrid As Integer
    'Width to pic2.line
Public YIndexGrid As Integer
    'Point + <-> Y
Public GragY As Integer, tmpDragY As Integer
    'Point [] Y to all X
Public YPointerIndex As Long, DblYPointerIndex As Long
Public SubYPointerIndex As Integer

Sub MouseX(tBGrid As BGrid, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountX As Integer)
Dim TmpCutGridX As Integer

'If Y < tBGrid.RangeY Then
    TmpCutGridX = MdlCountX
    Do
        If TmpCutGridX < tBGrid.CountX Then
            Cx = tBGrid.SelLeft(TmpCutGridX) + tBGrid.SelWidth(TmpCutGridX)
            If X < Cx Or TmpCutGridX > tBGrid.CountX - 1 Then
                If X > Cx - 5 Or X < tBGrid.SelLeft(TmpCutGridX) + 5 Then
                    If X < tBGrid.SelLeft(TmpCutGridX) + 5 Then az = 1
                        If TmpCutGridX - az < 0 Then az = 0
                        
                        ObjMe.MousePointer = 2
                        XIndexGrid = TmpCutGridX - az
                        XPointerIndex = TmpCutGridX - az
                Else
                    ObjMe.MousePointer = 0
                End If
            Exit Do
            End If
        Else
            Exit Do
        End If
        TmpCutGridX = TmpCutGridX + 1
    Loop
'Else
'    If ObjMe2.Visible = False Then ObjMe.MousePointer = 0
'End If
End Sub
Sub MouseUpX(tBGrid As BGrid, ObjMe1 As Object, ObjMe2 As Object)
Dim TmpGridWidth As Integer
    
    If ObjMe1.MousePointer = 2 Then
        ObjMe2.Visible = False
        ObjMe1.MousePointer = 0
        
        If tBGrid.SelWidth(XIndexGrid) = 0 And XMovGrid > 0 Then XIndexGrid = XIndexGrid - 1
        
        TmpGridWidth = (tBGrid.SelWidth(XIndexGrid)) + -XMovGrid
        If TmpGridWidth < 0 Then TmpGridWidth = 0
    
        If TmpGridWidth > tBGrid.GDWidth - 1 Then _
            tBGrid.SelWidth(XIndexGrid, True) = TmpGridWidth Else _
            tBGrid.SelWidth(XIndexGrid, True) = tBGrid.GDWidth
    End If
End Sub

Sub MouseY(tBGrid As BGrid, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountY As Integer)
Dim TmpCutGridY As Integer

'If X < tBGrid.RangeX Then
    TmpCutGridY = MdlCountY
    Do
        If TmpCutGridY < tBGrid.CountY Then
            Cy = tBGrid.SelTop(TmpCutGridY) + tBGrid.SelHeight(TmpCutGridY)
            If Y < Cy Or TmpCutGridY > tBGrid.CountY - 1 Then
                If Y > Cy - 5 Or Y < tBGrid.SelTop(TmpCutGridY) + 5 Then
                    If Y < tBGrid.SelTop(TmpCutGridY) + 5 Then az = 1
                        If TmpCutGridY - az < 0 Then az = 0

                        ObjMe.MousePointer = 2
                        YIndexGrid = TmpCutGridY - az
                        YPointerIndex = TmpCutGridY - az
                Else
                    ObjMe.MousePointer = 0
                End If
            Exit Do
            End If
        Else
            Exit Do
        End If
        TmpCutGridY = TmpCutGridY + 1
    Loop
'Else
'    If ObjMe2.Visible = False Then ObjMe.MousePointer = 0
'End If
End Sub
Sub MouseUpY(tBGrid As BGrid, ObjMe1 As Object, ObjMe2 As Object)
Dim TmpGridHeight As Integer
    
    If ObjMe1.MousePointer = 2 Then
        ObjMe2.Visible = False
        ObjMe1.MousePointer = 0
        
        If tBGrid.SelHeight(YIndexGrid) = 0 And YMovGrid > 0 Then YIndexGrid = YIndexGrid - 1
        
        TmpGridHeight = (tBGrid.SelHeight(YIndexGrid) - tBGrid.GDHeight) + -YMovGrid
        If TmpGridHeight < 0 Then TmpGridHeight = 0
    
        If TmpGridHeight > (tBGrid.GDHeight - tBGrid.GDHeight) - 1 Then _
            tBGrid.SelHeight(YIndexGrid, True) = TmpGridHeight Else _
            tBGrid.SelHeight(YIndexGrid, True) = tBGrid.GDHeight
    End If
End Sub

Function SizeY(tBGrid As BGrid, X As Single, Y As Single) As Boolean
    If X > tBGrid.RangePicSubX1 And X < tBGrid.RangePicSubX1 + tBGrid.RangePicSubX2 And _
    Y > tBGrid.SelTop(YPointerIndex) + tBGrid.RangePicSubY1 And Y < tBGrid.SelTop(YPointerIndex) + tBGrid.RangePicSubY1 + tBGrid.RangePicSubY2 Then
    SizeY = True
        If tBGrid.SelHeight(YPointerIndex) > tBGrid.GDHeight Then
            tBGrid.SizeUpDown YPointerIndex, False
        Else
            tBGrid.SizeUpDown YPointerIndex, True
        End If
    End If
End Function

Sub SearchPointer(tBGrid As BGrid, X As Single, TmpCutGridX As Long, Y As Single, TmpCutGridY As Long)
         'TmpCutGridX = Form1.Hsc.Value
    Do
         If X < tBGrid.SelLeft(TmpCutGridX) + tBGrid.SelWidth(TmpCutGridX) _
         Or TmpCutGridX >= tBGrid.GridDownX Then Exit Do
    TmpCutGridX = TmpCutGridX + 1
    Loop
        
        'TmpCutGridY = Form1.Vsc.Value
    'nForm(0).Text3.Text = ""
    Do
         If Y < tBGrid.SelTop(TmpCutGridY) + tBGrid.SelHeight(TmpCutGridY) _
         Or TmpCutGridY >= tBGrid.GridDownY Then Exit Do
        'nForm(0).Text3.Text = "nForm(0).Text3.Text" & TmpCutGridY & "." & vbCrLf
    TmpCutGridY = TmpCutGridY + 1
    Loop
End Sub

