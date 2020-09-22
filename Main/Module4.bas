Attribute VB_Name = "InvertPaint"
Private Type CCCPointInvert
    Px1 As Long
    Px2 As Long
    Py1 As Long
    Py2 As Long
End Type
Public TmpPxy As PointInvert
'Public Txy As PointInvert

Public Const ContVButton = 2
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public lrect As RECT



Sub DrawInvert(ObjMe As Object, iLeft As Long, iTop As Long, iWidth As Long, iHeight As Long)
    lrect.left = iLeft
    lrect.top = iTop
    lrect.right = iWidth
    lrect.bottom = iHeight
'    InvertRect ObjMe.hDC, lrect
End Sub

Sub CCDrawInvertToGrid(ByVal tBGrid As BGrid, ByVal ObjMe As Object, iPxy As PointInvert, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
Dim TmpPxy As PointInvert
Dim Cx As Long, Cy As Long

iPxy.Px1 = Px1
iPxy.Py1 = Py1
iPxy.Px2 = Px2
iPxy.Py2 = Py2

'    If iPxy.Px2 >= tBGrid.CountX Then
'        iPxy.Px1 = tBGrid.CountX - 1
'        iPxy.Px2 = tBGrid.CountX - 1
'    End If
'    If iPxy.Py2 >= tBGrid.CountY Then
'        iPxy.Py1 = tBGrid.CountY - 1
'        iPxy.Py2 = tBGrid.CountY - 1
'    End If
    If tBGrid.SelLeft(Px1) = 0 Then Px1 = tBGrid.GridUpX
    If tBGrid.SelTop(Py1) = 0 Then Py1 = tBGrid.GridUpY

    If tBGrid.GridUpX > Px1 Then Px1 = tBGrid.GridUpX
    If Px2 > tBGrid.GridDownX Then Px2 = tBGrid.GridDownX
    
    If tBGrid.GridUpY > Py1 Then Py1 = tBGrid.GridUpY
    If Py2 > tBGrid.GridDownY Then Py2 = tBGrid.GridDownY
    
    Cx = tBGrid.SelLeft(Px2) + tBGrid.SelWidth(Px2) - 1
    Cy = tBGrid.SelTop(Py2) + tBGrid.SelHeight(Py2) - 1
    
    If Cx = 0 Then
        Cx = tBGrid.SelLeft(tBGrid.GridDownX) + tBGrid.SelWidth(tBGrid.GridDownX)
    End If
    If Cy = 0 Then
        Cy = tBGrid.SelTop(tBGrid.GridDownY) + tBGrid.SelHeight(tBGrid.GridDownY)
    End If
    
    
    DrawInvert ObjMe, tBGrid.SelLeft(Px1) + 2, tBGrid.SelTop(Py1) + 2, Cx, Cy
      
'    DrawInvert ObjMe, tBGrid.SelLeft(XPointerIndex) + 3, tBGrid.SelTop(YPointerIndex + 2) + 3, _
    Cx - 2, Cy - 2
'End If
End Sub

Sub xDrawInvertToGrid(tBGrid As BGrid, ObjMe As Object, iPxy As PointInvert, Px1 As Long, Py1 As Long, Px2 As Long, Py2 As Long)
Dim Cx As Long, Cy As Long

'If tBGrid.GridType <> 1 And tBGrid.GridType <> 3 Then
'If CheckType(tBGrid.GridType, 0, 4) = False Then

iPxy.Px1 = Px1
iPxy.Py1 = Py1
iPxy.Px2 = Px2
iPxy.Py2 = Py2

    If iPxy.Px2 >= tBGrid.CountX Then
        iPxy.Px1 = tBGrid.CountX - 1
        iPxy.Px2 = tBGrid.CountX - 1
    End If
    If iPxy.Py2 >= tBGrid.CountY Then
        iPxy.Py1 = tBGrid.CountY - 1
        iPxy.Py2 = tBGrid.CountY - 1
    End If
    
    
    Cx = tBGrid.SelLeft(Px2) + tBGrid.SelWidth(Px2)
    Cy = tBGrid.SelTop(Py2) + tBGrid.SelHeight(Py2)
    
    If Cx = 0 Then Cx = tBGrid.SelLeft(tBGrid.GridDownX) + tBGrid.SelWidth(tBGrid.GridDownX)
    If Cy = 0 Then Cy = tBGrid.SelTop(tBGrid.GridDownY) + tBGrid.SelHeight(tBGrid.GridDownY)
    
    DrawInvert ObjMe, tBGrid.SelLeft(Px1) + 1, tBGrid.SelTop(Py1) + 1, Cx, Cy
'    DrawInvert ObjMe, tBGrid.SelLeft(XPointerIndex) + 3, tBGrid.SelTop(YPointerIndex + 2) + 3, _
    Cx - 2, Cy - 2
'End If
End Sub

Sub Lo(gg As BGrid)
gg.M0_Value(0, 0) = "OK"
End Sub

