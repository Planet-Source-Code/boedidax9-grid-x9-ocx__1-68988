Attribute VB_Name = "Click_Tambahan"

Sub ClickProses(MeTag As String, TXPointerIndex As Long, TYPointerIndex As Long)
    'Form3.ClikData MeTag, TXPointerIndex, TYPointerIndex
End Sub


'Sampah
'Private Sub DrawList(ByVal LIndex As Integer, TmpCutGridX As Integer, TmpCutGridY As Integer, ObjMe As Object, SizeGridList As Integer, CountGridList As Integer)
Private Sub DrawList(TmpCutGridX As Integer, TmpCutGridY As Integer, ObjMe As Object, Optional IndexLList As Integer, Optional MouseDowns As Boolean, Optional CloseBar As Boolean, Optional SubCloseBar As Boolean)       ', SizeGridList As Integer, CountGridList As Integer)
Dim tmpSizeGridList As Integer
Dim XRangeList As Integer
Dim UIndexs As Integer, LIndex As Integer
Dim CountGridList As Integer, SizeGridList As Integer
Dim Y As Single
Dim OverLos As Integer
Dim FormatD As String
Dim YyY As Integer
Dim nY As Single
Dim ColorHead As Long
Dim tmpCountContGList As Integer, tmpGLRange As Integer
Dim CountList As Integer, LoopCountList As Integer, LoopLine As Integer
Dim tmpGLShowCount As Integer
Dim Aab As Integer

Cx = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 0
Cy = GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight


GridXYData(3, 100).GridSubType.GLFrmtStly = "?pt.vis|1 ?pt.x2|150?"
GridXYData(0, 150).GridSubType.GLFrmtStly = "" '?pt.vis|1 ?pt.x2|150 ?pt.full|1 ?" '?pt.x2|50 ?pt.full|0 ?pl.txt.x1|10 ?"
GridXYData(2, 100).GridSubType.GLFrmtStly = "?pt.vis|1 ?pt.x2|150 ?lst.txt.x1|10 ?"

FormatD = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLFrmtStly
LoadFormatPT FormatD
If ListPT.PTVis = 0 Then ListPT.PTFull = 0
If ListPT.PTX2 < 15 Then ListPT.PTVis = 0

tmpCountContGList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1
tmpGLShowCount = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLShowCount - 1

If tmpGLShowCount > -1 Then
tmpCountContGList = tmpGLShowCount
If tmpGLShowCount > GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1 Then
tmpCountContGList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1
tmpGLShowCount = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1
End If
End If

Aab = 20

If tmpCountContGList = 0 Then CountList = GridY(TmpCutGridY).GridHeight - SellHeight_Def - 0 Else _
CountList = ((GridY(TmpCutGridY).GridHeight - SellHeight_Def) \ (tmpCountContGList + 1))
If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight <> 0 Then CountList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight

UIndexs = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLSubScrolIndex

If ListPT.PTVis = 1 Then
If ListPT.PTFull = 1 Then ListPT.PTX2 = GridX(TmpCutGridX).GridWidth

Picture8.Picture = Nothing
Picture8.Width = ListPT.PTX2 - 3
Picture8.Height = CountList - 2

Picture8.PaintPicture Picture7.Picture, 7, 7, ListPT.PTX2 - 3 - 14, CountList - 2 - 14, 7 * 1 + 1, 7 * 1 + 1, 6, 6
Picture8.PaintPicture Image8.Picture, 0, 0, ListPT.PTX2 - 3, , 7 * 1 + 1, 7 * 0 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, ListPT.PTX2 - 6 - 3, 0, , CountList - 3, 7 * 2 + 1, 7 * 1 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, 0, CountList - 6 - 2, ListPT.PTX2 - 3, , 7 * 1 + 1, 7 * 2 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, 0, 0, , CountList - 2, 7 * 0 + 1, 7 * 1 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, 0, 0, , , 7 * 0 + 1, 7 * 0 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, ListPT.PTX2 - 6 - 3, 0, , , 7 * 2 + 1, 7 * 0 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, ListPT.PTX2 - 6 - 3, CountList - 6 - 2, , , 7 * 2 + 1, 7 * 2 + 1, 6, 6
Picture8.PaintPicture Picture7.Picture, 0, CountList - 6 - 2, , , 7 * 0 + 1, 7 * 2 + 1, 6, 6
Picture8.Picture = Picture8.Image
End If

For X = IndexLList To tmpCountContGList
If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight = 0 Then OverLos = 0 Else _
OverLos = SellHeight_Def + GridY(TmpCutGridY).GridTop + GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight * (X + 1)

If tmpGLShowCount > -1 And tmpGLShowCount <> GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1 Then
XRangeList = 15 * 2 + 2
Else
XRangeList = 15 + 2
End If

SizeGridList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLRange
LIndex = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLScrolIndex

CountGridList = Int((CountList - Aab) / SizeGridList) + 0
'Update
Dim ErrorX As Integer
ErrorX = CountGridList - ((GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 0) - LIndex)
If ErrorX > 0 Then
LIndex = LIndex - ErrorX
GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLScrolIndex = LIndex
If LIndex < 0 Then
LIndex = 0
GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLScrolIndex = LIndex
End If
End If

'Update
If LIndex = -1 Then
LIndex = 0
GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(X + UIndexs).GLScrolIndex = LIndex
End If

If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 1 < CountGridList - 0 Then
If tmpGLShowCount = -1 Or tmpGLShowCount = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1 Then
XRangeList = 4
Else
XRangeList = 15 + 4
End If
End If

tmpGLRange = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLRange
LoopCountList = (CountList * (X + 1)) + ((GridY(TmpCutGridY).GridTop + SellHeight_Def))
LoopLine = (Abs(LoopCountList - (GridY(TmpCutGridY).GridTop + SellHeight_Def + CountList))) + Aab ' 20

If tmpGLShowCount > -1 Then
aXRangeList = 1
Else
aXRangeList = 0
End If

If ListPT.PTVis = 1 Then 'Format PicThumb
If ListPT.PTFull = 0 And GridX(TmpCutGridX).GridLeft + ListPT.PTX2 + 10 >= Cx - XRangeList + 2 Then
TMPx2_PicThumb = 2
GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).OpenThumb = False
Else
'Line Thumb
ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, _
(GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17)- _
(GridX(TmpCutGridX).GridLeft + ListPT.PTX2 - 2, LoopCountList), 0, B

Picture1.PaintPicture Picture8.Picture, GridX(TmpCutGridX).GridLeft + 2, _
(GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17

GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).OpenThumb = True
TMPx2_PicThumb = ListPT.PTX2
End If
Else
TMPx2_PicThumb = 2
End If

If ListPT.PTFull = 0 Then  'Update
If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.PointerListHead = X + UIndexs Then '&H00FF8080&
Picture1.FontBold = True
ColorHead = &HFFC0C0 'Color List Head Hit
Else
Picture1.FontBold = False
ColorHead = &HFF8080 'Color List Head Default
End If
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 17 < Cy Then
If (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 1) > Cy Then
YyY = (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 1) - Cy
Else
YyY = 0
End If

'Line Back Head List
ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 17)- _
(Cx - (15 * aXRangeList) - 2, (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 1) - YyY), ColorHead, BF

'Text Head List
TextEffect ObjMe.hDC, GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCaption & vbCrLf & "LLLLL", _
GridX(TmpCutGridX).GridLeft + 5 + TMPx2_PicThumb, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 14, _
Cx - (15 * aXRangeList) - 2, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - YyY, , GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Alignment, 0
End If

Command17.Visible = True
For Y = 0 To GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 1
If ListPT.PTFull <> 0 Then Exit For
If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLPointer = Y + LIndex Then
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine > LoopCountList Then
Rrrx = (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine) - LoopCountList
Else
Rrrx = 0
End If
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine - Rrrx > Cy Then
Rrrx = (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine - Rrrx) - Cy
End If
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine < Cy Then
ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)- _
(Cx - XRangeList + 2, ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine - Rrrx), &H8000000D, BF
End If
tmpColorGrid = vbWhite
Else
tmpColorGrid = 0
End If


tmpSizeGridList = SizeGridList
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine > _
LoopCountList Then
tmpSizeGridList = Abs((CountList - Aab) - (CountGridList * SizeGridList)) 'LoopCountList
End If

If Y + LIndex >= GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 0 Then
Exit For 'Exit Sub
End If

If OverLos > Cy Then
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine + tmpSizeGridList > Cy Then
tmpSizeGridList = (Cy) - (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)  'GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + (CountGridList * SizeGridList) - Cy 'Abs((CountList - Aab) - (CountGridList * SizeGridList))
End If
End If

Picture1.FontBold = False
TextEffect ObjMe.hDC, GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLText(Y + LIndex), _
GridX(TmpCutGridX).GridLeft + 5 + TMPx2_PicThumb, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine, _
GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - XRangeList, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine + tmpSizeGridList, , GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Alignment, tmpColorGrid

sd = 0
'If X > 0 Then sd = 0 '5
If GridY(TmpCutGridY).GridTop + SellHeight_Def + (SizeGridList * (Y + 1) + LoopLine) > _
LoopCountList Then
Exit For
End If


If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine > Cy Then
YyY = ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine
Else
YyY = ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine
End If

If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine < Cy Then
ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)- _
(Cx - XRangeList + 2, YyY), &HC0C0C0, B '&HC0C0C0
End If
'If ...GLHeight <> 0 And ...... Then Exit For
If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight <> 0 And ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine > Cy Then Exit For

Next Y

'Update
If Y + LIndex >= GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 0 Then
'Exit For 'Exit Sub
End If

If OverLos > Cy Then YyY = Cy Else YyY = LoopCountList
If GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine < Cy Then 'MsgBox "O "
ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
(GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine)- _
(Cx - XRangeList + 2, YyY - 0), vbRed, B
End If


If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 1 > CountGridList - 1 Then
'Tengah
If OverLos > Cy Then YyY = (CountList - Aab) - (LoopCountList - Cy) Else YyY = (CountList - Aab)
If YyY > 0 Then
ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, _
GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine, , YyY, 14 * 1, , 14, 13
End If

If CloseBar = False Then
'Dalam Tengah               ^ Nilai tambahan coba dikoreksi
MovingY = ((CountList - Aab + 1) - (13 * 3)) / _
(GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - (CountGridList + 0))

List1.AddItem MovingY & "K-K " & X & " " & Y
'MsgBox GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X).GLCount - CountGridList
'Command17.Caption = ((GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + 13) + (MovingY * GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLScrolIndex) + 0) - Cy

nY = (GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + 13) + _
(MovingY * GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLScrolIndex)

If nY + 13 > Cy Then YyY = Cy - nY Else YyY = 13
If YyY > 0 Then
ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, _
nY, , , 14 * 3, , 14, YyY
End If
End If
If GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + 13 > Cy Then YyY = (CountList - Aab) - (LoopCountList - Cy) Else YyY = 13
If YyY > 0 Then
ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, _
GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine, , , 14 * 0, , 14, YyY
End If
If OverLos > Cy Then YyY = 13 - (LoopCountList - Cy) Else YyY = 13
If YyY > 0 Then
ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, LoopCountList - 12, , , 14 * 2, , 14, YyY
End If

'Command17.Caption = "00000dfsf"
End If
End If

If MouseDowns = True Then Exit For
If OverLos > Cy Then Exit For
Next X

If tmpGLShowCount > -1 Then

ObjMe.PaintPicture Image3.Picture, Cx - 15, GridY(TmpCutGridY).GridTop + SellHeight_Def, , GridY(TmpCutGridY).GridHeight - SellHeight_Def, 14 * 1, 13 * 1, 14, 13



If SubCloseBar = False Then
SubMovingYs = (GridY(TmpCutGridY).GridHeight - SellHeight_Def - (13 * 3)) / _
((GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1) - tmpGLShowCount)

ObjMe.PaintPicture Image3.Picture, Cx - 15, _
(GridY(TmpCutGridY).GridTop + SellHeight_Def + 0 + 13) + _
(SubMovingYs * UIndexs), , , 14 * 3, 13 * 1, 14, 13
List1.AddItem "OJ"
End If

ObjMe.PaintPicture Image3.Picture, Cx - 15, GridY(TmpCutGridY).GridTop + SellHeight_Def, , , 14 * 0, , 14, 13

ObjMe.PaintPicture Image3.Picture, Cx - 15, Cy - 13, , , 14 * 2, , 14, 13
End If

End Sub



