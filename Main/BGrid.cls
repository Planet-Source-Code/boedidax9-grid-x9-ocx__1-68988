VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BGrid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'back to poprty page

'----------------------------------------------------------
'BGrid V 1.0
'By: BoedidaX9
'Email: Nurfaststar@yahoo.com
'----------------------------------------------------------
Private Type PointInvert
    InX1 As Single
    InY1 As Single
    InX2 As Single
    InY2 As Single
End Type
Private InvertME() As PointInvert


Private SetNewGrid As NewGridXY
Private GridX() As TypeGridX '99 ----------
Private GridY() As TypeGridY '3999 --------
Private GridXYData() As TypeGridXYData  '99,3999

Private PictureSub As Image
Private PictureXY  As Image
Private PicErr As Boolean

'Public SetNewGrid.GridSize.GDRangeX As Integer, SetNewGrid.GridSize.GDRangeY As Integer
'Public SetNewGrid.GridSize.GDHeight As Integer, SetNewGrid.GridSize.GDWidth As Integer
'Public SetNewGrid.GridSize.GDXYHeight As Single, SetNewGrid.GridSize.GDXYWidth As Single

Public GridUpX As Integer, GridUpY As Integer
Public GridDownX As Integer, GridDownY As Integer
'Public Tager As String
'Public SetNewGrid.GridSizePicSub.RangePicSubX1 As Integer,
'SetNewGrid.GridSizePicSub.RangePicSuby1 As Integer
'Public SetNewGrid.GridSizePicSub.RangePicSubX2 As Integer,
'SetNewGrid.GridSizePicSub.RangePicSubY2 As Integer

'Public RangePicGridX1 As Integer, RangePicGridy1 As Integer
'Public SetNewGrid.GridSizePicSub.RangePicGridx2 As Integer, SetNewGrid.GridSizePicSub.RangePicGridy2 As Integer
'Public SetNewGrid.GridSizePicSub.RangePicGridCW As Integer, SetNewGrid.GridSizePicSub.RangePicGridCH As Integer

'Property Get M0_Value(ByVal IndexX As Long, ByVal IndexY As Long) As String
'    M0_Value = GridXYData(IndexX, IndexY).GridXYValue
'End Property
'Property Let M0_Value(ByVal IndexX As Long, ByVal IndexY As Long, ByVal MAValue As String)
'    GridXYData(IndexX, IndexY).GridXYValue = MAValue
'End Property


'Property Get Tag(ByVal IndexX As Long, ByVal IndexY As Long) As String
'    Tag = GridXYData(IndexX, IndexY).GridTag
'End Property
'Property Let Tag(ByVal IndexX As Long, ByVal IndexY As Long, ByVal MAValue As String)
'    GridXYData(IndexX, IndexY).GridTag = MAValue
'End Property

Property Get PictureErr() As Boolean
    PictureErr = PicErr
End Property
Property Let PictureErr(ByVal MAValue As Boolean)
    PicErr = MAValue
End Property

Property Get FilePicture() As String
    FilePicture = SetNewGrid.GridFilePicture
End Property
Property Let FilePicture(ByVal MAValue As String)
    SetNewGrid.GridFilePicture = MAValue
End Property

Property Get GridType() As Integer
    GridType = SetNewGrid.GridType
End Property
Property Let GridType(ByVal MAValue As Integer)
    SetNewGrid.GridType = MAValue
End Property

Property Get RangePicGridX1() As Integer
    RangePicGridX1 = SetNewGrid.GridSizePic.RangePicGridX1
End Property
Property Let RangePicGridX1(ByVal MAValue As Integer)
    SetNewGrid.GridSizePic.RangePicGridX1 = MAValue
End Property

Property Get RangePicGridY1() As Integer
    RangePicGridY1 = SetNewGrid.GridSizePic.RangePicGridY1
End Property
Property Let RangePicGridY1(ByVal MAValue As Integer)
    SetNewGrid.GridSizePic.RangePicGridY1 = MAValue
End Property

Property Get RangePicGridX2() As Integer
    RangePicGridX2 = SetNewGrid.GridSizePic.RangePicGridX2
End Property
Property Let RangePicGridX2(ByVal MAValue As Integer)
    SetNewGrid.GridSizePic.RangePicGridX2 = MAValue
End Property

Property Get RangePicGridY2() As Integer
    RangePicGridY2 = SetNewGrid.GridSizePic.RangePicGridY2
End Property
Property Let RangePicGridY2(ByVal MAValue As Integer)
    SetNewGrid.GridSizePic.RangePicGridY2 = MAValue
End Property

Property Get RangePicGridCW() As Integer
    RangePicGridCW = SetNewGrid.GridSizePic.RangePicGridCW
End Property
Property Let RangePicGridCW(ByVal MAValue As Integer)
    SetNewGrid.GridSizePic.RangePicGridCW = MAValue
End Property

Property Get RangePicGridCH() As Integer
    RangePicGridCH = SetNewGrid.GridSizePic.RangePicGridCH
End Property
Property Let RangePicGridCH(ByVal MAValue As Integer)
    SetNewGrid.GridSizePic.RangePicGridCH = MAValue
End Property






Property Get RangePicSubX1() As Integer
    RangePicSubX1 = SetNewGrid.GridSizePicSub.RangePicSubX1
End Property
Property Let RangePicSubX1(ByVal MAValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubX1 = MAValue
End Property

Property Get RangePicSubY1() As Integer
    RangePicSubY1 = SetNewGrid.GridSizePicSub.RangePicSubY1
End Property
Property Let RangePicSubY1(ByVal MAValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubY1 = MAValue
End Property

Property Get RangePicSubX2() As Integer
    RangePicSubX2 = SetNewGrid.GridSizePicSub.RangePicSubX2
End Property
Property Let RangePicSubX2(ByVal MAValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubX2 = MAValue
End Property

Property Get RangePicSubY2() As Integer
    RangePicSubY2 = SetNewGrid.GridSizePicSub.RangePicSubY2
End Property
Property Let RangePicSubY2(ByVal MAValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubY2 = MAValue
End Property




Property Get RangeX() As Integer
    RangeX = SetNewGrid.GridSize.GDRangeX
End Property
Property Let RangeX(ByVal MAValue As Integer)
    SetNewGrid.GridSize.GDRangeX = MAValue
End Property

Property Get RangeY() As Integer
    RangeY = SetNewGrid.GridSize.GDRangeY
End Property
Property Let RangeY(ByVal MAValue As Integer)
    SetNewGrid.GridSize.GDRangeY = MAValue
End Property

Property Get GDWidth() As Integer
    GDWidth = SetNewGrid.GridSize.GDWidth
End Property
Property Let GDWidth(ByVal MAValue As Integer)
    SetNewGrid.GridSize.GDWidth = MAValue
End Property

Property Get GDHeight() As Integer
    GDHeight = SetNewGrid.GridSize.GDHeight
End Property
Property Let GDHeight(ByVal MAValue As Integer)
    SetNewGrid.GridSize.GDHeight = MAValue
End Property

Property Get GDXYWidth() As Single
    GDXYWidth = SetNewGrid.GridSize.GDXYWidth
End Property
Property Let GDXYWidth(ByVal MAValue As Single)
    SetNewGrid.GridSize.GDXYWidth = MAValue
End Property

Property Get GDXYHeight() As Single
    GDXYHeight = SetNewGrid.GridSize.GDXYHeight
End Property
Property Let GDXYHeight(ByVal MAValue As Single)
    SetNewGrid.GridSize.GDXYHeight = MAValue
End Property


Property Get NameTabel() As String
    NameTabel = SetNewGrid.GridXYName
End Property
Property Let NameTabel(ByVal MAValue As String)
    SetNewGrid.GridXYName = MAValue
End Property

Property Get CountX() As Long
    CountX = SetNewGrid.GridXCount
End Property
Property Get CountY() As Long
    CountY = SetNewGrid.GridYCount
End Property

Property Get M0_Value(ByVal IndexX As Long, ByVal IndexY As Long) As String
    M0_Value = GridXYData(IndexX, IndexY).GridXYValue
End Property
Property Let M0_Value(ByVal IndexX As Long, ByVal IndexY As Long, ByVal MAValue As String)
    GridXYData(IndexX, IndexY).GridXYValue = MAValue
End Property

Property Get M0_ValueSub(ByVal IndexX As Long, ByVal IndexY As Long) As String
    M0_ValueSub = GridXYData(IndexX, IndexY).GridXYValueSub
End Property
Property Let M0_ValueSub(ByVal IndexX As Long, ByVal IndexY As Long, ByVal MAValue As String)
    GridXYData(IndexX, IndexY).GridXYValueSub = MAValue
End Property

Property Get M0_PicIndexXY(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    M0_PicIndexXY = GridXYData(IndexX, IndexY).GridXYPicIndex
End Property
Property Let M0_PicIndexXY(ByVal IndexX As Long, ByVal IndexY As Long, ByVal MAValue As Integer)
    GridXYData(IndexX, IndexY).GridXYPicIndex = MAValue
End Property


Property Get SelLeft(ByVal Index As Long, Optional ByVal WDefault As Boolean) As Single
    SelLeft = GridX(Index).GridLeft
End Property
'Property Let SelLeft(ByVal Index As Long, Optional ByVal WDefault As Boolean, ByVal MAValue As Single)
'    GridX(Index).GridLeft = MAValue
'        GridX(Index).GWidthDefault = WDefault
'End Property

Property Get SelWidth(ByVal Index As Long, Optional ByVal WDefault As Boolean) As Single
    SelWidth = GridX(Index).GridWidth
        'GridX(Index).GWidthDefault = WDefault
End Property
Property Let SelWidth(ByVal Index As Long, Optional ByVal WDefault As Boolean, ByVal MAValue As Single)
    GridX(Index).GridWidth = MAValue
        GridX(Index).GWidthDefault = WDefault
End Property

Property Get SelValueX(ByVal Index As Long) As String
    SelValueX = GridX(Index).GridValue
End Property
Property Let SelValueX(ByVal Index As Long, ByVal MAValue As String)
    GridX(Index).GridValue = MAValue
End Property

Sub SizeLeftRight(Index As Long, MAValue As Boolean)

End Sub

Property Get SelXPicture(ByVal Index As Long) As Boolean
    SelXPicture = GridX(Index).GPicturePut
End Property
Property Let SelXPicture(ByVal Index As Long, ByVal MAValue As Boolean)
    GridX(Index).GPicturePut = MAValue
End Property



Property Get SelTop(ByVal Index As Long, Optional ByVal HDefault As Boolean) As Single
    SelTop = GridY(Index).GridTop
End Property
'Property Let SelTop(ByVal Index As Long, Optional ByVal HDefault As Boolean, ByVal MAValue As Single)
'    GridY(Index).GridTop = MAValue
'        GridY(Index).GHidthDefault = HDefault
'End Property

Property Get SelHeight(ByVal Index, Optional ByVal HDefault As Boolean) As Single
    'MsgBox UBound(GridY())
    SelHeight = GridY(Index).GridHeight
        'GridY(Index).GHidthDefault = HDefault
End Property
Property Let SelHeight(ByVal Index, Optional ByVal HDefault As Boolean, ByVal MAValue As Single)
    GridY(Index).GridHeight = SetNewGrid.GridSize.GDHeight + MAValue
        GridY(Index).GHeightDefault = HDefault
End Property

Property Get SelValueY(ByVal Index As Long) As String
    SelValueY = GridY(Index).GridValue
End Property
Property Let SelValueY(ByVal Index As Long, ByVal MAValue As String)
    GridY(Index).GridValue = MAValue
End Property

Sub SizeUpDown(Index As Long, MAValue As Boolean)
    If MAValue = False Then
        GridY(Index).GHSave = GridY(Index).GridHeight
        GridY(Index).GHeightDefault = False
    Else
        If GridY(Index).GHSave = 0 Then GridY(Index).GHSave = SetNewGrid.GridSize.GDHeight + 80 'Add
        GridY(Index).GridHeight = GridY(Index).GHSave
        GridY(Index).GHeightDefault = True
    End If
End Sub







Sub GridXY(ObjMe As Object, GridUpX As Integer, GridUpY As Integer, Optional ShowPic As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long

'ObjMe.Picture = Nothing
'ObjMe.Cls
'ObjMe.BackColor = 0

TmpCutGridX = GridUpX
TmpCutGridY = GridUpY

GridX(GridUpX).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

'Reset SetNewGrid.GridSize.GDHeight
'Form1.AutoRedraw

'Form1.AutoRedraw
'ObjMe.AutoRedraw = False

Do 'X
    If GridX(TmpCutGridX).GWidthDefault = False Then GridX(TmpCutGridX).GridWidth = SetNewGrid.GridSize.GDWidth '100
    If TmpCutGridX > GridUpX Then _
    GridX(TmpCutGridX).GridLeft = GridX(TmpCutGridX - 1).GridLeft + GridX(TmpCutGridX - 1).GridWidth
    
    Do 'Y
        If GridY(TmpCutGridY).GHeightDefault = False Then GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.GDHeight  '20
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight

        If DoitY = False Then DrawDrid ObjMe, TmpCutGridX, TmpCutGridY, 1
        DrawDrid ObjMe, TmpCutGridX, TmpCutGridY, 0
        
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.GDXYHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, TmpCutGridX, TmpCutGridY, 2
    
    If GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth >= SetNewGrid.GridSize.GDXYWidth Or _
    TmpCutGridX + 1 = SetNewGrid.GridXCount Then
        GridDownX = TmpCutGridX '- 10
        Exit Do
    End If
TmpCutGridX = TmpCutGridX + 1
DoitY = True
Loop

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
(ObjMe.ScaleWidth, ObjMe.ScaleHeight), &H8000000F, BF

ObjMe.Line (GridX(GridDownX).GridLeft + GridX(GridDownX).GridWidth + 1, 0)- _
(ObjMe.ScaleWidth, ObjMe.ScaleHeight), &H8000000F, BF

'Form1.ScaleWidth
'ObjMe.Picture = ObjMe.Image
'ObjMe.AutoRedraw = True
End Sub

Sub DrawDrid(ObjMe As Object, TmpCutGridX As Integer, TmpCutGridY As Integer, TypeDraw As Integer)
Dim TmpRangePicGridX1 As Integer, TmpRangePicGridX2 As Integer

Cx = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 0
Cy = GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight
        
Select Case TypeDraw
Case 0

    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 0, GridY(TmpCutGridY).GridTop + 0)- _
                  (Cx, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.GDHeight), &H8000000F, BF
    ObjMe.Line (GridX(TmpCutGridX).GridLeft - 1, GridY(TmpCutGridY).GridTop + 2)- _
                  (Cx, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.GDHeight), vbWhite, B
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 0, GridY(TmpCutGridY).GridTop + 0)- _
                  (Cx, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.GDHeight), 0, B
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 0, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.GDHeight + 1)- _
                  (Cx, Cy), vbWhite, BF
    ObjMe.Line (GridX(TmpCutGridX).GridLeft, GridY(TmpCutGridY).GridTop)- _
                  (Cx, Cy), , B

    If GridX(TmpCutGridX).GPicturePut = True Then
        TmpRangePicGridX1 = SetNewGrid.GridSizePic.RangePicGridX1
        TmpRangePicGridX2 = SetNewGrid.GridSizePic.RangePicGridX2
    End If
    
    TextEffect ObjMe.hdc, GridXYData(TmpCutGridX, TmpCutGridY).GridXYValue, _
    GridX(TmpCutGridX).GridLeft + TmpRangePicGridX1 + TmpRangePicGridX2 + 5, GridY(TmpCutGridY).GridTop + 3, , , 0
    
    TextEffect ObjMe.hdc, GridXYData(TmpCutGridX, TmpCutGridY).GridXYValueSub, _
    GridX(TmpCutGridX).GridLeft + 3, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.GDHeight + 1, , , 0
    

    If GridX(TmpCutGridX).GPicturePut = True Then
        DrawXY ObjMe, GridXYData(TmpCutGridX, TmpCutGridY).GridXYPicIndex, _
        TmpCutGridX, TmpCutGridY
    End If
   
Case 1
    ObjMe.Line (0, GridY(TmpCutGridY).GridTop + 0)-(SetNewGrid.GridSize.GDRangeX + 0, Cy), &H8000000F, BF
    ObjMe.Line (1, GridY(TmpCutGridY).GridTop + 1)-(SetNewGrid.GridSize.GDRangeX + 1, Cy), vbWhite, B
    ObjMe.Line (0, GridY(TmpCutGridY).GridTop)-(SetNewGrid.GridSize.GDRangeX + 0, Cy), , B
    If GridY(TmpCutGridY).GridHeight <= SetNewGrid.GridSize.GDHeight Then
        DrawSubY ObjMe, 0, 0, TmpCutGridY
        f = 1
    Else
        DrawSubY ObjMe, 1, 0, TmpCutGridY
        If GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.GDHeight + (SetNewGrid.GridSize.GDHeight / 4) Then _
        f = 2 Else f = 1
    End If
    
    cf = (GridY(TmpCutGridY).GridTop + ((GridY(TmpCutGridY).GridHeight - ObjMe.TextHeight(GridY(TmpCutGridY).GridValue)) \ f))
    D = GridY(TmpCutGridY).GridValue 'SetNewGrid.GridSize.GDHeight  ' GridY(TmpCutGridY).GridValue
    TextEffect ObjMe.hdc, D, _
    (SetNewGrid.GridSize.GDRangeX - ObjMe.TextWidth(D)) / 2, cf, , , 0

Case 2
    ObjMe.Line (GridX(TmpCutGridX).GridLeft, 0)-(Cx, SetNewGrid.GridSize.GDRangeY - 1), &H8000000F, BF
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 1, 1)-(Cx, SetNewGrid.GridSize.GDRangeY - 0), vbWhite, B
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 0, 0)-(Cx, SetNewGrid.GridSize.GDRangeY - 0), , B
    
    cf = (GridX(TmpCutGridX).GridLeft + ((GridX(TmpCutGridX).GridWidth - ObjMe.TextWidth(GridX(TmpCutGridX).GridValue)) / 2))
    TextEffect ObjMe.hdc, GridX(TmpCutGridX).GridValue, _
    cf, 2, , , 0
End Select
End Sub

Sub DrawSubY(ObjMe As Object, IndexX As Integer, IndexY As Integer, TmpCutGridY As Integer)
'If SetNewGrid.GridSizePicSub.RangePicSubX1 < 1 Then SetNewGrid.GridSizePicSub.RangePicSubX1 = 3
'If SetNewGrid.GridSizePicSub.RangePicSuby1 < 1 Then SetNewGrid.GridSizePicSub.RangePicSuby1 = 3
If SetNewGrid.GridSizePicSub.RangePicSubX2 < 1 Then SetNewGrid.GridSizePicSub.RangePicSubX2 = 9
If SetNewGrid.GridSizePicSub.RangePicSubY2 < 1 Then SetNewGrid.GridSizePicSub.RangePicSubY2 = 9

    ObjMe.PaintPicture PictureSub, _
    SetNewGrid.GridSizePicSub.RangePicSubX1, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePicSub.RangePicSubY1, , , SetNewGrid.GridSizePicSub.RangePicSubX2 * IndexX, SetNewGrid.GridSizePicSub.RangePicSubX2 * IndexY, SetNewGrid.GridSizePicSub.RangePicSubX2, SetNewGrid.GridSizePicSub.RangePicSubY2
End Sub
Sub SLinkPictureSub(ObjMe As Object)
Set PictureSub = ObjMe
End Sub

Sub DrawXY(ObjMe As Object, IndexXY As Integer, TmpCutGridX As Integer, TmpCutGridY As Integer)
Dim IndexX As Integer, IndexY As Integer

If SetNewGrid.GridSizePic.RangePicGridX1 < 1 Then SetNewGrid.GridSizePic.RangePicGridX1 = 3
If SetNewGrid.GridSizePic.RangePicGridY1 < 1 Then SetNewGrid.GridSizePic.RangePicGridY1 = (SetNewGrid.GridSize.GDHeight - SetNewGrid.GridSizePic.RangePicGridY2) / 2
If SetNewGrid.GridSizePic.RangePicGridX2 < 1 Then SetNewGrid.GridSizePic.RangePicGridX2 = 20
If SetNewGrid.GridSizePic.RangePicGridY2 < 1 Then SetNewGrid.GridSizePic.RangePicGridY2 = 20
    
    IndexY = IndexXY \ (SetNewGrid.GridSizePic.RangePicGridCW + 1)
    IndexX = IndexXY - (IndexY * (SetNewGrid.GridSizePic.RangePicGridCW + 1))
   
If PicErr = False Then
'On Error GoTo 10
    ObjMe.PaintPicture PictureXY, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePic.RangePicGridX1, _
    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePic.RangePicGridY1, _
    , , SetNewGrid.GridSizePic.RangePicGridX2 * IndexX, _
    SetNewGrid.GridSizePic.RangePicGridY2 * IndexY, SetNewGrid.GridSizePic.RangePicGridX2 - 0, SetNewGrid.GridSizePic.RangePicGridY2 - 0
Else
    ObjMe.PaintPicture PictureXY, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePic.RangePicGridX1, _
    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePic.RangePicGridY1 ', _
    , , SetNewGrid.GridSizePic.RangePicGridX2 * IndexX, _
    SetNewGrid.GridSizePic.RangePicGridY2 * IndexY, SetNewGrid.GridSizePic.RangePicGridX2 - 0, SetNewGrid.GridSizePic.RangePicGridY2 - 0
End If

Exit Sub
10
'MsgBox "Not SLinkPictureXY to PictureXY"
'End
End Sub


















'Sub iDrawXY(ObjMe As Object, IndexX As Integer, IndexY As Integer, TmpCutGridX As Integer, TmpCutGridY As Integer)
'If SetNewGrid.GridSizePicSub.RangePicGridX1 < 1 Then SetNewGrid.GridSizePicSub.RangePicGridX1 = 3
'If SetNewGrid.GridSizePicSub.RangePicGridY1 < 1 Then SetNewGrid.GridSizePicSub.RangePicGridY1 = (SetNewGrid.GridSize.GDHeight - 15) / 2
'If SetNewGrid.GridSizePicSub.RangePicGridX2 < 1 Then SetNewGrid.GridSizePicSub.RangePicGridX2 = 20
'If SetNewGrid.GridSizePicSub.RangePicGridY2 < 1 Then SetNewGrid.GridSizePicSub.RangePicGridY2 = 20
'
'On Error GoTo 10
'    ObjMe.PaintPicture PictureXY, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePicSub.RangePicGridX1, _
'    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePicSub.RangePicGridY1, _
'    , , (SetNewGrid.GridSizePicSub.RangePicGridX2 * IndexX - IndexX) + 0, _
'    (SetNewGrid.GridSizePicSub.RangePicGridY2 * IndexY - IndexY) + 1, SetNewGrid.GridSizePicSub.RangePicGridX2 - 0, SetNewGrid.GridSizePicSub.RangePicGridY2 - 0
'Exit Sub
'10
'MsgBox "Not SLinkPictureXY to PictureXY"
'End
'End Sub
Sub SLinkPictureXY(ObjMe As Object)
    If SetNewGrid.GridFilePicture = "" Then PicErr = True
    Set PictureXY = ObjMe
End Sub

Sub Initial()
    SetNewGrid.GridXCount = 1
    SetNewGrid.GridYCount = 1
    ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
    ReDim Preserve GridY(SetNewGrid.GridYCount - 1)
    ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
End Sub

Sub Add(Optional ByVal VX As Long, Optional ByVal VY As Long)
Dim TmpGridXYData() As TypeGridXYData
Dim TmpVX As Long, TmpVY As Long
Dim A As Boolean

'    TmpVX = SetNewGrid.GridXCount
'    TmpVY = SetNewGrid.GridYCount
'    If VX > SetNewGrid.GridXCount Then SetNewGrid.GridXCount = VX
'    If VX = 0 Then SetNewGrid.GridXCount = SetNewGrid.GridXCount + 1
'    If VY > SetNewGrid.GridYCount Then SetNewGrid.GridYCount = VY
'    If VY = 0 Then SetNewGrid.GridYCount = SetNewGrid.GridYCount + 1
        'If VX = 0 Then VX = SetNewGrid.GridXCount + 1
        'If VY = 0 Then VY = SetNewGrid.GridYCount + 1
        If VY < 0 Or VY < 0 Then Exit Sub
'        If VX <> TmpVX Then
        If VX > 0 Then
            ReDim Preserve GridX(VX - 1)
            
            TmpGridXYData() = GridXYData()
'            MsgBox TmpGridXYData(0, 0).GridXYValue
            If VY = 0 Then VY = SetNewGrid.GridYCount
            
            ReDim GridXYData(VX - 1, VY - 1)
            For X = 0 To VX - 1
                If X >= SetNewGrid.GridXCount - 0 Then Exit For
                For Y = 0 To VY - 1
                    If Y >= SetNewGrid.GridYCount - 0 Then Exit For
                    GridXYData(X, Y) = TmpGridXYData(X, Y)
                Next Y
            Next X
            SetNewGrid.GridXCount = VX
            If VY > 0 Then
                A = True
                SetNewGrid.GridYCount = VY
                ReDim Preserve GridY(VY - 1)
            End If
        End If
        If VY > 0 And A = False Then
            ReDim Preserve GridY(VY - 1)
            ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, VY - 1)
            SetNewGrid.GridYCount = VY
        End If

'        Else
'            MsgBox SetNewGrid.GridXCount & " " & SetNewGrid.GridYCount
'            ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
'        End If
End Sub

Sub Add1(Optional ByVal VX As Long, Optional ByVal VY As Long)
Dim TmpGridXYData() As TypeGridXYData
Dim TmpVX As Long, TmpVY As Long

    TmpVX = SetNewGrid.GridXCount
    TmpVY = SetNewGrid.GridYCount
    If VX > SetNewGrid.GridXCount Then SetNewGrid.GridXCount = VX
    If VX = 0 Then SetNewGrid.GridXCount = SetNewGrid.GridXCount + 1
    If VY > SetNewGrid.GridYCount Then SetNewGrid.GridYCount = VY
    If VY = 0 Then SetNewGrid.GridYCount = SetNewGrid.GridYCount + 1
    
        ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
        ReDim Preserve GridY(SetNewGrid.GridYCount - 1)
        
        If VX <> TmpVX Then
            TmpGridXYData() = GridXYData()
            ReDim GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
            For X = 0 To TmpVX - 1
                For Y = 0 To TmpVY - 1
                    GridXYData(X, Y) = TmpGridXYData(X, Y)
                Next Y
            Next X
        Else
'            MsgBox SetNewGrid.GridXCount & " " & SetNewGrid.GridYCount
            ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
        End If
End Sub

Sub Remove(Optional ByVal IndexX As Long = -1, Optional ByVal IndexY As Long = -1)
Dim TmpGridXYData() As TypeGridXYData
    If IndexX = 0 Then IndexX = SetNewGrid.GridXCount
    If IndexX > 0 Then
'        For X = IndexX To SetNewGrid.GridXCount - 1
'            GridX(X - 1) = GridX(X)
'        Next X
    
        TmpGridXYData() = GridXYData()
'        ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
        SetNewGrid.GridXCount = SetNewGrid.GridXCount - 1
        ReDim GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
        
        For X = 0 To SetNewGrid.GridXCount - 1
            If X >= IndexX Then GridX(X - 1) = GridX(X)
            For Y = 0 To SetNewGrid.GridYCount - 1
                If X >= IndexX Then
                    GridXYData(X - 1, Y) = TmpGridXYData(X, Y) '???????
                Else
                    GridXYData(X, Y) = TmpGridXYData(X, Y)
                End If
            Next Y
        Next X
'        SetNewGrid.GridXCount = SetNewGrid.GridXCount - 1
        ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
    End If
    
    If IndexY = 0 Then IndexY = SetNewGrid.GridYCount
    If IndexY > 0 Then
        'For Y = IndexY To SetNewGrid.GridYCount - 1
        '    GridY(Y - 1) = GridY(Y)
        'Next Y
    
        For X = 0 To SetNewGrid.GridXCount - 1
            For Y = IndexY To SetNewGrid.GridYCount - 1
                If X = 0 Then GridY(Y - 1) = GridY(Y)
                GridXYData(X, Y - 1) = GridXYData(X, Y)
            Next Y
        Next X
        SetNewGrid.GridYCount = SetNewGrid.GridYCount - 1
            ReDim Preserve GridY(SetNewGrid.GridYCount - 1)
            ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
    End If
End Sub
Private Sub ToolRemove()
End Sub

Sub RemoveAdd(IndexY As Long)
Dim TmpGridXYData() As TypeGridXYData
Dim TmpGridY() As TypeGridY
Dim Tmp(0, 0) As TypeGridXYData
Dim Tmp2(0) As TypeGridY

    TmpGridXYData() = GridXYData()
    TmpGridY() = GridY()
    
    Add -1, 0
    For X = 0 To SetNewGrid.GridXCount - 1
        GridXYData(X, IndexY) = Tmp(0, 0)
        GridY(IndexY) = Tmp2(0)
        For Y = IndexY + 1 To SetNewGrid.GridYCount - 1
            GridXYData(X, Y) = TmpGridXYData(X, Y - 1)
            If X = 0 Then GridY(Y) = TmpGridY(Y - 1)
        Next Y
    Next X
End Sub


Sub SetData(CFileAll As Integer)
'ReDim Preserve FileAll(CFileAll)
    SetNewGrid = FileAll(CFileAll).GridNew
    GridX() = FileAll(CFileAll).GridX()
    GridY() = FileAll(CFileAll).GridY()
    GridXYData() = FileAll(CFileAll).GridXYData()
End Sub

Sub SetDataX()
    SetNewGrid = J.GridNew
    GridX() = J.GridX()
    GridY() = J.GridY()
    GridXYData() = J.GridXYData()
End Sub

Sub GetData(CFileAll As Integer)
'ReDim Preserve FileAll(CFileAll)
    FileAll(CFileAll).GridNew = SetNewGrid
    FileAll(CFileAll).GridX() = GridX()
    FileAll(CFileAll).GridY() = GridY()
    FileAll(CFileAll).GridXYData() = GridXYData()
End Sub

'Public Sub SaveDataGrid_FileAll(SetNewGrid As NewGridXY, GridX() As TypeGridX, GridY() As TypeGridY, GridXYData() As TypeGridXYData, CountFileAll As Integer)
'    EditDataGrid_FileAll SetNewGrid, GridX(), GridY(), GridXYData(), CountFileAll
'    CountFileAll = CountFileAll + 1
'End Sub

'Public Sub EditDataGrid_FileAll(SetNewGrid As NewGridXY, GridX() As TypeGridX, GridY() As TypeGridY, GridXYData() As TypeGridXYData, IndexFileAll As Integer)
'    FileAll(IndexFileAll).GridNew = SetNewGrid
    'FileAll(IndexFileAll).GridCountX = UBound(GridX())
'    FileAll(IndexFileAll).GridX() = GridX()
    'FileAll(IndexFileAll).GridCountY = UBound(GridY())
'    FileAll(IndexFileAll).GridY() = GridY()
'    FileAll(IndexFileAll).GridXYData() = GridXYData()
'End Sub

'Public Sub OpenDataGrid_FileAll(SetNewGrid As NewGridXY, GridX() As TypeGridX, GridY() As TypeGridY, GridXYData() As TypeGridXYData, IndexFileAll As Integer)
'    SetNewGrid = FileAll(IndexFileAll).GridNew
'    GridX() = FileAll(IndexFileAll).GridX()
'    GridY() = FileAll(IndexFileAll).GridY()
'    GridXYData() = FileAll(IndexFileAll).GridXYData()
'End Sub



