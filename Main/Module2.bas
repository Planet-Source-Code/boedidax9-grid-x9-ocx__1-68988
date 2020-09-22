Attribute VB_Name = "TextEffectToGrid"
Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public m_bDoEffect As Boolean

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xsrc As Long, ByVal ysrc As Long, ByVal dwRop As Long) As Long

Public Declare Function TransparentBlt Lib "msimg32" _
                (ByVal hdcDest As Long, ByVal nXOriginDest As Long, _
                  ByVal nYOriginDest As Long, ByVal nWidthDest As Long, _
                  ByVal nHeightDest As Long, ByVal hdcSrc As Long, _
                  ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                 ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                 ByVal crTransparent As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hDC As Long, ByVal nCharExtra As Long) As Long

Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_BTNFACE = 15
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_DISPFILE = 6            '  Display-file
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1

Public Function TextEffect(iHdc As Long, _
    ByVal sText As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
    Optional ByVal RangeTxt As Integer, Optional ByVal Alignment As Integer = DT_LEFT, _
    Optional ByVal oColor As OLE_COLOR = vbWindowText) As Single

Dim tR As RECT
Dim lLen As Long
Dim hBrush As Long
Dim lCOlor As Long

Select Case Alignment
Case 0
    Alignment = DT_LEFT
Case 1
    Alignment = DT_RIGHT
Case 2
    Alignment = DT_CENTER
End Select

    tR.left = X1: tR.top = Y1: tR.right = X2: tR.bottom = Y2
    
    OleTranslateColor oColor, 0, lCOlor

    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
    lLen = Len(sText)
'errr rrrr  rrrrr rrrr
    SetTextColor iHdc, lCOlor
        SetTextCharacterExtra iHdc, RangeTxt
        TextEffect = DrawText(iHdc, sText, lLen, tR, Alignment)
    DeleteObject hBrush
End Function


Public Function ZTextEffect(iHdc As Long, _
    ByVal sText As String, ByVal X1 As Long, ByVal Y1 As Long, _
    Optional ByVal RangeTxt As Integer, Optional ByVal Alignment As Integer = DT_LEFT, _
    Optional ByVal oColor As OLE_COLOR = vbWindowText) As Single

Dim tR As RECT
Dim lLen As Long
Dim hBrush As Long
Dim lCOlor As Long

    tR.left = X1: tR.top = Y1: tR.right = X1: tR.bottom = Y1
    
    OleTranslateColor oColor, 0, lCOlor

    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
    lLen = Len(sText)

    SetTextColor iHdc, lCOlor
        SetTextCharacterExtra iHdc, RangeTxt
        TextEffect = DrawText(iHdc, sText, lLen, tR, DT_CALCRECT)
        TextEffect = DrawText(iHdc, sText, lLen, tR, Alignment)
    DeleteObject hBrush
End Function


