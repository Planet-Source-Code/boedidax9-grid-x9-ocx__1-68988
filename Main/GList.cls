VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type LScrolBar
    LBx1 As Integer
    LBy1 As Integer
    LBx2 As Integer
    LBy2 As Integer
'    LBIndex As Long
End Type

Private Type GTList
    GLText  As String
    GLType  As Integer
    GLScrolBar(2) As LScrolBar
    GLControl As String
    
    
    GLHeadColors As Long
    GLHeadColor_Cliks As Long
End Type

'Private Type GTListOnControl
'    IndexPicture  As String
'    bGLType  As GTList
'End Type

Private GLHeadColors As Long
Private GLHeadColorCliks As Long

Private GridList() As GTList
Private CountGridList As Long
Private GridListRange As Integer
Public GLCaption As String
'Private TTypeList As Integer

Private ScrolIndex As Integer
Private SubScrolIndex As Integer
Private PointerList As Integer
Private PointerListHead As Integer
Public OpenThumb As Boolean '==> buat Get Only

Sub Add(Optional ByVal nCount As Long)
    If nCount > 0 Then CountGridList = nCount
    ReDim Preserve GridList(CountGridList)
    CountGridList = CountGridList + 1
End Sub

Sub Clear()
    Erase GridList()
    CountGridList = 0
End Sub

Sub Clear_GLText()
    Erase GridList()
    CountGridList = 0
End Sub

Property Get GLCount() As Long
    GLCount = CountGridList
End Property

Property Get GLHeadColor() As Long
    GLHeadColor = GLHeadColors
End Property
Property Let GLHeadColor(ByVal NewValue As Long)
    GLHeadColors = NewValue
End Property

Property Get GLHeadColorClik() As Long
    GLHeadColorClik = GLHeadColorCliks
End Property
Property Let GLHeadColorClik(ByVal NewValue As Long)
    GLHeadColorCliks = NewValue
End Property

Property Get GLText(ByVal Index As Long) As String
    GLText = GridList(Index).GLText
End Property
Property Let GLText(ByVal Index As Long, ByVal NewValue As String)
    GridList(Index).GLText = NewValue
End Property

Property Get GLControl(ByVal Index As Long) As String
    GLControl = GridList(Index).GLControl
End Property
Property Let GLControl(ByVal Index As Long, ByVal NewValue As String)
    GridList(Index).GLControl = NewValue
End Property

'Property Get GLType(ByVal Index As Long) As Integer
'    GLType = GridList(Index).GLType
'End Property
'Property Let GLType(ByVal Index As Long, ByVal NewValue As Integer)
'    GridList(Index).GLType = NewValue
'End Property

Property Get GLScrolIndex() As Integer
    GLScrolIndex = ScrolIndex
End Property
Property Let GLScrolIndex(ByVal NewValue As Integer)
    ScrolIndex = NewValue
End Property

'Property Get GLSubScrolIndex() As Integer
'    GLSubScrolIndex = SubScrolIndex
'End Property
'Property Let GLSubScrolIndex(ByVal NewValue As Integer)
'    SubScrolIndex = NewValue
'End Property


Property Get GLRange() As Integer
    GLRange = GridListRange
End Property
Property Let GLRange(ByVal NewValue As Integer)
    GridListRange = NewValue
End Property

'Property Get TypeList() As Integer
'    TypeList = TTypeList
'End Property
'Property Let TypeList(ByVal NewValue As Integer)
'    TTypeList = NewValue
'End Property

Property Get GLPointer() As Integer
    GLPointer = PointerList
End Property
Property Let GLPointer(ByVal NewValue As Integer)
    PointerList = NewValue
End Property

'Property Get GLPointerHead() As Integer
'    GLPointerHead = PointerListHead
'End Property
'Property Let GLPointerHead(ByVal NewValue As Integer)
'    PointerListHead = NewValue
'End Property

