Attribute VB_Name = "GeneralVariabel"
'Del or Input must not GridXY, this ? _
 <<-081.2948.9130->>
Public Type NGeid 'membuat ini ya..............
    BackColor   As Long
    ForeColor   As Long
    Alignment   As Integer
    Bold        As Boolean
    Italic      As Boolean
    Underline   As Boolean
End Type
 
Public Type Size
    GDRangeX      As Integer
    GDRangeY      As Integer
    SellWidth_Def As Integer
    SellHeight_Def As Integer
    TableWidth    As Single
    TableHeight   As Single
End Type

Public Type SizePicSub
    RangePicSubX1 As Integer
    RangePicSubY1 As Integer
    RangePicSubX2 As Integer
    RangePicSubY2 As Integer
End Type

Public Type SizePic
    SellIconX1 As Integer
    SellIconY1 As Integer
    SellIconX2 As Integer
    SellIconY2 As Integer
    SellIconPicContColms As Integer
    SellIconPicContRows As Integer
End Type
    Public Type LengkapNewgridXY
        GridBackColGra(1) As Long '>>> OLE_COLOR
        GridForeColor    As Long
        'GridYBackColor    As Long
        'GridYForeColor    As Long
    End Type
    
    Public Type NewGridXY '--------------------> New Tool XY|
        GridSize           As Size
        GridSizePic        As SizePic
        GridSizePicSub     As SizePicSub
        GridXCount         As Integer
        GridYCount         As Integer
        GridXYName         As String
        
        GridLenkapX As LengkapNewgridXY
        GridLenkapY As LengkapNewgridXY
        
'        GridBackColGra(1) As OLE_COLOR
'        GridXForeColor   As Long
'        GridYBackColor    As Long
'        GridYForeColor   As Long
        
        GridXYBackColor    As Long
        GridXYBackColorSub As Long
        GridXYForeColor   As Long
        GridXYForeColorSub As Long
        GridXYFillColor  As Long
        GridXYFillColorSub  As Long

        GridType         As Integer
        'GridYBreakPoint  As Integer
        GridFilePicture  As String
        
        GridPassword     As String
        GridAuthor       As String
        GridMemo         As String
        
        GridStyleX       As Integer
        GridStyleY       As Integer
        
        GridTag          As String
    End Type


'Text =MaxLength DefaultValue PasswordChar

            Public Type TypeVariabel '-------------------->?
                VarName As String
                VarLoadType As String '--|
                    VarType() As String             '  |
                    'VarMaxLength As Integer
                    'VarPasswordChar  As String * 1
                    'VarDefaultValue As String
                    'VarAutoNomber As Integer
                    'VarFormula    As String
            End Type
        
        Public Type TypeGridX '--------------------> X
            GridFront        As Boolean
            GridLeft         As Long 'ganti dengan singgel
            GridWidth        As Long 'ganti dengan singgel
            GWidthDefault    As Boolean
            GWSave           As Single
            
            GridValue        As String
            GridStyle        As Integer
            GridColGra(1)    As Long '>>> OLE_COLOR
            
            GridRealPosisi   As Integer '> Untuk Mengetahui Posisi Yang Di Keluarkan
            GridRealOnPosisi As Integer '> Untuk Mengetahui Posisi Yang Default
            GridIndexHead    As Integer
            GridOnIndexHead  As Integer
            
            Tag              As String
            Visibles         As Boolean
            
                        
            'GridPicture      As New StdPicture
            GVariabel        As TypeVariabel
            
            GPicturePut      As Boolean
        End Type
        
        Public Type TypeGridY '--------------------> Y
            GridFront        As Boolean
            GridTop          As Long
            GridHeight       As Long
            GHeightDefault   As Boolean
            GHSave           As Single
            
            GridValue        As String
            GridStyle        As Integer
            GridColGra(1)    As Long '>>> OLE_COLOR
            
            Tag              As String
            Tmp_MultiSelect  As Long
            
            Visibles         As Boolean
            'GridPicture      As New StdPicture
            
            'GridVariabelY    As String

'            GridBreakPoint   As Boolean
'            GVariabel        As TypeVariabel
        End Type
    
                Public Type Lengkap
                    BackColor       As Long
                    ForeColor       As Long
                    FillColor       As Long
                    Alignment       As Integer
                    Bold            As Boolean
                    Italic          As Boolean
                    Underline       As Boolean
                    GColorDefault(1) As Boolean
                End Type
    
            Public Type GSubTypeControl
                TypeControl As Integer
                
                CountContGList As Integer
                ContGList() As New GList
                PointerListHead As Integer
                GLAuto As Boolean
                GLShowCount As Integer
                GLHeight As Integer
                GLSubScrolIndex As Integer
                
                GLFrmtStly As String  'New
'                GLFrmtStly As Integer 'New

'                BackColor  As Long
'                Color_Clik As Long
            End Type
    
    
            Public Type TypeGridXYData '--------------------> XY
                GridXYValue         As String
                GridTag             As String
                GridXYValueSub      As String
                GridTagSub          As String
'fdgfd fd gdfgfd
                GridXYPicIndex      As Integer

                GridXPicIndex       As Integer
                GridYPicIndex       As Integer
                                
                Grid                As Lengkap
                GridSub             As Lengkap
                GridSubType         As GSubTypeControl

'                vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv gggggggggggggg ggggggggg
'                GridXYBackColor     As Long
'                GridXYForeColor     As Long
'                GColorDefault(1)    As Boolean
'                GridAlignment       As Integer
'                GridBold            As Boolean
'                GridItalic          As Boolean
'                GridUnderline       As Boolean
                
'                GridXYBackColorSub  As Long
'                GridXYForeColorSub  As Long
'                GColorDefaultSub(1) As Boolean
'                GridAlignmentSub    As Integer
'                GridBoldSub         As Boolean
'                GridItalicSub       As Boolean
'                GridUnderlineSub    As Boolean

                GridXYFillColor     As Long
                GPasswordChar       As String
                'GridTag          As String
            End Type

Public Type TypeFileGridXY '--------------------> File All XY
    'GridXYName       As String
    'GridXYType       As String
    
        GridNew          As NewGridXY
        'GridCountX       As Integer
        GridX()          As TypeGridX
        'GridCountY       As Integer
        GridY()          As TypeGridY
        GridXYData()     As TypeGridXYData
        
End Type
Public FileAll()     As TypeFileGridXY
Public CountFileAll  As Integer
Public J As TypeFileGridXY


Private Type nIForm
    Index As String
    OnOff As Boolean
End Type
    
Public XForm() As nIForm 'buat menentukan jumlah FileAll ke nForm
Public nXForm As Integer

'Public CBGrid() As New BGrid
'Public nCBGrid  As Integer

'Public nForm() As New Form1
Public cnForm  As Integer


Public nPopPg As Boolean


'Public Sub SaveDataGrid_FileAll(SetNewGrid As NewGridXY, GridX() As TypeGridX, GridY() As TypeGridY, GridXYData() As TypeGridXYData, CountFileAll As Integer)
'ReDim Preserve FileAll(CountFileAll)
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

Function CheckType(ByVal CK As Integer, N As Integer) As Boolean
    'If Tpy = 0 Then Tpy = 1
    'For XY = Form6.Check1.Count - 1 To 0 Step -1
    '    If Tpy Mod (XY + 2) = 0 Then
    '        Tpy = Tpy / (XY + 2)
    '        If iXY = XY Then CheckType = True   'MsgBox "OK"
    '    End If
    'Next XY
    CheckType = False
    
'    For XY = Form6.Check1.Count - 1 To 0 Step -1
'        If CK >= Form6.Check1(XY).Tag Then
'            CK = CK - Form6.Check1(XY).Tag
'            If N = XY Then
'                CheckType = True
'                Exit For
'            End If
'        End If
'    Next XY
End Function


