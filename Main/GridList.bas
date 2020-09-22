Attribute VB_Name = "GridList"
Type GList
    GLText As String
    
End Type
Type nIcon
    File  As String
    Index As Integer
    X1    As Integer
    Y1    As Integer
End Type
Type PT
    PTVis As Integer
    PTFull As Integer
    PTX1 As Integer
    PTX2 As Integer
    PTY1 As Integer
    PTY2 As Integer
    PTIcon As nIcon
End Type
Public ListPT As PT
Dim tmpListPT As PT

Dim GridList() As GList


Function Get_Format(Names As String, TypeName As String) As String
Dim TInstr As Long, tagTInstr As Long, endTagTI As Long
    
    TInstr = InStr(TInstr + 1, Names, TypeName)
    If TInstr <> 0 Then
        tagTInstr = InStr(TInstr + 1, Names, "|")
        endTagTI = InStr(tagTInstr + 1, Names, "?")
    
        Get_Format = right(left(Names, endTagTI - 1), Abs(endTagTI - tagTInstr) - 1)
        'If Get_Format = "" Then Get_Format = "?"
    Else
        Get_Format = ""
    End If
End Function

Sub LoadFormatPT(nNames As String)
If nNames <> "" Then
    ListPT.PTVis = Val(Get_Format(nNames, "?pt.vis|"))
    ListPT.PTFull = Val(Get_Format(nNames, "?pt.full|"))
    ListPT.PTX1 = Val(Get_Format(nNames, "?pt.x1|"))
    ListPT.PTX2 = Val(Get_Format(nNames, "?pt.x2|"))
    ListPT.PTY1 = Val(Get_Format(nNames, "?pt.y1|"))
    ListPT.PTY2 = Val(Get_Format(nNames, "?pt.y2|"))
    
    ListPT.PTIcon.File = Get_Format(nNames, "?pt.ico.file|")
    ListPT.PTIcon.Index = Val(Get_Format(nNames, "?pt.ico.index|"))
Else
    ListPT = tmpListPT
End If

End Sub
